/**
 * The Microsoft Substrate search client — the backend Teams' own search box
 * uses, and the primary path of `teams_search_messages` (the client-side
 * chat-service scan in `tools/search.ts` is the fallback).
 *
 *   POST https://substrate.office.com/searchservice/api/v2/query
 *        [Authorization: Bearer <token>, Content-Type: application/json]
 *        { entityRequests: [{ entityType: "Message", contentSources: ["Teams"],
 *            query: { queryString, displayQueryString }, from, size,
 *            sort: [{ Field: "Time", SortDirection: "Desc" }] }],
 *          cvid, logicalId, scenario: { Name: "powerbar", … } }
 *   → { EntitySets: [{ ResultSets: [{ Total, MoreResultsAvailable,
 *          Results: [{ HitHighlightedSummary, Source: {
 *            ClientThreadId, ClientConversationId, InternetMessageId,
 *            DateTimeSent, Sender: { EmailAddress: { Name, Address } },
 *            Preview, Subject, ChannelName, TeamName, … } }] }] }] }
 *
 * Verified live 2026-10-12 (token captured off the Teams tab, replayed from
 * Node; read-only queries such as "deploy"; a redacted response is
 * `tests/fixtures/substrate-query-response.json`):
 *   - No header beyond the bearer and the JSON content type is needed. The
 *     token `substrate-auth.ts` captures off `/searchservice/*` worked (an
 *     app-GUID audience, ~1.4h lifetime); so did one captured off another
 *     substrate.office.com path (audience `https://outlook.office.com/`).
 *   - Body: `cvid`/`logicalId` and `scenario.Name` are REQUIRED (400 without);
 *     the sort must be `Field: "Time"` with capitalised keys (`DateTime`, or
 *     `field`/`sortDirection`, is a 400). `size` up to 200 works; `from`
 *     pages without overlap, newest first.
 *   - Response: the Teams conversation is `Source.ClientThreadId`;
 *     `ClientConversationId` is the same plus `;messageid=<root>` for a
 *     channel post or reply. `InternetMessageId` is the Teams message id
 *     (equal to the root id on a root post). `Sender` (= `From`) carries a
 *     display name and an SMTP address — NO MRI, and the address need not be
 *     the token's UPN, so `fromMe` cannot be decided from a hit. `Preview`
 *     (= `HitHighlightedSummary`, no highlight markup) is PLAIN TEXT capped
 *     at ~255 characters; there is no body. `Subject` is set on a channel
 *     root post with a subject. A query with no hits has no `Results` key.
 *     `Total` is an estimate (one phrase query said 86 and paged to 11).
 *   - Matching is Substrate's: words ANDed with stemming ("deploy" matches
 *     "deployed", "deployment"), `"phrases"`, `OR`; not window-bound (all
 *     history unless restricted).
 *   - Server-side restrictions in the query string, ANDed with a
 *     parenthesised user query: `ClientThreadId:"<conversation id>"` (also
 *     for `48:notes`) and `sent>=<ISO date-time, UNQUOTED>` (a quoted value
 *     was mis-read). `from:<word>` exists but matches name/address word
 *     PREFIXES, not the "display name contains" the tool promises, so `from`
 *     stays client-side.
 *   - A bad token is a 401.
 *
 * People lookup (`suggestPeople`, the directory fallback of
 * `teams_add_chat_members` / `teams_remove_chat_members`), verified live
 * 2026-11-09 with the SAME captured token — no new token or bridge scope:
 *
 *   POST https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar
 *        { EntityRequests: [{ EntityType: "People", Query: { QueryString },
 *            Size }], Cvid, LogicalId, Scenario: { Name: "powerbar" } }
 *   → { Groups: [{ Type: "People", Suggestions: [{ DisplayName,
 *          EmailAddresses: string[], UserPrincipalName, JobTitle, Department,
 *          PeopleSubtype, MRI, ExternalDirectoryObjectId, … }] }] }
 *
 *   - A surname, or "First Last", found the person; their
 *     `ExternalDirectoryObjectId` is their Entra oid, and their chat MRI is
 *     `8:orgid:<that oid>`.
 *   - `PeopleSubtype` seen: `OrganizationUser`, `PersonalContact`, `Unknown`,
 *     `PublicDistributionList` — only the first can be a chat member.
 *   - The v2 query endpoint with `EntityType: "People"` is a 400: not used.
 *
 * As with the other clients, the bridge is touched only to capture the token
 * (`substrate-auth.ts`); every request is a plain Node `fetch`.
 */
import { fetchBounded, McpToolError, parseBoolEnv } from '@chrischall/mcp-utils';
import { describeTransportFailure } from './network-error.js';
import { createCapturedTokenManager, type TokenManager } from '@chrischall/mcp-utils/session';
import { createSubstrateTokenCache } from './substrate-token-cache.js';

export const SUBSTRATE_QUERY_URL = 'https://substrate.office.com/searchservice/api/v2/query';
export const SUBSTRATE_SUGGESTIONS_URL = 'https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar';

/** Directory suggestions asked for when the caller names no size. */
const DEFAULT_SUGGESTIONS = 10;

/** Most suggestions one lookup asks for. */
const MAX_SUGGESTIONS = 25;

/**
 * How long after a failed capture (or a 401 on a freshly captured token) an
 * unforced search skips straight to failing
 * (so mode "auto" falls back to the scan at once) instead of holding the call
 * for another capture window. Mode "substrate" forces a fresh attempt.
 */
export const CAPTURE_RETRY_COOLDOWN_MS = 10 * 60 * 1000;

/** Per-request deadline, covering the body read. */
const REQUEST_TIMEOUT_MS = 30_000;

/** A search response is a page of summaries; anything near this is not one. */
const MAX_RESPONSE_BYTES = 8 * 1024 * 1024;

/** Capture lazily, so boot never loads the bridge (see `sharepoint.ts`). */
async function captureSubstrateTokenLazily(): Promise<string> {
  const mod = await import('./substrate-auth.js');
  return mod.captureSubstrateTokenViaFetchproxy();
}

export interface SubstrateSearchRequest {
  /** The user's query, passed through as typed (Substrate's own query syntax: words AND'ed, "phrases", OR). */
  query: string;
  /** Zero-based offset of the page. */
  from: number;
  /** Results per page (verified up to 200). */
  size: number;
  /** Server-side: only this Teams conversation (`19:…`, `48:notes`), via `ClientThreadId:"…"`. */
  conversationId?: string;
  /** Server-side: only messages sent at or after this epoch-ms instant, via `sent>=…` (whole seconds). */
  sentAfter?: number;
}

/** One message hit, normalised from a Substrate result. Every field but the conversation is best-effort. */
export interface SubstrateMessage {
  /** The Teams conversation (`19:…@thread.v2`, `48:notes`, …), without any `;messageid=` suffix. */
  conversationId: string;
  /** The Teams message id (`InternetMessageId`: the epoch-ms id the chat service uses). */
  messageId?: string;
  /** UTC ISO 8601. */
  time?: string;
  sender?: string;
  /** The sender's SMTP address. Substrate carries no MRI, so `fromMe` cannot be decided from a hit. */
  senderEmail?: string;
  /** Substrate's PLAIN-TEXT preview of the message (not HTML; at most ~255 characters, so often cut short). */
  text: string;
  subject?: string;
  /** A channel reply's root post id. */
  replyTo?: string;
}

export interface SubstratePage {
  results: SubstrateMessage[];
  /** Substrate's own estimate of all matches (approximate: it can exceed what paging returns). */
  total?: number;
  moreAvailable: boolean;
  /** Results that carried no Teams conversation id (and so were left out). */
  unmapped: number;
}

/** A person from the directory who can be a chat member. Every field but the MRI and name is best-effort. */
export interface DirectoryPerson {
  /** `8:orgid:<oid>`. */
  mri: string;
  name: string;
  email?: string;
  upn?: string;
  jobTitle?: string;
  department?: string;
}

/**
 * What `teams_search_messages` and the group-chat member tools need of a
 * Substrate client — tests pass a fake. `force` skips the capture cooldown.
 */
export interface SubstrateSearch {
  isConfigured(): boolean;
  searchMessages(req: SubstrateSearchRequest, opts?: { force?: boolean }): Promise<SubstratePage>;
  suggestPeople(text: string, opts?: { size?: number; force?: boolean }): Promise<DirectoryPerson[]>;
}

/** `sent>=` takes an unquoted ISO date-time (a quoted one was mis-read live); whole seconds, rounded down. */
function sentBound(ms: number): string {
  return new Date(Math.floor(ms / 1000) * 1000).toISOString().replace(/\.000Z$/, 'Z');
}

/**
 * The query string: the user's query, ANDed with the verified KQL-style
 * property restrictions. With a restriction, the user's query is
 * parenthesised so an `OR` in it cannot escape the filter (verified live).
 */
export function buildQueryString(req: Pick<SubstrateSearchRequest, 'query' | 'conversationId' | 'sentAfter'>): string {
  const filters: string[] = [];
  if (req.conversationId !== undefined) {
    if (/["\s]/.test(req.conversationId)) {
      throw new McpToolError('That conversation id cannot be expressed as a Substrate filter.', { kind: 'unknown' });
    }
    filters.push(`ClientThreadId:"${req.conversationId}"`);
  }
  if (req.sentAfter !== undefined && Number.isFinite(req.sentAfter)) filters.push(`sent>=${sentBound(req.sentAfter)}`);
  return filters.length === 0 ? req.query : [`(${req.query})`, ...filters].join(' ');
}

/**
 * The query body, as verified live 2026-10-12. Required (400 without them):
 * `cvid` + `logicalId`, and a `scenario` with a `Name`. `sort` must be
 * `{ Field: "Time", SortDirection: "Desc" }` — `DateTime`, or lower-case keys,
 * are a 400.
 */
export function buildQueryBody(req: SubstrateSearchRequest): Record<string, unknown> {
  const cvid = crypto.randomUUID();
  return {
    entityRequests: [
      {
        entityType: 'Message',
        contentSources: ['Teams'],
        query: { queryString: buildQueryString(req), displayQueryString: req.query },
        from: req.from,
        size: req.size,
        sort: [{ Field: 'Time', SortDirection: 'Desc' }],
      },
    ],
    cvid,
    logicalId: cvid,
    scenario: {
      Name: 'powerbar',
      Dimensions: [
        { DimensionName: 'QueryType', DimensionValue: 'Messages' },
        { DimensionName: 'FormFactor', DimensionValue: 'general.web.reactSearch' },
      ],
    },
    QueryAlterationOptions: { EnableSuggestion: false, EnableAlteration: false },
  };
}

type Obj = Record<string, unknown>;

function obj(v: unknown): Obj | undefined {
  return v !== null && typeof v === 'object' && !Array.isArray(v) ? (v as Obj) : undefined;
}

function str(v: unknown): string | undefined {
  return typeof v === 'string' && v.trim() !== '' ? v : undefined;
}

function isoOf(v: unknown): string | undefined {
  const s = str(v);
  if (s === undefined) return undefined;
  const ms = Date.parse(s);
  return Number.isFinite(ms) ? new Date(ms).toISOString() : undefined;
}

function mapResult(raw: unknown): SubstrateMessage | undefined {
  const r = obj(raw);
  const source = obj(r?.Source);
  if (r === undefined || source === undefined) return undefined;

  // `19:x@thread.tacv2;messageid=<root>` for a channel post or reply; the bare id for a chat.
  const clientConv = str(source.ClientConversationId) ?? str(source.ClientThreadId);
  if (clientConv === undefined) return undefined;
  const [conversationId, ...params] = clientConv.split(';');
  if (str(conversationId) === undefined) return undefined;
  const root = params.map((p) => /^messageid=(.+)$/i.exec(p.trim())?.[1]).find((v) => v !== undefined);

  const messageId = str(source.InternetMessageId);
  // `Sender` and `From` were identical on every hit seen live.
  const from = obj(obj(source.Sender)?.EmailAddress) ?? obj(obj(source.From)?.EmailAddress);
  const text = str(source.Preview) ?? str(r.HitHighlightedSummary) ?? '';
  const time = isoOf(source.DateTimeSent) ?? isoOf(source.DateTimeReceived) ?? isoOf(source.DateTimeCreated);
  const sender = str(from?.Name);
  const senderEmail = str(from?.Address);
  const subject = str(source.Subject);

  return {
    conversationId,
    ...(messageId !== undefined ? { messageId } : {}),
    ...(time !== undefined ? { time } : {}),
    ...(sender !== undefined ? { sender } : {}),
    ...(senderEmail !== undefined ? { senderEmail } : {}),
    text,
    ...(subject !== undefined ? { subject } : {}),
    ...(root !== undefined && root !== messageId ? { replyTo: root } : {}),
  };
}

/** Parse a v2 query response into a page of message hits. Never throws on shape. */
export function parseQueryResponse(body: unknown): SubstratePage {
  const sets = obj(body)?.EntitySets;
  const firstSet = Array.isArray(sets) ? obj(sets[0]) : undefined;
  const resultSets = firstSet?.ResultSets;
  const rs = Array.isArray(resultSets) ? obj(resultSets[0]) : undefined;
  if (rs === undefined) return { results: [], moreAvailable: false, unmapped: 0 };

  // A query with no hits has no `Results` key at all (seen live).
  const raw = Array.isArray(rs.Results) ? rs.Results : [];
  const results: SubstrateMessage[] = [];
  let unmapped = 0;
  for (const r of raw) {
    const m = mapResult(r);
    if (m === undefined) unmapped++;
    else results.push(m);
  }
  return {
    results,
    ...(typeof rs.Total === 'number' ? { total: rs.Total } : {}),
    moreAvailable: rs.MoreResultsAvailable === true,
    unmapped,
  };
}

/** The suggestions body, as verified live 2026-11-09 (capitalised keys, as Teams sends them). */
export function buildSuggestionsBody(text: string, size: number): Record<string, unknown> {
  return {
    EntityRequests: [{ EntityType: 'People', Query: { QueryString: text }, Size: size }],
    Cvid: crypto.randomUUID(),
    LogicalId: crypto.randomUUID(),
    Scenario: { Name: 'powerbar' },
  };
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const ORGID_MRI = /^8:orgid:\S+$/i;

/**
 * A suggestion as a directory person, or `undefined` when it cannot be a chat
 * member: only an `OrganizationUser` (or an entry with no subtype) with a
 * usable id — its `MRI` when that is `8:orgid:…`, else `8:orgid:<its
 * ExternalDirectoryObjectId, lower-cased>`.
 */
function mapSuggestion(raw: unknown): DirectoryPerson | undefined {
  const r = obj(raw);
  if (r === undefined) return undefined;
  const subtype = str(r.PeopleSubtype);
  if (subtype !== undefined && subtype !== 'OrganizationUser') return undefined;
  const mriField = str(r.MRI)?.trim();
  const oid = str(r.ExternalDirectoryObjectId)?.trim();
  const mri =
    mriField !== undefined && ORGID_MRI.test(mriField)
      ? mriField
      : oid !== undefined && GUID.test(oid)
        ? `8:orgid:${oid.toLowerCase()}`
        : undefined;
  if (mri === undefined) return undefined;
  const email = Array.isArray(r.EmailAddresses) ? r.EmailAddresses.map(str).find((e) => e !== undefined) : undefined;
  const upn = str(r.UserPrincipalName);
  const name = str(r.DisplayName)?.trim() ?? email ?? upn ?? mri;
  const jobTitle = str(r.JobTitle);
  const department = str(r.Department);
  return {
    mri,
    name,
    ...(email !== undefined ? { email } : {}),
    ...(upn !== undefined ? { upn } : {}),
    ...(jobTitle !== undefined ? { jobTitle } : {}),
    ...(department !== undefined ? { department } : {}),
  };
}

/** Parse a suggestions response into the people who can be chat members, one per MRI. Never throws on shape. */
export function parseSuggestionsResponse(body: unknown): DirectoryPerson[] {
  const groups = obj(body)?.Groups;
  if (!Array.isArray(groups)) return [];
  const byMri = new Map<string, DirectoryPerson>();
  for (const g of groups) {
    const group = obj(g);
    if (group?.Type !== 'People' || !Array.isArray(group.Suggestions)) continue;
    for (const raw of group.Suggestions) {
      const p = mapSuggestion(raw);
      if (p !== undefined && !byMri.has(p.mri.toLowerCase())) byMri.set(p.mri.toLowerCase(), p);
    }
  }
  return [...byMri.values()];
}

/** What a request is for: words its errors, and names its host on a transport failure. */
interface Lookup {
  url: string;
  /** "Substrate search" / "Substrate people lookup". */
  service: string;
  /** "searches" / "people lookups", for the 429. */
  plural: string;
}

const SEARCH: Lookup = { url: SUBSTRATE_QUERY_URL, service: 'Substrate search', plural: 'searches' };
const PEOPLE: Lookup = { url: SUBSTRATE_SUGGESTIONS_URL, service: 'Substrate people lookup', plural: 'people lookups' };

export interface SubstrateClientOptions {
  env?: NodeJS.ProcessEnv;
  /** Injected for tests so no suite ever touches the real bridge. */
  captureToken?: () => Promise<string>;
  fetchImpl?: typeof fetch;
}

export class SubstrateClient implements SubstrateSearch {
  readonly #fetchImpl: typeof fetch;
  readonly #configError: McpToolError | null = null;
  readonly #manager: TokenManager;
  /**
   * When the last capture failed, or a freshly captured token was still
   * refused (401), and how; cleared by a successful capture.
   */
  #captureFailure: { at: number; error: unknown } | null = null;

  constructor(opts: SubstrateClientOptions = {}) {
    const env = opts.env ?? process.env;
    const capture = opts.captureToken ?? captureSubstrateTokenLazily;
    // A receiver-safe wrapper, never the bare global (see `sharepoint.ts`).
    this.#fetchImpl = opts.fetchImpl ?? ((input, init) => fetch(input, init));

    const recapture = async (): Promise<string> => {
      try {
        const token = await capture();
        this.#captureFailure = null;
        return token;
      } catch (error) {
        this.#captureFailure = { at: Date.now(), error };
        throw error;
      }
    };
    this.#manager = createCapturedTokenManager({
      capture: recapture,
      persistence: createSubstrateTokenCache(env),
    });

    // Like SharePoint, there is no direct-token env var: the bridge is the only source.
    if (parseBoolEnv('TEAMS_DISABLE_FETCHPROXY', { env })) {
      this.#configError = new McpToolError('Substrate search is disabled (no token source).', {
        kind: 'no_credential',
        hint: 'Unset TEAMS_DISABLE_FETCHPROXY so a Substrate token can be captured from a signed-in Teams tab.',
      });
    }
  }

  /** Whether a token could be obtained at all (the bridge is not disabled). */
  isConfigured(): boolean {
    return this.#configError === null;
  }

  async searchMessages(req: SubstrateSearchRequest, opts: { force?: boolean } = {}): Promise<SubstratePage> {
    return parseQueryResponse(await this.#post(SEARCH, buildQueryBody(req), opts.force === true));
  }

  /**
   * People in the directory matching `text` (a name, part of one, or an
   * address), as Teams' search box suggests them — only those who can be a
   * chat member (see {@link parseSuggestionsResponse}). Blank text asks nothing.
   */
  async suggestPeople(text: string, opts: { size?: number; force?: boolean } = {}): Promise<DirectoryPerson[]> {
    const q = text.trim();
    if (q === '') return [];
    const size = Math.min(Math.max(1, Math.floor(opts.size ?? DEFAULT_SUGGESTIONS)), MAX_SUGGESTIONS);
    return parseSuggestionsResponse(await this.#post(PEOPLE, buildSuggestionsBody(q, size), opts.force === true));
  }

  /**
   * POST `body` to Substrate with the captured token and return the parsed
   * JSON: the config error, the capture cooldown (unless `force`), one
   * re-capture on a 401, a bounded read, no redirects, and the status mapped
   * to an `McpToolError` — the body is never echoed.
   */
  async #post(lookup: Lookup, body: Record<string, unknown>, force: boolean): Promise<unknown> {
    if (this.#configError) throw this.#configError;
    const failed = this.#captureFailure;
    if (!force && failed !== null && Date.now() - failed.at < CAPTURE_RETRY_COOLDOWN_MS) throw failed.error;

    let text = '';
    const res = await this.#manager.withAuth(async (token) => {
      let r: Awaited<ReturnType<typeof fetchBounded<'text'>>>;
      try {
        r = await fetchBounded(
          lookup.url,
          {
            method: 'POST',
            headers: {
              authorization: `Bearer ${token}`,
              'content-type': 'application/json',
              accept: 'application/json',
            },
            body: JSON.stringify(body),
            redirect: 'error',
          },
          {
            read: 'text',
            maxBytes: MAX_RESPONSE_BYTES,
            service: 'Substrate',
            timeoutMs: REQUEST_TIMEOUT_MS,
            fetchImpl: this.#fetchImpl,
          },
        );
      } catch (e) {
        throw describeTransportFailure(e, lookup.url, lookup.service) ?? e;
      }
      text = r.body;
      return r.res;
    });

    try {
      checkStatus(res.status, lookup);
    } catch (error) {
      // A 401 here is on a token `withAuth` just re-captured: the capture itself
      // is not yielding a usable token. Cool down like a failed capture, so each
      // unforced call does not hold the caller for another capture window.
      if (res.status === 401) this.#captureFailure = { at: Date.now(), error };
      throw error;
    }
    try {
      return JSON.parse(text);
    } catch {
      // Never echo the body: it holds message content or people's details.
      throw new McpToolError('Substrate returned a response that is not JSON.', { status: res.status, kind: 'http' });
    }
  }
}

/** Map a non-2xx to an `McpToolError` carrying its status. The body is never included. */
function checkStatus(status: number, lookup: Lookup): void {
  if (status >= 200 && status < 300) return;
  if (status === 401) {
    throw new McpToolError('Substrate rejected the search token (401) even after a fresh capture.', {
      status,
      kind: 'credential_rejected',
      hint: 'Run a search in a signed-in Teams tab to mint a fresh Substrate token, then retry.',
    });
  }
  if (status === 429) {
    throw new McpToolError(`Substrate is rate-limiting ${lookup.plural} (429).`, {
      status,
      kind: 'http',
      hint: 'Wait a minute before retrying.',
    });
  }
  throw new McpToolError(`${lookup.service} returned HTTP ${status}.`, {
    status,
    kind: 'http',
    ...(status >= 500 ? { hint: 'A Substrate-side failure; retry shortly.' } : {}),
  });
}
