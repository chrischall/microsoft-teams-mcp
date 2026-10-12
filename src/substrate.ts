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
 * As with the other clients, the bridge is touched only to capture the token
 * (`substrate-auth.ts`); every request is a plain Node `fetch`.
 */
import { decodeJwtExp, fetchBounded, McpToolError, parseBoolEnv } from '@chrischall/mcp-utils';
import { describeTransportFailure } from './network-error.js';
import { TokenManager, type BearerTokens } from '@chrischall/mcp-utils/session';
import { createSubstrateTokenCache } from './substrate-token-cache.js';

export const SUBSTRATE_QUERY_URL = 'https://substrate.office.com/searchservice/api/v2/query';

/**
 * How long after a failed capture (or a 401 on a freshly captured token) an
 * unforced search skips straight to failing
 * (so mode "auto" falls back to the scan at once) instead of holding the call
 * for another capture window. Mode "substrate" forces a fresh attempt.
 */
export const CAPTURE_RETRY_COOLDOWN_MS = 10 * 60 * 1000;

/** Fallback lifetime for a token whose `exp` we cannot read (as in `sharepoint.ts`). */
const UNKNOWN_TOKEN_TTL_MS = 10 * 60 * 1000;

/** Stands in for a refresh token this flow does not have (see `sharepoint.ts`). */
const RECAPTURE_SENTINEL = 'fetchproxy:recapture';

/** Per-request deadline, covering the body read. */
const REQUEST_TIMEOUT_MS = 30_000;

/** A search response is a page of summaries; anything near this is not one. */
const MAX_RESPONSE_BYTES = 8 * 1024 * 1024;

function toBearerTokens(accessToken: string): BearerTokens {
  let expiresAt: number;
  try {
    expiresAt = decodeJwtExp(accessToken) * 1000;
  } catch {
    expiresAt = Date.now() + UNKNOWN_TOKEN_TTL_MS;
  }
  return { accessToken, refreshToken: RECAPTURE_SENTINEL, expiresAt };
}

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

/** What `teams_search_messages` needs of a Substrate client — tests pass a fake. */
export interface SubstrateSearch {
  isConfigured(): boolean;
  searchMessages(req: SubstrateSearchRequest, opts?: { force?: boolean }): Promise<SubstratePage>;
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

    const recapture = async (): Promise<BearerTokens> => {
      try {
        const tokens = toBearerTokens(await capture());
        this.#captureFailure = null;
        return tokens;
      } catch (error) {
        this.#captureFailure = { at: Date.now(), error };
        throw error;
      }
    };
    const persistence = createSubstrateTokenCache(env);
    this.#manager = new TokenManager({
      initial: recapture,
      refresh: recapture,
      ...(persistence ? { persistence } : {}),
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
    if (this.#configError) throw this.#configError;
    const failed = this.#captureFailure;
    if (!opts.force && failed !== null && Date.now() - failed.at < CAPTURE_RETRY_COOLDOWN_MS) throw failed.error;

    let text = '';
    const res = await this.#manager.withAuth(async (token) => {
      let r: Awaited<ReturnType<typeof fetchBounded<'text'>>>;
      try {
        r = await fetchBounded(
          SUBSTRATE_QUERY_URL,
          {
            method: 'POST',
            headers: {
              authorization: `Bearer ${token}`,
              'content-type': 'application/json',
              accept: 'application/json',
            },
            body: JSON.stringify(buildQueryBody(req)),
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
        throw describeTransportFailure(e, SUBSTRATE_QUERY_URL, 'Substrate search') ?? e;
      }
      text = r.body;
      return r.res;
    });

    try {
      checkStatus(res.status);
    } catch (error) {
      // A 401 here is on a token `withAuth` just re-captured: the capture itself
      // is not yielding a usable token. Cool down like a failed capture, so each
      // unforced search does not hold the call for another capture window.
      if (res.status === 401) this.#captureFailure = { at: Date.now(), error };
      throw error;
    }
    try {
      return parseQueryResponse(JSON.parse(text));
    } catch {
      // Never echo the body: it holds message content.
      throw new McpToolError('Substrate returned a response that is not JSON.', { status: res.status, kind: 'http' });
    }
  }
}

/** Map a non-2xx to an `McpToolError` carrying its status. The body is never included. */
function checkStatus(status: number): void {
  if (status >= 200 && status < 300) return;
  if (status === 401) {
    throw new McpToolError('Substrate rejected the search token (401) even after a fresh capture.', {
      status,
      kind: 'credential_rejected',
      hint: 'Run a search in a signed-in Teams tab to mint a fresh Substrate token, then retry.',
    });
  }
  if (status === 429) {
    throw new McpToolError('Substrate is rate-limiting searches (429).', { status, kind: 'http', hint: 'Wait a minute before retrying.' });
  }
  throw new McpToolError(`Substrate search returned HTTP ${status}.`, {
    status,
    kind: 'http',
    ...(status >= 500 ? { hint: 'A Substrate-side failure; retry shortly.' } : {}),
  });
}
