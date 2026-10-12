/**
 * The Microsoft Substrate search client — the backend Teams' own search box
 * uses, and the primary path of `teams_search_messages` (the client-side
 * chat-service scan in `tools/search.ts` is the fallback).
 *
 *   POST https://substrate.office.com/searchservice/api/v2/query
 *        [Authorization: Bearer <Substrate token>]
 *        { entityRequests: [{ entityType: "Message", contentSources: ["Teams"],
 *          query: { queryString }, from, size, sort: [DateTime desc] }], … }
 *   → { EntitySets: [{ ResultSets: [{ Total, MoreResultsAvailable,
 *          Results: [{ HitHighlightedSummary, Source: { … } }] }] }] }
 *
 * LIVE STATUS, 2026-10-12 — only partly verified:
 *   - VERIFIED: the Teams web tab's search box issues exactly
 *     `POST https://substrate.office.com/searchservice/api/v2/query` (seen
 *     twice, status 200 on the earlier probe; no query string).
 *   - NOT VERIFIED: everything past the URL. The request is made from a worker,
 *     so its body and response were not observable from the page, and the
 *     token could not be captured to replay it: the capture never reached the
 *     extension, because the fetchproxy server refused the answering
 *     extension — its identity did not match the one `microsoft-teams-mcp` had
 *     pinned, and re-pinning (`FETCHPROXY_TRUST_NEW_EXTENSION=1`) is the
 *     user's decision, not this code's. So the request body, the required
 *     headers (if any beyond the bearer: `x-client-flights`, `Prefer`, …),
 *     the response field names below and the paging behaviour are the best
 *     public evidence of Substrate's v2 query API, not observed fact. The
 *     parser is deliberately defensive (every field optional, a hit without a
 *     Teams conversation id is counted as `unmapped` rather than guessed), and
 *     `teams_search_messages` in mode "auto" falls back to the scan on any
 *     failure here, so a wrong guess degrades to the old behaviour.
 *
 * As with the other clients, the bridge is touched only to capture the token
 * (`substrate-auth.ts`); every request is a plain Node `fetch`.
 */
import { decodeJwtExp, fetchBounded, McpToolError, parseBoolEnv } from '@chrischall/mcp-utils';
import { TokenManager, type BearerTokens } from '@chrischall/mcp-utils/session';
import { createSubstrateTokenCache } from './substrate-token-cache.js';

export const SUBSTRATE_QUERY_URL = 'https://substrate.office.com/searchservice/api/v2/query';

/**
 * How long after a failed capture an unforced search skips straight to failing
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
  /** Passed to Substrate as the query string, as typed. */
  query: string;
  /** Zero-based offset of the page. */
  from: number;
  /** Results per page. */
  size: number;
}

/** One message hit, normalised from a Substrate result. Every field but the conversation is best-effort. */
export interface SubstrateMessage {
  /** The Teams conversation (`19:…@thread.v2`, `48:notes`, …), without any `;messageid=` suffix. */
  conversationId: string;
  messageId?: string;
  /** UTC ISO 8601. */
  time?: string;
  sender?: string;
  senderEmail?: string;
  /** The sender's MRI (`8:orgid:<oid>`), when the hit carries one. */
  senderMri?: string;
  /** The message body (HTML), or the best excerpt of it Substrate returned. */
  html: string;
  /** True when `html` is a preview or highlighted summary rather than the whole body. */
  excerpt?: true;
  subject?: string;
  /** A channel reply's root post id. */
  replyTo?: string;
}

export interface SubstratePage {
  results: SubstrateMessage[];
  /** Substrate's own estimate of all matches, when it gives one. */
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

/** The query body. See the module note: best evidence, not observed live. */
export function buildQueryBody(req: SubstrateSearchRequest): Record<string, unknown> {
  const cvid = crypto.randomUUID();
  return {
    entityRequests: [
      {
        entityType: 'Message',
        contentSources: ['Teams'],
        propertySet: 'Optimized',
        query: { queryString: req.query, displayQueryString: req.query },
        from: req.from,
        size: req.size,
        sort: [{ field: 'DateTime', sortDirection: 'Desc' }],
        topResultsCount: 0,
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

/** The first non-empty string among an extensions bag's keys ending in `suffix`. */
function extension(source: Obj, suffix: string): string | undefined {
  const ext = obj(source.Extensions);
  if (ext === undefined) return undefined;
  for (const [k, v] of Object.entries(ext)) if (k.endsWith(suffix) && str(v) !== undefined) return v as string;
  return undefined;
}

function mapResult(raw: unknown): SubstrateMessage | undefined {
  const r = obj(raw);
  const source = obj(r?.Source);
  if (r === undefined || source === undefined) return undefined;

  // `19:x@thread.tacv2;messageid=<root>` for a channel thread; a bare id for a chat.
  const clientConv = str(source.ClientConversationId) ?? str(source.ClientThreadId);
  if (clientConv === undefined) return undefined;
  const [conversationId, ...params] = clientConv.split(';');
  if (str(conversationId) === undefined) return undefined;
  const root = params.map((p) => /^messageid=(.+)$/i.exec(p.trim())?.[1]).find((v) => v !== undefined);

  const messageId = str(source.ClientMessageId) ?? extension(source, '_SkypeMessageId');
  const from = obj(obj(source.From)?.EmailAddress) ?? obj(obj(source.Sender)?.EmailAddress);
  const body = str(obj(source.Body)?.Content);
  const preview = str(source.Preview);
  const summary = str(r.HitHighlightedSummary) ?? str(r.Summary);
  const html = body ?? preview ?? summary ?? '';
  const time = isoOf(source.DateTimeSent) ?? isoOf(source.DateTimeReceived) ?? isoOf(source.DateTimeCreated);
  const sender = str(from?.Name) ?? str(source.FromName);
  const senderEmail = str(from?.Address);
  const senderMri = extension(source, '_FromSkypeInternalId');
  const subject = str(source.Subject);

  return {
    conversationId,
    ...(messageId !== undefined ? { messageId } : {}),
    ...(time !== undefined ? { time } : {}),
    ...(sender !== undefined ? { sender } : {}),
    ...(senderEmail !== undefined ? { senderEmail } : {}),
    ...(senderMri !== undefined ? { senderMri } : {}),
    html,
    ...(body === undefined ? { excerpt: true as const } : {}),
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
  /** When the last capture failed, and how; cleared by a successful one. */
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
      const r = await fetchBounded(
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
      text = r.body;
      return r.res;
    });

    checkStatus(res.status);
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
