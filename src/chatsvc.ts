/**
 * The Teams chat-service REST client.
 *
 * Auth is a bearer token whose only source is the user's browser (see
 * `auth-fetchproxy.ts`), or a `TEAMS_ACCESS_TOKEN` supplied directly. Once
 * held, every request is a plain Node `fetch` — the bridge is never on the hot
 * path.
 *
 * Read-only by construction: the only methods are `listConversations`,
 * `getMessages` and `fetchImage`. Nothing here writes the read marker
 * (`consumptionhorizon`), so reading a chat through this MCP never clears the
 * user's unread badges.
 *
 * Verified live 2026-10-10 against a real Teams tenant — see the "Verified
 * live facts" section of `docs/plans/2026-10-10-chatsvc-api.md` for the
 * response shapes the types below describe.
 */
import {
  decodeJwtClaim,
  decodeJwtExp,
  fetchBounded,
  McpToolError,
  parseBoolEnv,
  readEnvVar,
} from '@chrischall/mcp-utils';
import { TokenManager, type BearerTokens } from '@chrischall/mcp-utils/session';
import { createTokenCache, reportCacheWriteFailure } from './token-cache.js';

/**
 * Load the bridge bootstrap lazily.
 *
 * `auth-fetchproxy.ts` pulls in `@chrischall/mcp-utils/fetchproxy`, which
 * imports the OPTIONAL `@fetchproxy/server` peer. The `.mcpb` bundle ships no
 * `node_modules`, so a top-level import here would throw
 * `ERR_MODULE_NOT_FOUND` before the server could answer `initialize`. Deferring
 * it means a user who supplies `TEAMS_ACCESS_TOKEN` never needs it installed.
 */
async function captureTokenLazily(): Promise<string> {
  const mod = await import('./auth-fetchproxy.js');
  return mod.captureTokenViaFetchproxy();
}

/** The chat-service region this tenant was verified on. */
export const DEFAULT_REGION = 'amer';

/**
 * The chat-service base URL: `TEAMS_CHATSVC_BASE` verbatim when set, else the
 * New Teams proxy path for `TEAMS_CHATSVC_REGION` (default `amer`).
 */
export function chatServiceBase(env: NodeJS.ProcessEnv = process.env): string {
  const override = readEnvVar('TEAMS_CHATSVC_BASE', { env });
  if (override !== undefined) return override.replace(/\/+$/, '');
  const region = readEnvVar('TEAMS_CHATSVC_REGION', { env }) ?? DEFAULT_REGION;
  return `https://teams.cloud.microsoft/api/chatsvc/${encodeURIComponent(region)}/v1`;
}

/**
 * Fallback lifetime for a token whose `exp` we cannot read.
 *
 * Real tokens are JWTs carrying `exp` (~23.5h out), so this is only reached
 * for a hand-supplied opaque `TEAMS_ACCESS_TOKEN`. Deliberately short: being
 * wrong in this direction costs one extra 401-and-replay.
 */
const UNKNOWN_TOKEN_TTL_MS = 10 * 60 * 1000;

/**
 * Sentinel standing in for a refresh token this flow does not have.
 *
 * The token comes from MSAL's hidden-iframe flow; there is no refresh grant we
 * can drive, so "refreshing" means capturing a new one. `TokenManager` only
 * calls `refresh` when a refresh token is present, hence the sentinel.
 */
const RECAPTURE_SENTINEL = 'fetchproxy:recapture';

/** Per-request deadline, covering the body read. */
const REQUEST_TIMEOUT_MS = 60_000;

/** Paging caps — a runaway `backwardLink` chain must not hold a tool call open. */
const MAX_CONVERSATION_PAGES = 10;
const MAX_MESSAGE_PAGES = 20;

/** Largest inline image we will buffer (they measured 8–18 KB live). */
const MAX_IMAGE_BYTES = 5 * 1024 * 1024;

/**
 * Inline images live on Skype's AMS object store (`us-api.asm.skype.com`, the
 * host varies by region). The bearer is sent ONLY to hosts under this suffix.
 */
const IMAGE_HOST_SUFFIX = '.asm.skype.com';

function toBearerTokens(accessToken: string): BearerTokens {
  let expiresAt: number;
  try {
    // decodeJwtExp returns seconds; a non-JWT throws.
    expiresAt = decodeJwtExp(accessToken) * 1000;
  } catch {
    expiresAt = Date.now() + UNKNOWN_TOKEN_TTL_MS;
  }
  return { accessToken, refreshToken: RECAPTURE_SENTINEL, expiresAt };
}

/** The last message a conversation listing carries. Every field is a string from the wire. */
export interface RawLastMessage {
  id?: string;
  composetime?: string;
  originalarrivaltime?: string;
  content?: string;
  /** A URL ending in the sender's MRI. */
  from?: string;
  imdisplayname?: string;
  messagetype?: string;
  conversationid?: string;
  sequenceId?: number;
}

/** One entry of `GET /users/ME/conversations`. */
export interface RawConversation {
  /** `…@thread.v2` group/meeting, `…@thread.tacv2` channel, `…@unq.gbl.spaces` 1:1. */
  id: string;
  type?: string;
  version?: number;
  targetLink?: string;
  lastMessage?: RawLastMessage;
  memberProperties?: Record<string, unknown>;
  /** `consumptionhorizon` is `"<lastReadArrivalMs>;<clientMs>;<messageId>"`. */
  properties?: Record<string, string | undefined>;
  threadProperties?: {
    topic?: string;
    threadType?: string;
    productThreadType?: string;
    spaceId?: string;
    groupId?: string;
    createdat?: string;
    lastjoinat?: string;
    privacy?: string;
    [key: string]: string | undefined;
  };
}

/** One entry of `GET /users/ME/conversations/{id}/messages`. */
export interface RawMessage {
  /** Numeric string: the arrival time in epoch ms. */
  id: string;
  clientmessageid?: string;
  composetime?: string;
  originalarrivaltime?: string;
  content?: string;
  contenttype?: string;
  /** `Text` / `RichText/Html` are user messages; anything else is a system event. */
  messagetype?: string;
  from?: string;
  imdisplayname?: string;
  /** In a channel, a reply's link ends `;messageid=<rootId>`. */
  conversationLink?: string;
  conversationid?: string;
  /** `files` is a JSON STRING; `deletetime` set means deleted. */
  properties?: Record<string, unknown>;
  sequenceId?: number;
  version?: string;
  type?: string;
}

interface Page {
  _metadata?: { backwardLink?: string };
}
interface ConversationsPage extends Page {
  conversations?: RawConversation[];
}
interface MessagesPage extends Page {
  messages?: RawMessage[];
}

export interface GetMessagesOptions {
  /**
   * Stop paging at the first message (newest-first) this accepts. That message
   * is NOT included — it is the boundary (the read horizon, a `since` cutoff).
   */
  stopAt?: (m: RawMessage) => boolean;
  /** Most messages to return. */
  max: number;
}

export interface ChatServiceClientOptions {
  env?: NodeJS.ProcessEnv;
  /** Injected for tests so no suite ever touches the real bridge. */
  captureToken?: () => Promise<string>;
  fetchImpl?: typeof fetch;
}

export class ChatServiceClient {
  readonly #baseUrl: string;
  /**
   * Deferred configuration error. The constructor never throws, so the server
   * still boots (and answers `tools/list`) with no credential present.
   */
  readonly #configError: McpToolError | null = null;
  readonly #env: NodeJS.ProcessEnv;
  readonly #capture: () => Promise<string>;
  readonly #fetchImpl: typeof fetch;
  readonly #direct: string | undefined;
  #tokens: TokenManager | null = null;

  constructor(opts: ChatServiceClientOptions = {}) {
    this.#env = opts.env ?? process.env;
    this.#capture = opts.captureToken ?? captureTokenLazily;
    // A receiver-safe wrapper, never the bare global: older undici throws
    // `Illegal invocation` when `fetch` is called with a non-global receiver.
    this.#fetchImpl = opts.fetchImpl ?? ((input, init) => fetch(input, init));
    this.#baseUrl = chatServiceBase(this.#env);
    this.#direct = readEnvVar('TEAMS_ACCESS_TOKEN', { env: this.#env });

    const bridgeOff = parseBoolEnv('TEAMS_DISABLE_FETCHPROXY', { env: this.#env });
    if (this.#direct === undefined && bridgeOff) {
      this.#configError = new McpToolError('No Teams credential is configured.', {
        kind: 'no_credential',
        hint:
          'Set TEAMS_ACCESS_TOKEN, or unset TEAMS_DISABLE_FETCHPROXY so the token can ' +
          'be captured from a signed-in Teams browser tab.',
      });
    }
  }

  /** Non-secret description of where the credential comes from. */
  get tokenSource(): string {
    return this.#direct !== undefined ? 'TEAMS_ACCESS_TOKEN' : 'browser capture';
  }

  /** The API base this client talks to. Non-secret; safe for healthcheck output. */
  get apiBase(): string {
    return this.#baseUrl;
  }

  /** Epoch ms at which the current token expires, or `null` if none is held. */
  tokenExpiresAt(): number | null {
    return this.#tokens ? this.#tokens.getExpiresAt() : null;
  }

  /**
   * The signed-in user's MRI, `8:orgid:<oid>`, from the token's `oid` claim.
   * Needed to tell the user's own messages from everyone else's.
   */
  async meMri(): Promise<string> {
    this.#requireConfigured();
    const token = await this.#tokenManager().getAccessToken();
    let oid: unknown;
    try {
      oid = decodeJwtClaim(token, 'oid');
    } catch {
      oid = undefined;
    }
    if (typeof oid !== 'string' || oid === '') {
      throw new McpToolError('The Teams token does not identify the signed-in user (no oid claim).', {
        hint:
          'Capture a fresh token from a signed-in teams.cloud.microsoft tab, or supply a ' +
          'TEAMS_ACCESS_TOKEN that is the chat-service JWT itself.',
      });
    }
    return `8:orgid:${oid}`;
  }

  /**
   * Every conversation the user is in, across up to 10 pages, minus `48:*`
   * system threads (notes, calendar, …) which are not chats.
   */
  async listConversations(): Promise<RawConversation[]> {
    this.#requireConfigured();
    const out: RawConversation[] = [];
    let url: string | undefined = `${this.#baseUrl}/users/ME/conversations?view=msnp24Equivalent&pageSize=200`;
    for (let page = 0; url !== undefined && page < MAX_CONVERSATION_PAGES; page++) {
      const body: ConversationsPage = await this.#getJson<ConversationsPage>(url);
      const convs = body.conversations ?? [];
      if (convs.length === 0) break;
      for (const c of convs) {
        if (typeof c?.id === 'string' && !c.id.startsWith('48:')) out.push(c);
      }
      url = this.#nextLink(body);
    }
    return out;
  }

  /**
   * A conversation's messages, newest first, paging back until `stopAt`
   * accepts a message, `max` are collected, the pages run out, or 20 pages
   * have been read.
   */
  async getMessages(id: string, opts: GetMessagesOptions): Promise<RawMessage[]> {
    this.#requireConfigured();
    const out: RawMessage[] = [];
    let url: string | undefined =
      `${this.#baseUrl}/users/ME/conversations/${encodeURIComponent(id)}/messages` +
      '?view=msnp24Equivalent|supportsMessageProperties&pageSize=200&startTime=1';
    for (let page = 0; url !== undefined && page < MAX_MESSAGE_PAGES; page++) {
      const body: MessagesPage = await this.#getJson<MessagesPage>(url);
      const msgs = body.messages ?? [];
      if (msgs.length === 0) break;
      for (const m of msgs) {
        if (opts.stopAt?.(m)) return out;
        out.push(m);
        if (out.length >= opts.max) return out;
      }
      url = this.#nextLink(body);
    }
    return out;
  }

  /**
   * Download an inline image (`<img src>` in a message body).
   *
   * The URL comes out of message content — i.e. it is attacker-influenced —
   * so it is checked BEFORE a token is even resolved: https only, and a host
   * under `.asm.skype.com`. Anything else would hand the user's bearer to
   * whoever wrote the message.
   */
  async fetchImage(url: string): Promise<{ mimeType: string; data: Buffer }> {
    this.#requireConfigured();
    let target: URL;
    try {
      target = new URL(url);
    } catch {
      throw new McpToolError('Refusing to fetch an image link that is not a URL.', {
        hint: 'Pass an image ref exactly as a message listing returned it.',
      });
    }
    if (target.protocol !== 'https:' || !target.hostname.endsWith(IMAGE_HOST_SUFFIX)) {
      throw new McpToolError(`Refusing to fetch an image from ${target.protocol}//${target.hostname}.`, {
        hint: `Only https images on *${IMAGE_HOST_SUFFIX} (Teams' own image store) are fetched.`,
      });
    }
    const { status, headers, body } = await this.#send(target.toString(), 'bytes', MAX_IMAGE_BYTES);
    this.#checkStatus(status);
    const mimeType = (headers.get('content-type') ?? 'application/octet-stream').split(';')[0].trim();
    return { mimeType, data: Buffer.from(body as Uint8Array) };
  }

  #requireConfigured(): void {
    if (this.#configError) throw this.#configError;
  }

  #tokenManager(): TokenManager {
    if (this.#tokens) return this.#tokens;

    const direct = this.#direct;
    const persistence = direct === undefined ? createTokenCache(this.#env) : null;
    const mint = async (): Promise<BearerTokens> => toBearerTokens(direct ?? (await this.#capture()));

    this.#tokens = new TokenManager({
      initial: mint,
      // There is no refresh grant — re-capture instead. But a token supplied as
      // TEAMS_ACCESS_TOKEN must NOT be "refreshed" from the browser: that would
      // silently swap the operator's chosen credential for whichever account a
      // tab is signed into.
      refresh: async () => {
        if (direct !== undefined) throw this.#directRejected();
        return toBearerTokens(await this.#capture());
      },
      // A failed cache write must not fail the call that triggered it.
      ...(persistence
        ? {
            persistence: {
              load: () => persistence.load(),
              save: (t: BearerTokens) => {
                try {
                  persistence.save(t);
                } catch (e) {
                  reportCacheWriteFailure(e);
                }
              },
              clear: () => {
                try {
                  persistence.clear();
                } catch {
                  /* nothing useful to do; the next capture overwrites it */
                }
              },
            },
          }
        : {}),
    });
    return this.#tokens;
  }

  #directRejected(): McpToolError {
    return new McpToolError('TEAMS_ACCESS_TOKEN was rejected by Teams (401).', {
      status: 401,
      kind: 'credential_rejected',
      hint:
        'A directly-supplied token cannot be refreshed automatically. Replace ' +
        'TEAMS_ACCESS_TOKEN with a current one, or unset it to capture from a signed-in ' +
        'Teams browser tab instead.',
    });
  }

  /**
   * One authenticated GET. `TokenManager.withAuth` supplies the token and, on a
   * 401, re-captures once and replays once — so a 401 that reaches the caller
   * is the second one.
   *
   * The body is read as text/bytes rather than JSON so a 401's non-JSON body
   * cannot throw inside the call and pre-empt that replay.
   */
  async #send(
    url: string,
    read: 'text' | 'bytes',
    maxBytes?: number,
  ): Promise<{ status: number; headers: Headers; body: string | Uint8Array }> {
    let body: string | Uint8Array = '';
    const res = await this.#tokenManager().withAuth(async (token) => {
      const r = await fetchBounded(
        url,
        {
          method: 'GET',
          headers: {
            authorization: `Bearer ${token}`,
            // Verified live: without it the service answers some failures with
            // a redirect to a sign-in page instead of a status we can act on.
            behavioroverride: 'redirectAs404',
            accept: read === 'text' ? 'application/json' : '*/*',
          },
          // Never let a redirect carry the bearer somewhere we did not vet.
          redirect: 'error',
        },
        {
          read,
          maxBytes,
          service: 'Teams',
          timeoutMs: REQUEST_TIMEOUT_MS,
          fetchImpl: this.#fetchImpl,
        },
      );
      body = r.body as string | Uint8Array;
      return r.res;
    });
    return { status: res.status, headers: res.headers, body };
  }

  async #getJson<T>(url: string): Promise<T> {
    const { status, body } = await this.#send(url, 'text');
    this.#checkStatus(status);
    try {
      return JSON.parse(body as string) as T;
    } catch {
      // Never echo the body: it is the user's message content.
      throw new McpToolError('Teams chat service returned a response that is not JSON.', {
        status,
        kind: 'http',
      });
    }
  }

  /**
   * Map a non-2xx to an `McpToolError` carrying its status. No retry loop: a
   * 429 or 5xx is reported, and the model decides whether to try again. The
   * body is never included — it may be message content.
   */
  #checkStatus(status: number): void {
    if (status >= 200 && status < 300) return;
    if (status === 401) {
      if (this.#direct !== undefined) throw this.#directRejected();
      throw new McpToolError('Teams rejected the access token (401) even after a fresh capture.', {
        status,
        kind: 'credential_rejected',
        hint:
          'Re-open Teams in a signed-in teams.cloud.microsoft tab (sign in again if it ' +
          'asks), click into any chat, then retry.',
      });
    }
    if (status === 429) {
      throw new McpToolError('Teams chat service is rate-limiting requests (429).', {
        status,
        kind: 'http',
        hint: 'Wait a minute before retrying.',
      });
    }
    throw new McpToolError(`Teams chat service returned HTTP ${status}.`, {
      status,
      kind: 'http',
      ...(status >= 500 ? { hint: 'A Teams-side failure; retry shortly.' } : {}),
    });
  }

  /**
   * The next (older) page's URL, re-rooted onto the configured base.
   *
   * The service hands back an absolute `backwardLink`. Rather than trust its
   * origin with the bearer token, only its `/users/…` path and query are kept
   * and requested against our own base — so a link naming another host (the
   * service's internal one, or anything else) can never receive the token. A
   * link with no `/v1/users/` path is refused outright.
   */
  #nextLink(page: Page): string | undefined {
    const link = page._metadata?.backwardLink;
    if (typeof link !== 'string' || link === '') return undefined;
    let target: URL;
    try {
      target = new URL(link);
    } catch {
      throw new McpToolError('Refusing to follow a Teams paging link that is not a URL.');
    }
    const at = target.pathname.indexOf('/v1/users/');
    if (at < 0) {
      throw new McpToolError(`Refusing to follow a Teams paging link to ${target.origin}${target.pathname}.`, {
        hint: 'Paging links must be chat-service /v1/users/… paths.',
      });
    }
    return `${this.#baseUrl}${target.pathname.slice(at + '/v1'.length)}${target.search}`;
  }
}
