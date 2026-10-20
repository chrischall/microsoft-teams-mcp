/**
 * The Teams chat-service REST client.
 *
 * Auth is a bearer token whose only source is the user's browser (see
 * `auth-fetchproxy.ts`), or a `TEAMS_ACCESS_TOKEN` supplied directly. Once
 * held, every request is a plain Node `fetch` — the bridge is never on the hot
 * path.
 *
 * Reads are `listConversations`, `getMessages` (which also reads the Activity
 * feed, conversation `48:notifications`), `getMessage` and
 * `fetchImage`; the writes are `sendMessage`, `setReaction`,
 * `removeReaction` and `setConsumptionHorizon`, which only the confirm-gated
 * `teams_send_message`, `teams_react` and `teams_mark_read` (and the read
 * tools' opt-in `mark_read`) call. A read never writes the read marker
 * (`consumptionhorizon`) on its own, so reading a chat through this MCP leaves
 * the user's unread badges alone unless they ask for it to be marked read.
 *
 * Verified live 2026-10-10 against a real Teams tenant — see the "Verified
 * live facts" section of `docs/plans/2026-10-10-chatsvc-api.md` for the
 * response shapes the types below describe.
 */
import {
  decodeJwtClaim,
  fetchBounded,
  errorKindOf,
  McpToolError,
  parseBoolEnv,
  readEnvVar,
  RequestTimeoutError,
  WriteOutcomeUnknownError,
} from '@chrischall/mcp-utils';
import { createCapturedTokenManager, type TokenManager } from '@chrischall/mcp-utils/session';
import { createTokenCache, reportCacheWriteFailure } from './token-cache.js';
import { SELF_CHAT } from './model.js';
import { describeTransportFailure } from './network-error.js';

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

/** Per-request deadline, covering the body read. */
const REQUEST_TIMEOUT_MS = 60_000;

/** Paging caps — a runaway `backwardLink` chain must not hold a tool call open. */
const MAX_CONVERSATION_PAGES = 10;
const MAX_MESSAGE_PAGES = 20;
/** Messages per `getMessages` request unless the caller asks for fewer. */
const DEFAULT_MESSAGE_PAGE_SIZE = 200;

/** Largest inline image we will buffer (they measured 8–18 KB live). */
const MAX_IMAGE_BYTES = 5 * 1024 * 1024;

/**
 * Inline images live on Skype's AMS object store (`us-api.asm.skype.com`, the
 * host varies by region). The bearer is sent ONLY to hosts under this suffix.
 */
const IMAGE_HOST_SUFFIX = '.asm.skype.com';

/**
 * The Teams gateway in front of the same image store. Verified live
 * 2026-10-12: `GET https://us-prod.asyncgw.teams.microsoft.com/v1/objects/<id>/views/imgo`
 * with the chat-service bearer returns the identical bytes for objects of
 * either region (`0-sin-…` and `0-eus-…`). Some networks reset every
 * connection to `*.asm.skype.com` while allowing teams.microsoft.com, so an
 * image the store cannot be REACHED at is retried here — never one the store
 * answered. Only ever reached by rewriting a vetted `*.asm.skype.com` URL,
 * never from a URL in message content.
 */
const IMAGE_GATEWAY_HOST = 'us-prod.asyncgw.teams.microsoft.com';

/** What a transport failure calls the far side: the image store, or the chat service. */
function serviceFor(url: string): string {
  try {
    const host = new URL(url).hostname;
    if (host.endsWith(IMAGE_HOST_SUFFIX)) return "Teams' image store";
    return host === IMAGE_GATEWAY_HOST ? 'the Teams image gateway' : 'the Teams chat service';
  } catch {
    return 'the Teams chat service';
  }
}

/**
 * A write that failed after it may have reached Teams. A timeout or a
 * transport failure (dropped connection, broken body, the caller cancelling)
 * leaves the outcome unknown; an `McpToolError` raised before anything was
 * sent (an oversized body cap, a bad URL) is a real failure and passes through.
 */
function outcomeUnknown(method: string, e: unknown, hint: string): unknown {
  if (e instanceof RequestTimeoutError) {
    return new WriteOutcomeUnknownError('Teams', method, { timeoutMs: e.timeoutMs, cause: e, hint });
  }
  if (e instanceof McpToolError) return e;
  return new WriteOutcomeUnknownError('Teams', method, { cause: e, hint });
}

const SEND_OUTCOME_HINT =
  'The message may already have been posted — check the conversation (teams_list_messages) before retrying; do not resend blindly.';

const REACTION_OUTCOME_HINT =
  'The reaction change may already have been applied — check the message (teams_list_messages) before retrying.';

const HORIZON_OUTCOME_HINT =
  'The read marker may already have moved — check the conversation (teams_list_messages) before retrying; ' +
  'moving it again to the same message is harmless.';

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
  /** As on a full message: `deletetime` set means deleted. */
  properties?: Record<string, unknown>;
}

/** One entry of `GET /users/ME/conversations`. */
export interface RawConversation {
  /** `…@thread.v2` group/meeting, `…@thread.tacv2` channel, `…@unq.gbl.spaces` 1:1. */
  id: string;
  type?: string;
  version?: number;
  targetLink?: string;
  lastMessage?: RawLastMessage;
  /**
   * The arrival (id, epoch ms) of the conversation's most recently changed
   * message, and that change's time. Live (2026-10-12), in a channel this
   * moves for a new REPLY while `lastMessage` and `lastimreceivedtime` stay on
   * the newest root post — so only these show that a thread got a reply. Seen
   * as numbers; a numeric string is read too.
   */
  lastUpdatedMessageId?: number | string;
  lastUpdatedMessageVersion?: number | string;
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
  /**
   * When the message last changed (epoch ms, a numeric string): its arrival,
   * or a later edit or reaction. Live (2026-10-12), pages come newest CHANGE
   * first, so an old message edited recently is served ahead of newer ones.
   */
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
  /** Most pages to request (default 20) — a read that must reach far back asks for more. */
  maxPages?: number;
  /**
   * Messages per request (default 200, the service's largest page). A read
   * that wants only a few recent messages asks for fewer, so a cold read of
   * many conversations is not 200 messages each.
   */
  pageSize?: number;
  /**
   * Called when the read ended at `max` or at the page cap rather than at a
   * `stopAt` boundary or the start of the history — i.e. there may be more.
   */
  onCapped?: () => void;
  /**
   * Called between pages — after page `pages` is read (with `messages`
   * collected so far) and before the next is requested. A long read reports
   * progress here; a throw ends the read with that error and fetches nothing more.
   */
  onPage?: (pages: number, messages: number) => void | Promise<void>;
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

  /**
   * Epoch ms at which the current token expires, or `null` if none is held —
   * including after a capture that failed, when the token manager exists but
   * holds nothing (it reports that as expiry 0, i.e. 1970).
   */
  tokenExpiresAt(): number | null {
    const at = this.#tokens?.getExpiresAt();
    return at !== undefined && Number.isFinite(at) && at > 0 ? at : null;
  }

  /**
   * Hold a token — the cached one, or a fresh capture — without calling the
   * API. The healthcheck resolves the credential with this before probing, so
   * a capture failure is reported as one rather than as a failed request.
   */
  async ensureToken(): Promise<void> {
    this.#requireConfigured();
    await this.#tokenManager().getAccessToken();
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
   * The signed-in user's display name, from the token's `name` claim — what a
   * sent message carries as `imdisplayname`. `''` when the token has none.
   */
  async meDisplayName(): Promise<string> {
    this.#requireConfigured();
    const token = await this.#tokenManager().getAccessToken();
    try {
      const name = decodeJwtClaim(token, 'name');
      return typeof name === 'string' ? name : '';
    } catch {
      return '';
    }
  }

  /**
   * Every conversation the user is in, across up to 10 pages, minus the
   * `48:*` system threads (calendar, notifications, …) which are not chats —
   * except `48:notes`, the user's self-chat, which is.
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
        if (typeof c?.id === 'string' && (!c.id.startsWith('48:') || c.id === SELF_CHAT)) out.push(c);
      }
      url = this.#nextLink(body);
    }
    return out;
  }

  /**
   * A conversation's messages, newest first, paging back until `stopAt`
   * accepts a message, `max` are collected, the pages run out, or `maxPages`
   * (default 20) have been read.
   */
  async getMessages(id: string, opts: GetMessagesOptions): Promise<RawMessage[]> {
    this.#requireConfigured();
    const out: RawMessage[] = [];
    if (!(opts.max > 0)) return out;
    const seen = new Set<string>();
    let url: string | undefined =
      `${this.#baseUrl}/users/ME/conversations/${encodeURIComponent(id)}/messages` +
      `?view=msnp24Equivalent|supportsMessageProperties&pageSize=${opts.pageSize ?? DEFAULT_MESSAGE_PAGE_SIZE}&startTime=1`;
    for (let page = 0; url !== undefined; page++) {
      if (page >= (opts.maxPages ?? MAX_MESSAGE_PAGES)) {
        opts.onCapped?.();
        break;
      }
      const body: MessagesPage = await this.#getJson<MessagesPage>(url);
      const msgs = body.messages ?? [];
      if (msgs.length === 0) break;
      // A page is newest first but not strictly ordered (seen live
      // 2026-10-10), so the page `stopAt` fires in is finished — skipping
      // each message it rejects — and only the NEXT page is not fetched.
      let stopped = false;
      for (const m of msgs) {
        // Live (2026-10-14) the service served one message — same id, same version — on two pages
        // (2 of 567). The first copy stands; a repeat is neither returned nor shown to `stopAt`,
        // whose limit counting would otherwise count it twice.
        if (seen.has(m.id)) continue;
        seen.add(m.id);
        if (opts.stopAt?.(m)) {
          stopped = true;
          continue;
        }
        out.push(m);
        if (out.length >= opts.max) {
          opts.onCapped?.();
          return out;
        }
      }
      if (stopped) return out;
      url = this.#nextLink(body);
      if (url !== undefined) await opts.onPage?.(page + 1, out.length);
    }
    return out;
  }

  /**
   * One message by id, or `undefined` when the service answers 404. Used to
   * quote a message (its sender and text go into the quote) and to fill in a
   * channel thread's root post when it is older than a read's window — a
   * channel root fetched by its id answers like any message (verified live
   * 2026-10-12). An id the channel does not hold was answered 403, not 404,
   * live, so that surfaces as an error.
   */
  async getMessage(conversationId: string, messageId: string): Promise<RawMessage | undefined> {
    this.#requireConfigured();
    const url =
      `${this.#baseUrl}/users/ME/conversations/${encodeURIComponent(conversationId)}` +
      `/messages/${encodeURIComponent(messageId)}`;
    const { status, body } = await this.#send(url, 'text');
    if (status === 404) return undefined;
    return this.#parseJson<RawMessage>(status, body as string);
  }

  /**
   * Post a message (the body `compose.ts` builds). `conversationId` is the
   * chat or channel — or, for a threaded channel reply, `<channelId>;messageid=<rootId>`,
   * encoded as one path segment. Resolves the new message's id from the
   * `Location` header when the service sends one.
   *
   * Not idempotent, so a request that was SENT but got no usable answer (a
   * timeout, a dropped connection, a 5xx from a gateway or a failing backend)
   * is a `WriteOutcomeUnknownError` — the post may already be in the chat,
   * and a blind retry would post it twice. A 4xx is a refusal and stays a
   * plain error. A 401 is still replayed once after a re-capture: Teams
   * refused that request, so nothing was posted.
   */
  async sendMessage(conversationId: string, body: unknown): Promise<{ id?: string }> {
    this.#requireConfigured();
    const url = `${this.#baseUrl}/users/ME/conversations/${encodeURIComponent(conversationId)}/messages`;
    const { status, headers } = await this.#send(url, 'text', undefined, { method: 'POST', body, hint: SEND_OUTCOME_HINT });
    this.#checkWriteStatus('POST', status, SEND_OUTCOME_HINT);
    const id = /\/messages\/([^/?#]+)\/?(?:[?#].*)?$/.exec(headers.get('location') ?? '')?.[1];
    return id !== undefined ? { id: decodeURIComponent(id) } : {};
  }

  /**
   * React to a message: `PUT …/messages/{id}/properties?name=emotions` with
   * `{ emotions: { key, value: <epoch ms> } }` (captured live). Only the
   * message's `emotions` property is written — never the user's
   * `recentReactions` preference, which Teams updates as a UI nicety.
   *
   * A timeout, a dropped connection or a 5xx after the request was sent is a
   * `WriteOutcomeUnknownError`, as for a post: the change may have landed.
   * Repeating it is harmless (the same key twice is still one reaction), but
   * the tool reads the message back rather than reporting a plain failure for
   * a change that may well be visible already.
   */
  async setReaction(conversationId: string, messageId: string, key: string, at: number = Date.now()): Promise<void> {
    await this.#writeEmotions('PUT', conversationId, messageId, { key, value: at });
  }

  /** Withdraw the user's reaction: `DELETE` on the same property (unverified live — see the plan). */
  async removeReaction(conversationId: string, messageId: string, key: string): Promise<void> {
    await this.#writeEmotions('DELETE', conversationId, messageId, { key });
  }

  async #writeEmotions(
    method: 'PUT' | 'DELETE',
    conversationId: string,
    messageId: string,
    emotions: { key: string; value?: number },
  ): Promise<void> {
    this.#requireConfigured();
    const url =
      `${this.#baseUrl}/users/ME/conversations/${encodeURIComponent(conversationId)}` +
      `/messages/${encodeURIComponent(messageId)}/properties?name=emotions`;
    const { status } = await this.#send(url, 'text', undefined, { method, body: { emotions }, hint: REACTION_OUTCOME_HINT });
    this.#checkWriteStatus(method, status, REACTION_OUTCOME_HINT);
  }

  /**
   * Move the user's read marker in a conversation to `messageId`:
   * `PUT …/conversations/{id}/properties?name=consumptionhorizon` with
   * `{ consumptionhorizon: "<id>;<client ms>;<id>" }` — the shape `readHorizon`
   * parses, its first segment the message's arrival ms (which is its id).
   * Teams clears the unread badge for everything up to it.
   *
   * Verified live 2026-10-12 on `48:notes`: 200, and the listing then carries
   * the new horizon. A move BACKWARDS is also answered 200 but ignored, so a
   * 200 does not prove the marker moved; never sending one is the caller's
   * rule (`teams_mark_read`), which reports it as a no-op instead. A lost
   * answer is a `WriteOutcomeUnknownError`, as for a reaction.
   */
  async setConsumptionHorizon(conversationId: string, messageId: string, at: number = Date.now()): Promise<void> {
    this.#requireConfigured();
    const url =
      `${this.#baseUrl}/users/ME/conversations/${encodeURIComponent(conversationId)}` +
      '/properties?name=consumptionhorizon';
    const body = { consumptionhorizon: `${messageId};${at};${messageId}` };
    const { status } = await this.#send(url, 'text', undefined, { method: 'PUT', body, hint: HORIZON_OUTCOME_HINT });
    this.#checkWriteStatus('PUT', status, HORIZON_OUTCOME_HINT);
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
    try {
      return await this.#getImage(target);
    } catch (e) {
      if (errorKindOf(e) !== 'transport') throw e;
      const viaGateway = new URL(target.pathname + target.search, `https://${IMAGE_GATEWAY_HOST}`);
      try {
        return await this.#getImage(viaGateway);
      } catch (e2) {
        if (errorKindOf(e2) !== 'transport') throw e2;
        throw new McpToolError(
          `${(e as McpToolError).message} The Teams gateway was tried too: ${(e2 as McpToolError).message}`,
          { kind: 'transport', hint: (e as McpToolError).hint, cause: e },
        );
      }
    }
  }

  async #getImage(target: URL): Promise<{ mimeType: string; data: Buffer }> {
    const { status, headers, body } = await this.#send(target.toString(), 'bytes', MAX_IMAGE_BYTES);
    this.#checkStatus(status);
    const mimeType = (headers.get('content-type') ?? 'application/octet-stream').split(';')[0].trim();
    return { mimeType, data: Buffer.from(body as Uint8Array) };
  }

  #requireConfigured(): void {
    if (this.#configError) throw this.#configError;
  }

  #tokenManager(): TokenManager {
    // Captures on first use (or loads a still-valid cached capture), re-captures
    // on a 401 or inside the skew window, and expires a JWT at its `exp` (an
    // opaque token after 10 minutes). A token supplied as TEAMS_ACCESS_TOKEN is
    // never cached and never "refreshed" from the browser: that would silently
    // swap the operator's chosen credential for whichever account a tab is
    // signed into. A failed cache write is reported, never fatal.
    this.#tokens ??= createCapturedTokenManager({
      capture: () => this.#capture(),
      direct: this.#direct,
      onDirectRejected: () => this.#directRejected(),
      persistence: this.#direct === undefined ? createTokenCache(this.#env) : null,
      onPersistError: reportCacheWriteFailure,
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
    /** A write: a lost answer is a `WriteOutcomeUnknownError` carrying `hint`. */
    write?: { method: 'POST' | 'PUT' | 'DELETE'; body: unknown; hint: string },
  ): Promise<{ status: number; headers: Headers; body: string | Uint8Array }> {
    let body: string | Uint8Array = '';
    const method = write?.method ?? 'GET';
    const res = await this.#tokenManager().withAuth(async (token) => {
      const init: RequestInit = {
        method,
        headers: {
          authorization: `Bearer ${token}`,
          // Verified live: without it the service answers some failures with
          // a redirect to a sign-in page instead of a status we can act on.
          behavioroverride: 'redirectAs404',
          accept: read === 'text' ? 'application/json' : '*/*',
          ...(write ? { 'content-type': 'application/json' } : {}),
        },
        ...(write ? { body: JSON.stringify(write.body) } : {}),
        // Never let a redirect carry the bearer somewhere we did not vet.
        redirect: 'error',
      };
      const opts = { read, maxBytes, service: 'Teams', timeoutMs: REQUEST_TIMEOUT_MS, fetchImpl: this.#fetchImpl };
      let r: Awaited<ReturnType<typeof fetchBounded<'text' | 'bytes'>>>;
      try {
        r = await fetchBounded(url, init, opts);
      } catch (e) {
        if (write) throw outcomeUnknown(method, e, write.hint);
        throw describeTransportFailure(e, url, serviceFor(url)) ?? e;
      }
      body = r.body as string | Uint8Array;
      return r.res;
    });
    return { status: res.status, headers: res.headers, body };
  }

  async #getJson<T>(url: string): Promise<T> {
    const { status, body } = await this.#send(url, 'text');
    return this.#parseJson<T>(status, body as string);
  }

  #parseJson<T>(status: number, body: string): T {
    this.#checkStatus(status);
    try {
      return JSON.parse(body) as T;
    } catch {
      // Never echo the body: it is the user's message content.
      throw new McpToolError('Teams chat service returned a response that is not JSON.', {
        status,
        kind: 'http',
      });
    }
  }

  /**
   * A write's status: a 5xx means the request reached Teams and failed
   * somewhere past the front door — a gateway that lost the backend's answer
   * (502/504) or a backend that may have committed before failing — so the
   * outcome is unknown, never a plain "retry shortly". Anything else maps as
   * for a read.
   */
  #checkWriteStatus(method: string, status: number, hint: string): void {
    if (status >= 500 && status < 600) {
      throw new WriteOutcomeUnknownError('Teams', method, {
        cause: new McpToolError(`Teams chat service returned HTTP ${status}.`, { status, kind: 'http' }),
        hint,
      });
    }
    this.#checkStatus(status);
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
