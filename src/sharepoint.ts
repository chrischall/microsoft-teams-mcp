/**
 * The SharePoint download client — fetches the CONTENTS of a shared file that
 * lives in a Teams chat.
 *
 * A shared file is not in the Teams AMS object store inline images come from; it
 * is in SharePoint, and reading it needs a SharePoint token, NOT the chat-service
 * token `chatsvc.ts` holds (chatsvc's token returns 401 on SharePoint, and Graph
 * `/shares` returns 403). Verified live 2026-10-11, the one route that works:
 *
 *   GET https://{host}/_api/v2.1/shares/u!{base64url(objectUrl)}/driveItem
 *       ?$select=id,size,@content.downloadUrl      [Authorization: Bearer <SP>]
 *   GET {@content.downloadUrl}                     [no auth — pre-signed]
 *   → the file's bytes
 *
 * The SP token is captured per host from the user's browser (see
 * `sharepoint-auth.ts`), is short-lived (~53 min) and host-family-scoped, so this
 * client holds one `TokenManager` per host — each with the same capture-on-401
 * recovery `ChatServiceClient` has. A 403 is expected (an account that cannot
 * open the file, or a host the captured token does not cover) and degrades
 * gracefully; it never crashes the tool.
 *
 * As with `chatsvc.ts`, once a token is held every request is a plain Node
 * `fetch` — the bridge is only ever touched to capture the token.
 */
import { decodeJwtExp, fetchBounded, McpToolError, parseBoolEnv } from '@chrischall/mcp-utils';
import { TokenManager, type BearerTokens } from '@chrischall/mcp-utils/session';
import { createSharePointTokenCache, type SharePointTokenCache } from './sharepoint-token-cache.js';

/**
 * Capture a SharePoint token for `host` lazily.
 *
 * `sharepoint-auth.ts` pulls in `@chrischall/mcp-utils/fetchproxy`, whose
 * optional `@fetchproxy/server` peer the `.mcpb` bundle ships no `node_modules`
 * for — so a top-level import would throw `ERR_MODULE_NOT_FOUND` before the
 * server could answer `initialize`. Deferring it (exactly as `chatsvc.ts` does
 * for its own capture) keeps boot free of the bridge.
 */
async function captureSharePointTokenLazily(host: string): Promise<string> {
  const mod = await import('./sharepoint-auth.js');
  return mod.captureSharePointTokenViaFetchproxy(host);
}

/**
 * Fallback lifetime for a token whose `exp` we cannot read. Real SP tokens are
 * JWTs carrying `exp` (~53 min out), so this is only reached for a malformed
 * capture; deliberately short — being wrong here costs one extra 401-and-replay.
 */
const UNKNOWN_TOKEN_TTL_MS = 10 * 60 * 1000;

/**
 * Sentinel standing in for a refresh token this flow does not have. The token
 * comes from the browser; "refreshing" means capturing a new one. `TokenManager`
 * only calls `refresh` when a refresh token is present, hence the sentinel.
 * (Copied from `chatsvc.ts`, where it is module-private.)
 */
const RECAPTURE_SENTINEL = 'fetchproxy:recapture';

/** Per-request deadline, covering the body read. */
const REQUEST_TIMEOUT_MS = 60_000;

/** Copied faithfully from `chatsvc.ts`: a captured token as `TokenManager` input. */
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

/** The subset of the driveItem response the download needs. */
interface DriveItemMeta {
  id?: string;
  size?: number;
  /** A pre-signed link, fetched with no auth. */
  '@content.downloadUrl'?: string;
}

export interface SharePointClientOptions {
  env?: NodeJS.ProcessEnv;
  /** Injected for tests so no suite ever touches the real bridge. */
  captureToken?: (host: string) => Promise<string>;
  fetchImpl?: typeof fetch;
}

export class SharePointClient {
  readonly #env: NodeJS.ProcessEnv;
  readonly #capture: (host: string) => Promise<string>;
  readonly #fetchImpl: typeof fetch;
  readonly #cache: SharePointTokenCache | null;
  /**
   * Deferred configuration error. The constructor never throws, so the server
   * still boots (and answers `tools/list`) with no credential present.
   */
  readonly #configError: McpToolError | null = null;
  /** One `TokenManager` per SharePoint host — SP tokens are host-scoped. */
  readonly #managers = new Map<string, TokenManager>();

  constructor(opts: SharePointClientOptions = {}) {
    this.#env = opts.env ?? process.env;
    this.#capture = opts.captureToken ?? captureSharePointTokenLazily;
    // A receiver-safe wrapper, never the bare global: older undici throws
    // `Illegal invocation` when `fetch` is called with a non-global receiver.
    this.#fetchImpl = opts.fetchImpl ?? ((input, init) => fetch(input, init));
    this.#cache = createSharePointTokenCache(this.#env);

    // Unlike chatsvc there is NO direct-token env var: a SharePoint token is
    // only ever captured from the browser, so with the bridge off there is no
    // way to get one. Defer that as a clean no_credential rather than letting a
    // download attempt the bridge.
    if (parseBoolEnv('TEAMS_DISABLE_FETCHPROXY', { env: this.#env })) {
      this.#configError = new McpToolError('SharePoint file download is disabled (no token source).', {
        kind: 'no_credential',
        hint:
          'Unset TEAMS_DISABLE_FETCHPROXY so a SharePoint token can be captured from a ' +
          'signed-in Teams browser tab when a shared file is opened.',
      });
    }
  }

  /** Whether a token could be obtained at all (the bridge is not disabled). */
  isConfigured(): boolean {
    return this.#configError === null;
  }

  /**
   * Download the contents of a shared file by its `objectUrl` (the `url` a file
   * attachment carries). Returns the bytes, the mime type and the byte count.
   *
   * The URL comes out of message content — i.e. it is attacker-influenced — so
   * it is vetted BEFORE a token is resolved: https only, and a concrete
   * `*.sharepoint.com` tenant host. Anything else would hand the SharePoint
   * bearer to whoever wrote the message.
   */
  async downloadFile(objectUrl: string, maxBytes: number): Promise<{ mimeType: string; data: Buffer; size: number }> {
    this.#requireConfigured();

    let target: URL;
    try {
      target = new URL(objectUrl);
    } catch {
      throw new McpToolError('Refusing to download a file whose link is not a URL.', {
        hint: 'Pass a file url exactly as a message listing returned it.',
      });
    }

    // `isSharePointHost` lives in the lazily-imported auth module (see
    // `captureSharePointTokenLazily`); loaded here too so the host is vetted with
    // the one authoritative check rather than a duplicate.
    const { isSharePointHost } = await import('./sharepoint-auth.js');
    const host = target.hostname;
    if (target.protocol !== 'https:' || !isSharePointHost(host)) {
      throw new McpToolError(`Refusing to download a file from ${target.protocol}//${host}.`, {
        hint: 'Only https files on a *.sharepoint.com tenant host are downloaded.',
      });
    }

    // `u!` + unpadded base64url of the full objectUrl (Node's `base64url` already
    // uses `-`/`_` and drops padding), matching the live spike.
    const share = `u!${Buffer.from(objectUrl, 'utf8').toString('base64url')}`;
    const driveItemUrl = `https://${host}/_api/v2.1/shares/${share}/driveItem?$select=id,size,@content.downloadUrl`;

    const meta = await this.#getDriveItem(host, driveItemUrl);
    const downloadUrl = meta['@content.downloadUrl'];
    if (typeof downloadUrl !== 'string' || downloadUrl === '') {
      throw new McpToolError('SharePoint returned no download link for the file.', {
        kind: 'http',
        hint: 'The file may be a folder or a format with no direct content; open it with its link.',
      });
    }

    return this.#fetchContent(downloadUrl, maxBytes);

    // fallback: sharepointIds route — the proven shares route above has not 404'd
    // in testing; the documented driveId/siteId/listItemUniqueId addressing stays
    // out until a real 404 warrants it (plan Task 3, optional).
  }

  #requireConfigured(): void {
    if (this.#configError) throw this.#configError;
  }

  /** The `TokenManager` for `host`, created (and cached) on first use. */
  #managerFor(host: string): TokenManager {
    const existing = this.#managers.get(host);
    if (existing) return existing;

    // The per-host cache view already swallows-and-reports its own write
    // failures (see `sharepoint-token-cache.ts`), so it is passed straight
    // through — no inline wrapping as chatsvc needs for its raw persistence.
    const persistence = this.#cache?.forHost(host);
    const mgr = new TokenManager({
      initial: async () => toBearerTokens(await this.#capture(host)),
      // No refresh grant — re-capture instead.
      refresh: async () => toBearerTokens(await this.#capture(host)),
      ...(persistence ? { persistence } : {}),
    });
    this.#managers.set(host, mgr);
    return mgr;
  }

  /**
   * The authed driveItem GET, with chatsvc's discipline: `behavioroverride:
   * redirectAs404` so a failure is a status rather than a sign-in redirect, and
   * `redirect: 'error'` so a redirect never carries the bearer somewhere unvetted.
   * `TokenManager.withAuth` supplies the token and, on a 401, re-captures once
   * and replays once — so a 401 that reaches here is the second one.
   */
  async #getDriveItem(host: string, url: string): Promise<DriveItemMeta> {
    let body = '';
    const res = await this.#managerFor(host).withAuth(async (token) => {
      const r = await fetchBounded(
        url,
        {
          method: 'GET',
          headers: {
            authorization: `Bearer ${token}`,
            accept: 'application/json',
            behavioroverride: 'redirectAs404',
          },
          redirect: 'error',
        },
        { read: 'text', service: 'SharePoint', timeoutMs: REQUEST_TIMEOUT_MS, fetchImpl: this.#fetchImpl },
      );
      body = r.body;
      return r.res;
    });

    const status = res.status;
    if (status === 403) {
      // mcp-utils' `kind` vocabulary has no 'forbidden'; the 403 is carried by
      // `status`, which the attachment tool reads to degrade to metadata + link.
      throw new McpToolError('SharePoint refused access to the file (403).', {
        status,
        kind: 'http',
        hint:
          'Your account cannot open this file, or it is on a SharePoint host your captured ' +
          'token does not cover. Open it once in Teams, then retry.',
      });
    }
    if (status === 404) {
      throw new McpToolError('SharePoint has no such file (404).', {
        status,
        kind: 'http',
        hint: 'The file may have moved or been deleted; open it with its link to check.',
      });
    }
    return this.#parseJson<DriveItemMeta>(status, body);
  }

  /**
   * Fetch the pre-signed download link with NO Authorization header. It 302s to
   * a CDN/blob host, which is safe to follow precisely because no bearer is
   * attached; the read is bounded to `maxBytes`. https only.
   */
  async #fetchContent(downloadUrl: string, maxBytes: number): Promise<{ mimeType: string; data: Buffer; size: number }> {
    let target: URL;
    try {
      target = new URL(downloadUrl);
    } catch {
      throw new McpToolError('SharePoint returned a download link that is not a URL.', { kind: 'http' });
    }
    if (target.protocol !== 'https:') {
      throw new McpToolError(`Refusing to fetch a download link over ${target.protocol}.`, {
        kind: 'http',
        hint: 'SharePoint download links are expected to be https.',
      });
    }

    const r = await fetchBounded(
      target.toString(),
      { method: 'GET', headers: { accept: '*/*' }, redirect: 'follow' },
      { read: 'bytes', maxBytes, service: 'SharePoint', timeoutMs: REQUEST_TIMEOUT_MS, fetchImpl: this.#fetchImpl },
    );
    this.#checkStatus(r.status);
    const mimeType = (r.headers.get('content-type') ?? 'application/octet-stream').split(';')[0].trim();
    const data = Buffer.from(r.body);
    return { mimeType, data, size: data.length };
  }

  #parseJson<T>(status: number, body: string): T {
    this.#checkStatus(status);
    try {
      return JSON.parse(body) as T;
    } catch {
      // Never echo the body: it may be the file's metadata or an error page.
      throw new McpToolError('SharePoint returned a response that is not JSON.', { status, kind: 'http' });
    }
  }

  /** Map a non-2xx to an `McpToolError` carrying its status. The body is never included. */
  #checkStatus(status: number): void {
    if (status >= 200 && status < 300) return;
    if (status === 401) {
      throw new McpToolError('SharePoint rejected the access token (401) even after a fresh capture.', {
        status,
        kind: 'credential_rejected',
        hint: 'Open the shared file once in a signed-in Teams tab to mint a fresh SharePoint token, then retry.',
      });
    }
    if (status === 429) {
      throw new McpToolError('SharePoint is rate-limiting requests (429).', {
        status,
        kind: 'http',
        hint: 'Wait a minute before retrying.',
      });
    }
    throw new McpToolError(`SharePoint returned HTTP ${status}.`, {
      status,
      kind: 'http',
      ...(status >= 500 ? { hint: 'A SharePoint-side failure; retry shortly.' } : {}),
    });
  }
}
