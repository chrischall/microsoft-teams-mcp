/**
 * Bootstrap-only fetchproxy auth for SharePoint shared-file download.
 *
 * A shared file in a Teams chat lives in SharePoint, not the Teams AMS object
 * store that holds inline images, so downloading its contents needs a
 * SharePoint token — NOT the chat-service token `auth-fetchproxy.ts` captures.
 * When the user opens such a file in Teams, the browser requests the file's own
 * SharePoint host (`{tenant}-my.sharepoint.com` or `{tenant}.sharepoint.com`)
 * with `Authorization: Bearer …`; we snapshot that header, then talk to the
 * SharePoint v2.1 shares API with a plain Node `fetch`.
 *
 * Verified live 2026-10-11 against a real Teams tenant:
 *   - the working token has `aud` = the SharePoint/Office 365 resource GUID
 *     (`00000003-0000-0ff1-ce00-000000000000`), not the chat service's — the
 *     chatsvc token returns 401 on SharePoint, and Graph `/shares` returns 403;
 *   - that token authenticates a plain server-side GET of
 *     `/_api/v2.1/shares/u!{base64url(objectUrl)}/driveItem`, whose
 *     `@content.downloadUrl` is a pre-signed link fetched with no auth;
 *   - it is short-lived (~53 min) and host-family-scoped: a
 *     `{tenant}-my.sharepoint.com` token also worked on `{tenant}.sharepoint.com`
 *     but got 403 on a separate OneDrive-region host. So each distinct SP host
 *     needs its own captured token.
 *
 * The capture layer cannot wildcard the host: `validateCaptureHeaderDecls` runs
 * each capture host through `HOSTNAME_RE`, which rejects `*.sharepoint.com`. The
 * capture host is therefore the file's concrete host, derived at download time.
 * The pairing TRUST SCOPE may still be the broad `sharepoint.com` —
 * `hostMatchesAnyDomain` accepts any `*.sharepoint.com` subdomain via
 * `endsWith('.' + domain)` — so one domain covers every tenant SP host without
 * re-pairing per host.
 *
 * Import this module LAZILY (`await import('./sharepoint-auth.js')`) from the
 * code that needs a token, as `auth-fetchproxy.ts` is, so a server run with
 * `TEAMS_DISABLE_FETCHPROXY` never has to load the bridge to boot.
 */
import { createBootstrapOpts, createFetchproxyTransport } from '@chrischall/mcp-utils/fetchproxy';
import { bridgeDeadlineMs, captureWindowMs } from './capture-window.js';
import { getWsPort, stripBearer } from './auth-fetchproxy.js';
import { PACKAGE_NAME, VERSION } from './version.js';

/**
 * The extension's trust boundary for the SharePoint path: the broad family
 * domain. `hostMatchesAnyDomain` treats this as "any `*.sharepoint.com`", so a
 * single paired domain covers every tenant SP host, while the per-host capture
 * decl below stays a concrete hostname (the capture layer rejects a wildcard).
 */
export const SHAREPOINT_TRUST_DOMAINS = ['sharepoint.com'] as const;

/**
 * True for a concrete SharePoint tenant host (`{label}.sharepoint.com`), such
 * as `transunion-my.sharepoint.com` or `transunion.sharepoint.com`. Rejects a
 * bare `sharepoint.com`, a suffix spoof (`sharepoint.com.evil.com`), and any
 * non-SharePoint host. Host is lowercased before matching.
 */
export function isSharePointHost(host: string): boolean {
  return /^[a-z0-9][a-z0-9-]*\.sharepoint\.com$/.test(host.toLowerCase());
}

/**
 * Shortest plausible token. Real ones are several KB of JWT; this only guards
 * against an empty or echoed header, not against a wrong-audience token.
 */
const MIN_TOKEN_LENGTH = 100;

/**
 * Capture a SharePoint access token for `host` from the user's signed-in Teams
 * tab. Returns the bare token (no `Bearer ` prefix).
 *
 * `host` must be a concrete SharePoint host — it becomes the capture decl, and
 * sending a bearer later depends on it being one we vetted.
 */
export async function captureSharePointTokenViaFetchproxy(
  host: string,
  env: NodeJS.ProcessEnv = process.env,
): Promise<string> {
  if (!isSharePointHost(host)) {
    throw new Error(
      `refusing to capture a SharePoint token for ${JSON.stringify(host)}: not a ` +
        '*.sharepoint.com host',
    );
  }

  // The one request we capture from: any call the browser makes to this file's
  // SharePoint host. A concrete host (never a wildcard) — see the module note.
  const captureDecl = { host, path: '/*', headerName: 'authorization' } as const;

  // Read per capture, not at import: see `capture-window.ts`.
  const windowMs = captureWindowMs(env);
  const transport = createFetchproxyTransport({
    // `createBootstrapOpts` derives the capability set from the declarations:
    // `capture_request_header` beside mcp-utils' always-on `fetch` base verb.
    // Same verb set as the chatsvc capture, so an existing pairing that already
    // trusts `sharepoint.com` needs no re-approval.
    ...createBootstrapOpts({
      domains: [...SHAREPOINT_TRUST_DOMAINS],
      bootstrap: { captureHeaders: [{ ...captureDecl }] },
    }),
    port: getWsPort(),
    // Both: the explicit deadline for every bridge version, and the declared
    // window so mcp-utils' own derivation agrees with it.
    fetchTimeoutMs: bridgeDeadlineMs(windowMs),
    captureWindowMs: windowMs,
    serverName: PACKAGE_NAME,
    version: VERSION,
  });

  try {
    // `start()` loads the identity keypair. It does not bind the port or dial —
    // that stays lazy until the first verb — but skipping it makes the first
    // verb throw "ensureConnected called before listen()", which reads exactly
    // like a pairing problem and is not one.
    await transport.start();

    try {
      const raw = await transport.server.captureRequestHeader({
        ...captureDecl,
        timeoutMs: windowMs,
      });
      if (typeof raw !== 'string') {
        throw new Error(`${host}: capture returned no header`);
      }
      const token = stripBearer(raw);
      if (token.length < MIN_TOKEN_LENGTH) {
        throw new Error(
          `${host}: authorization header was present but too short to be a ` +
            `token (${token.length} chars)`,
        );
      }
      return token;
    } catch (e) {
      const detail = (e as Error).message;
      throw new Error(
        `could not capture a SharePoint token for ${host} (${detail}). ` +
          `Open the shared file once in Teams — that makes the browser request ${host}, ` +
          'which is what the SharePoint token is captured from — then retry.',
      );
    }
  } finally {
    try {
      await transport.close();
    } catch {
      /* the token, or the original failure, is what matters */
    }
  }
}
