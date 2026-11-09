/**
 * Bootstrap-only fetchproxy auth for SharePoint shared-file download.
 *
 * A shared file in a Teams chat lives in SharePoint, not the Teams AMS object
 * store that holds inline images, so downloading its contents needs a
 * SharePoint token — NOT the chat-service token `auth-fetchproxy.ts` captures.
 * The token is captured from the user's OneDrive host
 * (`{tenant}-my.sharepoint.com`): the browser authenticates that host during
 * normal Teams use and its token reads team-site files too, whereas it never
 * issues a capturable request to a tenant-root team-site host. The caller
 * resolves this capture host from the file's host via `captureHostFor` and
 * passes it as `host` here. We snapshot the `Authorization: Bearer …` header off
 * a request to it, then talk to the SharePoint v2.1 shares API with a plain Node
 * `fetch`.
 *
 * Verified live 2026-10-11 against a real Teams tenant:
 *   - the working token has `aud` = the SharePoint/Office 365 resource GUID
 *     (`00000003-0000-0ff1-ce00-000000000000`), not the chat service's — the
 *     chatsvc token returns 401 on SharePoint, and Graph `/shares` returns 403;
 *   - that token authenticates a plain server-side GET of
 *     `/_api/v2.1/shares/u!{base64url(objectUrl)}/driveItem`, whose
 *     `@content.downloadUrl` is a pre-signed link fetched with no auth;
 *   - it is short-lived (~53 min) and host-family-scoped: the
 *     `{tenant}-my.sharepoint.com` OneDrive token also worked on
 *     `{tenant}.sharepoint.com`, but got 403 on a separate OneDrive-region host
 *     — so that region needs its own captured token.
 *
 * The capture layer cannot wildcard the host: `validateCaptureHeaderDecls` runs
 * each capture host through `HOSTNAME_RE`, which rejects `*.sharepoint.com`. The
 * capture host is therefore a concrete OneDrive host, resolved from the file's
 * host by `captureHostFor` at download time.
 * The pairing TRUST SCOPE may still be the broad `sharepoint.com` —
 * `hostMatchesAnyDomain` accepts any `*.sharepoint.com` subdomain via
 * `endsWith('.' + domain)` — so one domain covers every tenant SP host without
 * re-pairing per host. That domain is one of the one bridge scope every capture
 * of this server declares (`bridge-scope.ts`); a capture host is added to it,
 * and remembered, once a capture of it succeeds.
 *
 * Import this module LAZILY (`await import('./sharepoint-auth.js')`) from the
 * code that needs a token, as `auth-fetchproxy.ts` is, so a server run with
 * `TEAMS_DISABLE_FETCHPROXY` never has to load the bridge to boot.
 */
import { captureBearerHeader } from '@chrischall/mcp-utils/fetchproxy';
import { captureWindowMs } from './capture-window.js';
import { bridgeScopedTransport, getWsPort } from './auth-fetchproxy.js';
import { BRIDGE_TRUST_DOMAINS, isSharePointHost, rememberSharePointHost, sharePointCaptureDecl } from './bridge-scope.js';
import { PACKAGE_NAME, VERSION } from './version.js';

/**
 * The domains the SharePoint capture trusts: the one bridge scope
 * (`bridge-scope.ts`), whose `sharepoint.com` covers every tenant SP host, while
 * the per-host capture decl below stays a concrete hostname (the capture layer
 * rejects a wildcard).
 */
export const SHAREPOINT_TRUST_DOMAINS = [...BRIDGE_TRUST_DOMAINS];

export { isSharePointHost } from './bridge-scope.js';

/**
 * The host to capture a SharePoint token FROM, for a file that lives on
 * `fileHost`.
 *
 * Verified live 2026-10-11: the browser authenticates the user's OneDrive host
 * (`<tenant>-my.sharepoint.com`) during normal Teams use, but never issues a
 * capturable page request to a tenant-root team-site host
 * (`<tenant>.sharepoint.com`) — so capturing from the file's own host times out
 * for a team-site file. The `-my` token, however, downloads team-site files too
 * (it worked cross-host to `transunion.sharepoint.com` in the spike). So:
 *   - a file already on a `*-my` host → capture from that host (incl. a separate
 *     region OneDrive like `<tenant>ind-my`, which needs its own token);
 *   - a file on a `<tenant>.sharepoint.com` root → capture from `<tenant>-my`.
 * The result always satisfies {@link isSharePointHost}. One `-my` token thus
 * serves every file on both that OneDrive and its tenant root, which is why the
 * client keys its per-host token cache by this capture host.
 */
export function captureHostFor(fileHost: string): string {
  const h = fileHost.toLowerCase();
  if (/-my\.sharepoint\.com$/.test(h)) return h;
  const m = /^([a-z0-9][a-z0-9-]*)\.sharepoint\.com$/.exec(h);
  return m ? `${m[1]}-my.sharepoint.com` : h;
}

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

  const token = await captureBearerHeader({
    // The one request we capture from: any call the browser makes to this
    // OneDrive host (the caller resolves it from the file's host via
    // `captureHostFor`). A concrete host, never a wildcard — see the module note.
    captures: [sharePointCaptureDecl(host)],
    // The one bridge scope (`bridge-scope.ts`), with this host's capture added
    // to the remembered ones so the hello asks for exactly what it is about to
    // read.
    domains: SHAREPOINT_TRUST_DOMAINS,
    // Read per capture, not at import: see `capture-window.ts`.
    windowMs: captureWindowMs(env),
    port: getWsPort(),
    serverName: PACKAGE_NAME,
    version: VERSION,
    what: `a SharePoint token for ${host}`,
    guidance:
      `Reload your signed-in Teams tab and retry while it loads — that makes the ` +
      `page call ${host} with the SharePoint token we capture. Opening a file does ` +
      'not reliably do this.',
    createTransport: bridgeScopedTransport(env, host),
  });
  // Granted and captured: every later hello declares this host too.
  rememberSharePointHost(host, env);
  return token;
}
