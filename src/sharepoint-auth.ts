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
 * re-pairing per host.
 *
 * Import this module LAZILY (`await import('./sharepoint-auth.js')`) from the
 * code that needs a token, as `auth-fetchproxy.ts` is, so a server run with
 * `TEAMS_DISABLE_FETCHPROXY` never has to load the bridge to boot.
 */
import { captureBearerHeader } from '@chrischall/mcp-utils/fetchproxy';
import { captureWindowMs } from './capture-window.js';
import { createCaptureTransport, getWsPort } from './auth-fetchproxy.js';
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

  return captureBearerHeader({
    // The one request we capture from: any call the browser makes to this
    // OneDrive host (the caller resolves it from the file's host via
    // `captureHostFor`). A concrete host, never a wildcard — see the module note.
    captures: [{ host, path: '/*', headerName: 'authorization' }],
    // The broad family domain as the trust scope. Same verb set as the chatsvc
    // capture, so an existing pairing that already trusts `sharepoint.com`
    // needs no re-approval.
    domains: SHAREPOINT_TRUST_DOMAINS,
    // Read per capture, not at import: see `capture-window.ts`.
    windowMs: captureWindowMs(env),
    port: getWsPort(),
    serverName: PACKAGE_NAME,
    version: VERSION,
    what: `a SharePoint token for ${host}`,
    guidance:
      `Click around Teams or open one of your own OneDrive files — that makes the ` +
      `browser request ${host}, which is where the SharePoint token is captured ` +
      'from — then retry.',
    createTransport: createCaptureTransport,
  });
}
