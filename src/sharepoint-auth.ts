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
 *   - it is short-lived.
 *
 * Verified live 2026-11-09: one token serves every SharePoint host of the
 * tenant. Tokens captured from `{tenant}-my` and from a separate OneDrive-region
 * host `{tenant}ind-my` each returned 200 on both of those hosts and on the
 * `{tenant}` team-site root. (The 2026-10-11 spike read a 403 on the region host
 * as the token not covering it; it was the file's own permissions.)
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
 * and remembered, once a capture of it succeeds. Since any granted host yields
 * a token every host accepts, later captures read off the granted hosts and
 * never widen the scope.
 *
 * Import this module LAZILY (`await import('./sharepoint-auth.js')`) from the
 * code that needs a token, as `auth-fetchproxy.ts` is, so a server run with
 * `TEAMS_DISABLE_FETCHPROXY` never has to load the bridge to boot.
 */
import { captureBearerHeader, type CaptureTransport } from '@chrischall/mcp-utils/fetchproxy';
import { McpToolError } from '@chrischall/mcp-utils';
import { captureWindowMs } from './capture-window.js';
import { bridgeScopedTransport, getWsPort } from './auth-fetchproxy.js';
import { BRIDGE_TRUST_DOMAINS, isSharePointHost, rememberSharePointHost, sharePointCaptureDecl } from './bridge-scope.js';
import { PACKAGE_NAME, VERSION } from './version.js';

/**
 * The domains the SharePoint capture trusts: the one bridge scope
 * (`bridge-scope.ts`), whose `sharepoint.com` covers every tenant SP host, while
 * each capture decl below stays a concrete hostname (the capture layer
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
 *   - a file already on a `*-my` host → capture from that host;
 *   - a file on a `<tenant>.sharepoint.com` root → capture from `<tenant>-my`.
 * The result always satisfies {@link isSharePointHost}. Used only for the
 * first capture, before any host has been granted: after that the client
 * captures off the granted hosts, whose token serves every host.
 */
export function captureHostFor(fileHost: string): string {
  const h = fileHost.toLowerCase();
  if (/-my\.sharepoint\.com$/.test(h)) return h;
  const m = /^([a-z0-9][a-z0-9-]*)\.sharepoint\.com$/.exec(h);
  return m ? `${m[1]}-my.sharepoint.com` : h;
}

/**
 * Capture a SharePoint access token from the user's signed-in Teams tab, off
 * the first request the page makes to any of `hosts`. Returns the bare token
 * (no `Bearer ` prefix).
 *
 * One token serves every SharePoint host of the tenant, so the caller passes
 * the hosts the bridge has already granted (any of them yields a usable
 * token), or the file's own capture host when none has been granted yet. Each
 * must be a concrete SharePoint host — it becomes a capture decl.
 *
 * A host the bridge has not granted is refused by the extension, which queues
 * a scope-update offer in its popup meanwhile. The bridge's own advice for that
 * ("revoke and re-pair") is heavier than the fix, so a capture that failed
 * with a refusal throws an `McpToolError` saying to approve the offer.
 */
export async function captureSharePointTokenViaFetchproxy(
  hosts: string | readonly string[],
  env: NodeJS.ProcessEnv = process.env,
): Promise<string> {
  const list = [...new Set((typeof hosts === 'string' ? [hosts] : [...hosts]).map((h) => h.toLowerCase()))];
  const bad = list.find((h) => !isSharePointHost(h));
  if (list.length === 0 || bad !== undefined) {
    throw new Error(
      `refusing to capture a SharePoint token for ${JSON.stringify(bad ?? '')}: not a ` +
        '*.sharepoint.com host',
    );
  }
  const label = list.join(' or ');

  // `captureBearerHeader` flattens a capture failure into a plain Error, so the
  // refusal's type is noted here, where it is still a typed error.
  let refused = false;
  const scoped = bridgeScopedTransport(env, list);
  let token: string;
  try {
    token = await captureBearerHeader({
      // Any call the browser makes to one of these hosts. Concrete hosts, never
      // a wildcard — see the module note.
      captures: list.map(sharePointCaptureDecl),
      // The one bridge scope (`bridge-scope.ts`), with these hosts' captures
      // added to the remembered ones so the hello asks for exactly what it is
      // about to read.
      domains: SHAREPOINT_TRUST_DOMAINS,
      // Read per capture, not at import: see `capture-window.ts`.
      windowMs: captureWindowMs(env),
      port: getWsPort(),
      serverName: PACKAGE_NAME,
      version: VERSION,
      what: `a SharePoint token for ${label}`,
      guidance:
        `Reload your signed-in Teams tab and retry while it loads — that makes the ` +
        `page call ${label} with the SharePoint token we capture.`,
      createTransport: (opts) => noteScopeRefusal(scoped(opts), () => (refused = true)),
    });
  } catch (err) {
    if (!refused) throw err;
    throw new McpToolError(`The ContextMint Bridge has not yet allowed a SharePoint token capture from ${label}.`, {
      hint:
        `Open the ContextMint Bridge extension popup and approve the pending scope update for ${PACKAGE_NAME}, ` +
        `which adds ${label} (no need to remove or re-pair it). Then retry, and reload your Teams tab while it waits.`,
      cause: err,
    });
  }
  // Granted and captured: every later hello declares these hosts too.
  for (const h of list) rememberSharePointHost(h, env);
  return token;
}

/** Call `onRefused` when the bridge refuses a capture for being outside the granted scope. */
function noteScopeRefusal(transport: CaptureTransport, onRefused: () => void): CaptureTransport {
  const server = transport.server;
  const capture = server.captureRequestHeader.bind(server);
  return {
    start: () => transport.start(),
    close: () => transport.close(),
    server: {
      captureRequestHeader: async (...args: Parameters<typeof capture>) => {
        try {
          return await capture(...args);
        } catch (err) {
          // @fetchproxy/server's FetchproxyScopeError, matched by name: this
          // module must not import the optional bridge package for a class.
          if (err instanceof Error && err.name === 'FetchproxyScopeError') onRefused();
          throw err;
        }
      },
    },
  };
}
