/**
 * Bootstrap-only fetchproxy auth for the Teams chat-service API.
 *
 * The Teams web worker sends `Authorization: Bearer …` on every request to
 * `https://teams.cloud.microsoft/api/chatsvc/*`. We snapshot that header off a
 * request the page already makes, and from then on this MCP talks to the
 * chat service with a plain Node `fetch`.
 *
 * The bridge is therefore touched ONCE per token (~23h), never in the hot
 * path — and nothing ever reads the page itself: the Activity feed is a
 * chat-service conversation too (`tools/activity.ts`).
 *
 * Verified live 2026-10-10 against a real Teams tenant:
 *   - the token has `aud: https://ic3.teams.office.com`,
 *     `scp: Teams.AccessAsUser.All` and a ~23.5h lifetime;
 *   - that token plus `behavioroverride: redirectAs404` authenticates plain
 *     server-side requests to the chat service with no bridge involved;
 *   - no `authentication`/`x-skypetoken` header is needed.
 *
 * We do NOT try to mint a token by calling an endpoint ourselves — it comes
 * from MSAL's hidden-iframe flow, which a bridge fetch cannot drive. Capture is
 * the only path, which is why a failed capture is a hard error with actionable
 * guidance rather than a fallback.
 *
 * Import this module LAZILY (`await import('./auth-fetchproxy.js')`) from the
 * code that needs a token, as Outlook's client does, so a server run with
 * `TEAMS_ACCESS_TOKEN` never has to load the bridge to boot.
 */
import { captureBearerHeader, createFetchproxyTransport, type CaptureTransport } from '@chrischall/mcp-utils/fetchproxy';
import { readPortEnv } from '@chrischall/mcp-utils';
import { captureWindowMs } from './capture-window.js';
import { PACKAGE_NAME, VERSION } from './version.js';

export { stripBearer } from '@chrischall/mcp-utils/fetchproxy';

/**
 * fetchproxy concentrator port — one for the whole fleet. `TEAMS_WS_PORT`
 * overrides it for local development and for a hosted bridged registration
 * (where mcp-host names this variable in `bridgePortEnv`).
 */
const DEFAULT_WS_PORT = 37_149;
export function getWsPort(): number {
  return readPortEnv('TEAMS_WS_PORT', DEFAULT_WS_PORT);
}

/**
 * The one request we capture from: any chat-service call on the New Teams
 * host. An idle tab may make none — opening a chat is what makes the worker
 * issue one, hence the guidance in the capture error below.
 */
export const CAPTURE_DECL = {
  host: 'teams.cloud.microsoft',
  path: '/api/chatsvc/*',
  headerName: 'authorization',
} as const;

/**
 * The extension's trust boundary for the chat-service path: exactly the host we
 * capture from. Every API call afterwards is a plain server-side `fetch` the
 * bridge never sees, so nothing wider is needed here.
 *
 * Shared-file download adds a SECOND, independent trust scope —
 * `SHAREPOINT_TRUST_DOMAINS = ['sharepoint.com']` in `sharepoint-auth.ts` — used
 * only when a file is opened. It is additive: the SharePoint capture declares
 * `sharepoint.com` on its own transport and does NOT widen this chat-service
 * capture decl, which stays `teams.cloud.microsoft` alone. Granting the broader
 * SharePoint scope takes a one-time re-pair of `microsoft-teams-mcp` in the
 * ContextMint Bridge popup; until then shared files degrade to metadata + link.
 */
export const TRUST_DOMAINS = [CAPTURE_DECL.host];

/**
 * Build the capture transport through this module's import binding, so a
 * test's `vi.mock('@chrischall/mcp-utils/fetchproxy')` stub is the one used.
 * Shared by the SharePoint and Substrate captures.
 */
export const createCaptureTransport = (opts: Parameters<typeof createFetchproxyTransport>[0]): CaptureTransport =>
  createFetchproxyTransport<CaptureTransport>(opts);

/**
 * Capture a chat-service access token from the user's signed-in Teams tab.
 * Returns the bare token (no `Bearer ` prefix).
 *
 * `captureBearerHeader` declares only `capture_request_header` (beside the
 * `fetch` base verb mcp-utils always keeps) — the same set as before the
 * Activity feed left the page, so an existing pairing needs no re-approval —
 * sets the bridge deadline to the capture window + 15 s, starts the transport
 * (skipping `start()` makes the first verb throw "ensureConnected called before
 * listen()", which reads exactly like a pairing problem), guards against a
 * header too short to be a token, and closes the transport whatever happened.
 */
export async function captureTokenViaFetchproxy(env: NodeJS.ProcessEnv = process.env): Promise<string> {
  return captureBearerHeader({
    captures: [CAPTURE_DECL],
    domains: TRUST_DOMAINS,
    // Read per capture, not at import: see `capture-window.ts`.
    windowMs: captureWindowMs(env),
    port: getWsPort(),
    serverName: PACKAGE_NAME,
    version: VERSION,
    what: 'a Teams token',
    guidance:
      'Open a signed-in teams.cloud.microsoft tab and click into any chat — ' +
      'that is what makes the page call the chat service; an idle tab never ' +
      'does — then retry.',
    createTransport: createCaptureTransport,
  });
}
