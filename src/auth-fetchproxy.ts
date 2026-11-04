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
import { BRIDGE_TRUST_DOMAINS, CAPTURE_DECL, withBridgeScope } from './bridge-scope.js';
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
 * issue one, hence the guidance in the capture error below. Defined in
 * `bridge-scope.ts`, beside the other captures it is declared with.
 */
export { CAPTURE_DECL } from './bridge-scope.js';

/**
 * The domains every capture of this server trusts — see `bridge-scope.ts` for
 * why the chat-service capture declares the SharePoint and Substrate domains
 * too, and every capture header beside its own.
 */
export const TRUST_DOMAINS = [...BRIDGE_TRUST_DOMAINS];

/**
 * Build the capture transport through this module's import binding, so a
 * test's `vi.mock('@chrischall/mcp-utils/fetchproxy')` stub is the one used.
 * Shared by the SharePoint and Substrate captures.
 */
export const createCaptureTransport = (opts: Parameters<typeof createFetchproxyTransport>[0]): CaptureTransport =>
  createFetchproxyTransport<CaptureTransport>(opts);

/**
 * A capture transport factory that declares the one bridge scope
 * (`withBridgeScope`), whichever single header the capture reads.
 * `sharePointHost` is the host a SharePoint capture is about to try.
 */
export const bridgeScopedTransport =
  (env: NodeJS.ProcessEnv, sharePointHost?: string) =>
  (opts: Parameters<typeof createFetchproxyTransport>[0]): CaptureTransport =>
    createCaptureTransport(withBridgeScope(opts, env, sharePointHost));

/**
 * Capture a chat-service access token from the user's signed-in Teams tab.
 * Returns the bare token (no `Bearer ` prefix).
 *
 * `captureBearerHeader` declares only `capture_request_header` (beside the
 * `fetch` base verb mcp-utils always keeps), widened to the one bridge scope
 * (`bridge-scope.ts`) so this capture never re-pairs a grant the search or
 * SharePoint capture made — sets the bridge deadline to the capture window + 15 s, starts the transport
 * (skipping `start()` makes the first verb throw "ensureConnected called before
 * listen()", which reads exactly like a pairing problem), guards against a
 * header too short to be a token, and closes the transport whatever happened.
 */
export async function captureTokenViaFetchproxy(env: NodeJS.ProcessEnv = process.env): Promise<string> {
  return captureBearerHeader({
    captures: [{ ...CAPTURE_DECL }],
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
    createTransport: bridgeScopedTransport(env),
  });
}
