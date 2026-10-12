/**
 * Bootstrap-only fetchproxy auth for Microsoft Substrate search — the backend
 * the Teams web client's own search box uses (`teams_search_messages`).
 *
 * Substrate does not accept the chat-service token (its audience is
 * `https://ic3.teams.office.com`), so search needs a THIRD token, captured the
 * same way the other two are: we snapshot the `Authorization: Bearer …` header
 * off a request the signed-in Teams tab makes to
 * `substrate.office.com/searchservice/*` — which it does when the user runs a
 * search in Teams' search box — and then talk to Substrate with a plain Node
 * `fetch` (`substrate.ts`). Requests only: nothing here, or anywhere in this
 * MCP, reads a token out of page storage.
 *
 * The capture path is restricted to `/searchservice/*` rather than `/*`: other
 * Substrate APIs the page calls are minted for other scopes, and a token lifted
 * off one of those could fail on the search API in ways that look like a
 * Substrate outage.
 *
 * NOT YET VERIFIED LIVE (2026-10-12): the capture attempt never reached the
 * extension — the fetchproxy server refused the answering extension because its
 * identity did not match the one `microsoft-teams-mcp` had pinned (re-pairing
 * needs `FETCHPROXY_TRUST_NEW_EXTENSION=1`, the user's call). The search
 * request itself was seen live (`POST /searchservice/api/v2/query`), so the
 * host and path are known; the token's audience and lifetime are not.
 *
 * Import this module LAZILY (`await import('./substrate-auth.js')`), as
 * `sharepoint-auth.ts` is, so a server run with `TEAMS_DISABLE_FETCHPROXY`
 * never has to load the bridge to boot.
 */
import { createBootstrapOpts, createFetchproxyTransport } from '@chrischall/mcp-utils/fetchproxy';
import { bridgeDeadlineMs, captureWindowMs } from './capture-window.js';
import { getWsPort, stripBearer } from './auth-fetchproxy.js';
import { PACKAGE_NAME, VERSION } from './version.js';

/** The one request we capture from: a search-service call the Teams tab makes. */
export const SUBSTRATE_CAPTURE_DECL = {
  host: 'substrate.office.com',
  path: '/searchservice/*',
  headerName: 'authorization',
} as const;

/**
 * The extension's trust boundary for search: exactly the host we capture from.
 * A third, independent scope beside `teams.cloud.microsoft` and
 * `sharepoint.com` — granting it takes a one-time re-pair of
 * `microsoft-teams-mcp` in the ContextMint Bridge popup; until then search falls
 * back to the client-side scan.
 */
export const SUBSTRATE_TRUST_DOMAINS = [SUBSTRATE_CAPTURE_DECL.host] as const;

/** Shortest plausible token; guards against an empty or echoed header only. */
const MIN_TOKEN_LENGTH = 100;

/**
 * Capture a Substrate search token from the user's signed-in Teams tab.
 * Returns the bare token (no `Bearer ` prefix).
 */
export async function captureSubstrateTokenViaFetchproxy(env: NodeJS.ProcessEnv = process.env): Promise<string> {
  // Read per capture, not at import: see `capture-window.ts`.
  const windowMs = captureWindowMs(env);
  const transport = createFetchproxyTransport({
    // Same verb set as the other two captures (`capture_request_header` beside
    // the always-on `fetch`); only the trusted domain differs.
    ...createBootstrapOpts({
      domains: [...SUBSTRATE_TRUST_DOMAINS],
      bootstrap: { captureHeaders: [{ ...SUBSTRATE_CAPTURE_DECL }] },
    }),
    port: getWsPort(),
    fetchTimeoutMs: bridgeDeadlineMs(windowMs),
    captureWindowMs: windowMs,
    serverName: PACKAGE_NAME,
    version: VERSION,
  });

  try {
    // Loads the identity keypair; skipping it makes the first verb throw
    // "ensureConnected called before listen()" (see `auth-fetchproxy.ts`).
    await transport.start();

    try {
      const raw = await transport.server.captureRequestHeader({ ...SUBSTRATE_CAPTURE_DECL, timeoutMs: windowMs });
      if (typeof raw !== 'string') {
        throw new Error(`${SUBSTRATE_CAPTURE_DECL.host}: capture returned no header`);
      }
      const token = stripBearer(raw);
      if (token.length < MIN_TOKEN_LENGTH) {
        throw new Error(
          `${SUBSTRATE_CAPTURE_DECL.host}: authorization header was present but too short to be a ` +
            `token (${token.length} chars)`,
        );
      }
      return token;
    } catch (e) {
      const detail = (e as Error).message;
      throw new Error(
        `could not capture a Substrate search token from the browser (${detail}). ` +
          'In a signed-in teams.cloud.microsoft tab, type anything into the search box at the top ' +
          'and press Enter — that is what makes Teams call substrate.office.com, where the token ' +
          'is read from — then retry. A "not in declared set" error means the ContextMint Bridge ' +
          'pairing does not trust substrate.office.com yet: re-pair microsoft-teams-mcp once.',
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
