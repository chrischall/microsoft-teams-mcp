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
 * Verified live 2026-10-12: after a one-time pair approval in the ContextMint
 * Bridge popup (the `substrate.office.com` scope), running any search in the
 * Teams search box during the capture window yielded a token off
 * `/searchservice/*` (an app-GUID audience, ~1.4h lifetime) that
 * `substrate.ts` replayed successfully from Node. Until the pair is approved,
 * the capture fails fast with "pairing required" and search falls back to the
 * scan.
 *
 * Import this module LAZILY (`await import('./substrate-auth.js')`), as
 * `sharepoint-auth.ts` is, so a server run with `TEAMS_DISABLE_FETCHPROXY`
 * never has to load the bridge to boot.
 */
import { captureBearerHeader } from '@chrischall/mcp-utils/fetchproxy';
import { captureWindowMs } from './capture-window.js';
import { createCaptureTransport, getWsPort } from './auth-fetchproxy.js';
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

/**
 * Capture a Substrate search token from the user's signed-in Teams tab.
 * Returns the bare token (no `Bearer ` prefix).
 */
export async function captureSubstrateTokenViaFetchproxy(env: NodeJS.ProcessEnv = process.env): Promise<string> {
  return captureBearerHeader({
    // Same verb set as the other two captures (`capture_request_header` beside
    // the always-on `fetch`); only the trusted domain differs.
    captures: [SUBSTRATE_CAPTURE_DECL],
    domains: SUBSTRATE_TRUST_DOMAINS,
    // Read per capture, not at import: see `capture-window.ts`.
    windowMs: captureWindowMs(env),
    port: getWsPort(),
    serverName: PACKAGE_NAME,
    version: VERSION,
    what: 'a Substrate search token',
    guidance:
      'In a signed-in teams.cloud.microsoft tab, type anything into the search box at the top ' +
      'and press Enter — that is what makes Teams call substrate.office.com, where the token ' +
      'is read from — then retry. A "not in declared set" error means the ContextMint Bridge ' +
      'pairing does not trust substrate.office.com yet: re-pair microsoft-teams-mcp once.',
    createTransport: createCaptureTransport,
  });
}
