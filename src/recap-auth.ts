/**
 * Bootstrap-only fetchproxy auth for meeting recaps — the Meeting Collab
 * Platform Service (`teams.cloud.microsoft/api/mcps/*`) that Teams' Recap tab
 * reads (`teams_get_meeting_recap`).
 *
 * MCPS accepts neither the chat-service token (a 401) nor the Substrate one: it
 * wants its own (audience `6bc3b958-689b-49f5-9006-36d165f30e00`, scopes
 * `User.Read user_impersonation`, about an hour). So, as with the other
 * three, we snapshot the `Authorization` header off a request the signed-in
 * Teams tab already makes — here, the `readcollabobject` call its Recap tab
 * issues on opening — and talk to MCPS with a plain Node `fetch` (`recap.ts`).
 *
 * Verified live 2026-11-10: opening a meeting chat's Recap tab during the
 * capture window yielded that token, which `recap.ts` replayed from Node.
 *
 * Import LAZILY (`await import('./recap-auth.js')`), like the other captures,
 * so boot never loads the bridge.
 */
import { captureBearerHeader } from '@chrischall/mcp-utils/fetchproxy';
import { captureWindowMs } from './capture-window.js';
import { bridgeScopedTransport, getWsPort } from './auth-fetchproxy.js';
import { BRIDGE_TRUST_DOMAINS, RECAP_CAPTURE_DECL } from './bridge-scope.js';
import { PACKAGE_NAME, VERSION } from './version.js';

export { RECAP_CAPTURE_DECL } from './bridge-scope.js';

/**
 * Capture a meeting-recap (MCPS) token from the user's signed-in Teams tab.
 * Returns the bare token (no `Bearer ` prefix).
 */
export async function captureRecapTokenViaFetchproxy(env: NodeJS.ProcessEnv = process.env): Promise<string> {
  return captureBearerHeader({
    // The one bridge scope every capture of this server declares (`bridge-scope.ts`).
    captures: [{ ...RECAP_CAPTURE_DECL }],
    domains: [...BRIDGE_TRUST_DOMAINS],
    windowMs: captureWindowMs(env),
    port: getWsPort(),
    serverName: PACKAGE_NAME,
    version: VERSION,
    what: 'a Teams meeting-recap token',
    guidance:
      "In a signed-in teams.cloud.microsoft tab, open any meeting chat and click its Recap tab — that is what " +
      'makes Teams call its meeting-recap service, where the token is read from — then retry. If the ContextMint ' +
      'Bridge asks to re-pair microsoft-teams-mcp, approve it once: the recap capture is new to its scope.',
    createTransport: bridgeScopedTransport(env),
  });
}
