/**
 * How long a token capture may wait. (`captureBearerHeader` derives the
 * bridge deadline that has to outlast it: the window + 15 s, never under the
 * library's 30 s.)
 *
 * Kept free of any bridge import, so it can be read and tested without
 * loading `auth-fetchproxy.ts` (the one transport this MCP builds) and its
 * optional `@fetchproxy/server` peer. It is read at CALL time, not module
 * load, so `TEAMS_CAPTURE_TIMEOUT` set by the host's env (or a `.env` loaded
 * after import) is honoured.
 */
import { readTtlMsEnv } from '@chrischall/mcp-utils';

/**
 * The capture window when `TEAMS_CAPTURE_TIMEOUT` is unset, unparseable or 0.
 *
 * The window IS the mechanism: capture resolves on the NEXT matching request,
 * so nothing the page did before the listener came up counts. 30s is long
 * enough for the user to click into a chat, short enough that an idle tab does
 * not hang a tool call for minutes.
 */
export const DEFAULT_CAPTURE_WINDOW_MS = 30_000;

/** `TEAMS_CAPTURE_TIMEOUT` (seconds) as milliseconds. */
export function captureWindowMs(env: NodeJS.ProcessEnv = process.env): number {
  const ms = readTtlMsEnv('TEAMS_CAPTURE_TIMEOUT', DEFAULT_CAPTURE_WINDOW_MS, { env });
  return ms > 0 ? ms : DEFAULT_CAPTURE_WINDOW_MS;
}

