import { McpToolError } from '@chrischall/mcp-utils';

/**
 * Classify a chat-service failure. The client already raises actionable
 * `McpToolError`s (401 after re-capture, 429, 5xx), which pass through as-is;
 * what is left is the token capture, which fails as a plain `Error` from
 * `auth-fetchproxy.ts`.
 */
export function wrapApiError(err: unknown, action: string): McpToolError {
  if (err instanceof McpToolError) return err;
  const message = err instanceof Error ? err.message : String(err);
  // A pairing prompt surfaces inside the capture error; approving it is the
  // fix, not opening a tab.
  if (/pair code|not granted|not in declared/i.test(message)) {
    return new McpToolError(`The Teams browser bridge is not yet approved: ${message}`, {
      hint: 'Approve the pairing request in the ContextMint Bridge extension popup, then retry.',
    });
  }
  if (/could not capture a Teams token/i.test(message)) {
    return new McpToolError(`Could not get a Teams access token to ${action}.`, {
      hint:
        'Open a signed-in teams.cloud.microsoft tab and click into any chat (that is what makes ' +
        'Teams call its chat service, which is where the token is read from), then retry.',
      cause: err,
    });
  }
  return new McpToolError(`Failed to ${action}: ${message}`);
}
