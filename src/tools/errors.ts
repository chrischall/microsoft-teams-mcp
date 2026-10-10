import { McpToolError } from '@chrischall/mcp-utils';
import { FetchproxyCapabilityUnavailableError } from '@fetchproxy/server';

/** Classify a bridge-read failure (the DOM-read Activity feed) into an actionable McpToolError. */
export function wrapBridgeError(err: unknown, action: string): McpToolError {
  // A browser limitation (e.g. Safari lacks an API the verb needs) — not a
  // pairing problem and not a bug in this MCP, so it gets its own message.
  if (err instanceof FetchproxyCapabilityUnavailableError) {
    const where = err.platform ? `this browser (${err.platform})` : 'this browser';
    return new McpToolError(
      `Cannot ${action}: ${where} cannot serve the "${err.capability}" capability.`,
      {
        hint: 'Nothing is wrong with the pairing or this MCP. Use Teams in a browser that supports it (for example Chrome) with ContextMint Bridge installed.',
        cause: err,
      },
    );
  }
  const message = err instanceof Error ? err.message : String(err);
  if (/no tab matching|could not reach a signed-in/i.test(message)) {
    return new McpToolError(`Could not reach a signed-in Teams tab to ${action}.`, {
      hint: 'Open teams.cloud.microsoft in Chrome, make sure you are signed in, and retry.',
    });
  }
  if (/pair code|not granted|not in declared/i.test(message)) {
    return new McpToolError(`The Teams browser bridge is not yet approved: ${message}`, {
      hint: 'Approve the pairing request in the ContextMint Bridge extension popup, then retry.',
    });
  }
  return new McpToolError(`Failed to ${action}: ${message}`);
}

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
