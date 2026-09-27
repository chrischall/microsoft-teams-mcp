import { McpToolError } from '@chrischall/mcp-utils';
import { FetchproxyCapabilityUnavailableError } from '@fetchproxy/server';

/** Classify a bridge-read failure into an actionable McpToolError. Shared by every tool. */
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
