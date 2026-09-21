import { McpToolError } from '@chrischall/mcp-utils';

/** Classify a bridge-read failure into an actionable McpToolError. Shared by every tool. */
export function wrapBridgeError(err: unknown, action: string): McpToolError {
  const message = err instanceof Error ? err.message : String(err);
  if (/no tab matching|could not reach a signed-in/i.test(message)) {
    return new McpToolError(`Could not reach a signed-in Teams tab to ${action}.`, {
      hint: 'Open teams.cloud.microsoft in Chrome, make sure you are signed in, and retry.',
    });
  }
  if (/pair code|not granted|not in declared/i.test(message)) {
    return new McpToolError(`The Teams browser bridge is not yet approved: ${message}`, {
      hint: 'Approve the pairing request in the Transporter extension popup, then retry.',
    });
  }
  return new McpToolError(`Failed to ${action}: ${message}`);
}
