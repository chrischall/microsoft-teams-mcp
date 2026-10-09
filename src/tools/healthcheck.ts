/**
 * Bridge healthcheck. Every real tool rides the fetchproxy bridge — there is
 * no credential to check independently of it — so this probes the same way
 * `teams_list_chats` does rather than a separate HTTP round-trip.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { registerBridgeHealthcheckTool } from '@chrischall/mcp-utils/fetchproxy';
import type { TeamsClient } from '../client.js';

export function registerHealthcheckTool(server: McpServer, client: TeamsClient): void {
  registerBridgeHealthcheckTool({
    server,
    prefix: 'teams',
    hostLabel: 'teams.cloud.microsoft',
    transport: client.transport,
    probePath: 'chat list (read_dom_list against the currently open Teams tab)',
    probeFn: async () => {
      const rows = await client.listChats();
      // The chat list renders only while the Chat view is open, and on any
      // other view (Teams, Activity, Calendar…) the read returns [] rather
      // than throwing — so an empty read means "wrong view", not "healthy".
      if (rows.length === 0) throw new ChatViewNotOpenError();
      return `${rows.length} chat(s) visible in the sidebar`;
    },
    classifyThrown: (err) =>
      err instanceof ChatViewNotOpenError ? { kind: 'chat_view_not_open', hint: CHAT_VIEW_NOT_OPEN_HINT } : undefined,
  });
}

const CHAT_VIEW_NOT_OPEN_HINT =
  'The bridge reached the teams.cloud.microsoft tab, but no chat list is rendered: the Chat view ' +
  'is not open (the tab is probably on Teams, Activity or Calendar). Ask the user to click Chat ' +
  'in the left nav, then retry.';

class ChatViewNotOpenError extends Error {
  constructor() {
    super('bridge OK but no chat list rendered — is the Chat view open?');
    this.name = 'ChatViewNotOpenError';
  }
}
