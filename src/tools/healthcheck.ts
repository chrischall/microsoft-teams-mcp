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
      return `${rows.length} chat(s) visible in the sidebar`;
    },
  });
}
