// The full served tool surface, registered against a stub client, for tests
// that assert on what `tools/list` actually publishes.
import { createTestHarness } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import type { TeamsClient } from '../src/client.js';
import { registerChatTools } from '../src/tools/chat.js';
import { registerChannelTools } from '../src/tools/channels.js';
import { registerActivityTools } from '../src/tools/activity.js';
import { registerHealthcheckTool } from '../src/tools/healthcheck.js';

export async function servedTools() {
  const client = {
    listChats: async () => [],
    getOpenChatMessages: async () => [],
    listTeamsAndChannels: async () => [],
    getOpenChannelPosts: async () => [],
    getActivity: async () => [],
    transport: { status: () => ({}), runProbe: async () => ({ ok: true, elapsed_ms: 0, bridge: {} }) },
  } as unknown as TeamsClient;

  const h = await createTestHarness((server: McpServer) => {
    registerChatTools(server, client);
    registerChannelTools(server, client);
    registerActivityTools(server, client);
    registerHealthcheckTool(server, client);
  });
  try {
    return (await h.client.listTools()).tools;
  } finally {
    await h.close();
  }
}
