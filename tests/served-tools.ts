// The full served tool surface, registered against a stub client, for tests
// that assert on what `tools/list` actually publishes.
import { createTestHarness } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import type { TeamsClient } from '../src/client.js';
import type { ChatServiceClient } from '../src/chatsvc.js';
import { registerChatTools } from '../src/tools/chat.js';
import { registerChannelTools } from '../src/tools/channels.js';
import { registerActivityTools } from '../src/tools/activity.js';
import { registerHealthcheckTool } from '../src/tools/healthcheck.js';

export async function servedTools() {
  const deps = {
    teams: {
      getActivity: async () => [],
      transport: { status: () => ({}), runProbe: async () => ({ ok: true, elapsed_ms: 0, bridge: {} }) },
    } as unknown as TeamsClient,
    chatsvc: { listConversations: async () => [], meMri: async () => '8:orgid:x' } as unknown as ChatServiceClient,
  };

  const h = await createTestHarness((server: McpServer) => {
    registerChatTools(server, deps);
    registerChannelTools(server, deps);
    registerActivityTools(server, deps);
    registerHealthcheckTool(server, deps);
  });
  try {
    return (await h.client.listTools()).tools;
  } finally {
    await h.close();
  }
}
