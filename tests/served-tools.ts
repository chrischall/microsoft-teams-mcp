// The full served tool surface, registered against a stub client, for tests
// that assert on what `tools/list` actually publishes.
import { createTestHarness } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import type { ChatServiceClient } from '../src/chatsvc.js';
import { registerChatTools } from '../src/tools/chat.js';
import { registerChannelTools } from '../src/tools/channels.js';
import { registerConversationTools } from '../src/tools/conversation.js';
import { registerUnreadTools } from '../src/tools/unread.js';
import { registerAttachmentTools } from '../src/tools/attachment.js';
import { registerSendTools } from '../src/tools/send.js';
import { registerReactTools } from '../src/tools/react.js';
import { registerActivityTools } from '../src/tools/activity.js';
import { registerHealthcheckTool } from '../src/tools/healthcheck.js';

export async function servedTools() {
  const deps = {
    chatsvc: { listConversations: async () => [], meMri: async () => '8:orgid:x' } as unknown as ChatServiceClient,
  };

  const h = await createTestHarness((server: McpServer) => {
    registerChatTools(server, deps);
    registerChannelTools(server, deps);
    registerConversationTools(server, deps);
    registerUnreadTools(server, deps);
    registerAttachmentTools(server, deps);
    registerSendTools(server, deps);
    registerReactTools(server, deps);
    registerActivityTools(server, deps);
    registerHealthcheckTool(server, deps);
  });
  try {
    return (await h.client.listTools()).tools;
  } finally {
    await h.close();
  }
}
