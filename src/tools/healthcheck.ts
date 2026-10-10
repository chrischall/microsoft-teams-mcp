/**
 * Healthcheck. Chats and channels are read through the chat-service API, so
 * the probe is one conversations listing — which also exercises the token
 * capture on a cold start. The bridge status section still matters: the token
 * is captured through it, and the Activity feed reads the page through it.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { registerBridgeHealthcheckTool } from '@chrischall/mcp-utils/fetchproxy';
import type { TeamsDeps } from '../deps.js';

export function registerHealthcheckTool(server: McpServer, { teams, chatsvc }: TeamsDeps): void {
  registerBridgeHealthcheckTool({
    server,
    prefix: 'teams',
    hostLabel: 'teams.cloud.microsoft',
    transport: teams.transport,
    probePath: 'chat-service conversation listing (GET /users/ME/conversations)',
    probeFn: async () => {
      const conversations = await chatsvc.listConversations();
      return `${conversations.length} conversation(s) listed`;
    },
    classifyThrown: (err) =>
      /could not capture a Teams token/i.test(err instanceof Error ? err.message : String(err))
        ? { kind: 'token_capture', hint: TOKEN_CAPTURE_HINT }
        : undefined,
  });
}

const TOKEN_CAPTURE_HINT =
  'No Teams access token could be captured. Open a signed-in teams.cloud.microsoft tab and click ' +
  'into any chat (that is what makes Teams call its chat service, where the token is read from), ' +
  'then retry.';
