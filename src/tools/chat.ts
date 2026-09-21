/**
 * Chat tools — both read the user's live, signed-in `teams.microsoft.com`
 * tab through the fetchproxy bridge. See `client.ts`'s module doc for why
 * there is no way to select a specific chat: the bridge can only READ the
 * DOM, never navigate it, so `teams_get_open_chat_messages` returns whatever
 * conversation is currently displayed.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { minifiedResult, toolAnnotations, McpToolError } from '@chrischall/mcp-utils';
import type { TeamsClient } from '../client.js';

function wrapBridgeError(err: unknown, action: string): McpToolError {
  const message = err instanceof Error ? err.message : String(err);
  if (/no tab matching|could not reach a signed-in/i.test(message)) {
    return new McpToolError(`Could not reach a signed-in Teams tab to ${action}.`, {
      hint:
        'Open teams.microsoft.com (or teams.cloud.microsoft) in Chrome, make sure you are ' +
        'signed in, and retry.',
    });
  }
  if (/pair code|not granted|not in declared/i.test(message)) {
    return new McpToolError(`The Teams browser bridge is not yet approved: ${message}`, {
      hint: 'Approve the pairing request in the Transporter extension popup, then retry.',
    });
  }
  return new McpToolError(`Failed to ${action}: ${message}`);
}

export function registerChatTools(server: McpServer, client: TeamsClient): void {
  server.registerTool(
    'teams_list_chats',
    {
      title: 'List Teams chats',
      description:
        'List the chats in the signed-in user\'s Teams chat sidebar (1:1s, group chats, meeting ' +
        'chats), each with its title, last-message preview, last-activity time, and a ' +
        'conversationKey (Teams\' internal thread id — not usable to open the chat, there is no ' +
        'navigate capability). Reads the currently-rendered list from the user\'s live browser tab.',
      annotations: toolAnnotations({ readOnly: true }),
      inputSchema: {},
    },
    async () => {
      try {
        const rows = await client.listChats();
        return minifiedResult(rows);
      } catch (err) {
        throw wrapBridgeError(err, 'list chats');
      }
    },
  );

  server.registerTool(
    'teams_get_open_chat_messages',
    {
      title: 'Read the currently open Teams chat',
      description:
        'Read the message thread of whichever Teams chat is CURRENTLY OPEN in the user\'s ' +
        'signed-in browser tab — sender, ISO-8601 time, message id, and text for each message. ' +
        'There is no way to choose a different chat from here (the bridge can only read the DOM, ' +
        'not navigate it): ask the user to open the chat they mean first, or use ' +
        'teams_list_chats to show them what is available. A quoted-reply message\'s text may ' +
        'include the quoted preview concatenated in.',
      annotations: toolAnnotations({ readOnly: true }),
      inputSchema: {},
    },
    async () => {
      try {
        const rows = await client.getOpenChatMessages();
        return minifiedResult(rows);
      } catch (err) {
        throw wrapBridgeError(err, 'read the open chat');
      }
    },
  );
}
