/**
 * Chat tools — both read the user's live, signed-in `teams.cloud.microsoft`
 * tab through the fetchproxy bridge. See `client.ts`'s module doc for why
 * there is no way to select a specific chat: the bridge can only READ the
 * DOM, never navigate it, so `teams_get_open_chat_messages` returns whatever
 * conversation is currently displayed.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { toolAnnotations } from '@chrischall/mcp-utils';
import type { TeamsClient } from '../client.js';
import { wrapBridgeError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

export function registerChatTools(server: McpServer, client: TeamsClient): void {
  server.registerTool(
    'teams_list_chats',
    {
      title: 'List Teams chats',
      description:
        'List the chats in the signed-in user\'s Teams chat sidebar (1:1s, group chats, meeting ' +
        'chats), each with its title, last-message preview, last-activity time, and a ' +
        'conversationKey (Teams\' internal thread id — not usable to open the chat, there is no ' +
        'navigate capability). Reads the currently-rendered list from the user\'s live browser tab.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true }),
      inputSchema: {},
    },
    async () => {
      try {
        const rows = await client.listChats();
        return untrustedResult({ rows });
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
        'include the quoted preview concatenated in.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true }),
      inputSchema: {},
    },
    async () => {
      try {
        const rows = await client.getOpenChatMessages();
        return untrustedResult({ rows });
      } catch (err) {
        throw wrapBridgeError(err, 'read the open chat');
      }
    },
  );
}
