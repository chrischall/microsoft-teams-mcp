/**
 * `teams_list_chats` — every chat the user is in (group, meeting, 1:1), read
 * from the Teams chat-service API. Unlike the old sidebar read it needs no
 * particular Teams view open, and it is not limited to what the sidebar has
 * rendered: the listing is the user's whole conversation set.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { toolAnnotations } from '@chrischall/mcp-utils';
import type { TeamsDeps } from '../deps.js';
import type { ConversationSummary } from '../model.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';
import { limitArg, trimRows } from './limit.js';
import { listSummaries, queryArg, unreadOnlyArg } from './listing.js';

/** Up to 10 pages of 200 come back; this bounds one result, not the listing. */
const MAX_CHATS = 500;

/** Newest activity first; a chat with no messages at all sorts last. */
function byLastActivityDesc(a: ConversationSummary, b: ConversationSummary): number {
  return (b.lastActivity ?? '').localeCompare(a.lastActivity ?? '');
}

export function registerChatTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_list_chats',
    {
      title: 'List Teams chats',
      description:
        'List every chat the signed-in user is in — 1:1s, group chats and meeting chats, not ' +
        'just the ones visible in the Teams sidebar — most recent activity first. Each has an ' +
        'id, kind (oneOnOne / group / meeting, or self for the user\'s own 48:notes self-chat), ' +
        'title, unread flag, last-activity time (ISO 8601), and a last-message preview and sender. ' +
        'Pass an id (or the title) to teams_read_conversation to read the messages. Use ' +
        'unread_only to find chats with messages the user has not read — as in Teams, a call ' +
        'recording, transcript or other system event never makes a chat unread; listing never ' +
        'marks anything read. Channels are listed by teams_list_teams_and_channels instead.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        unread_only: unreadOnlyArg,
        query: queryArg('chat title'),
        limit: limitArg(MAX_CHATS, 'chats', 'most recent activity first kept'),
      },
    },
    async ({ unread_only, query, limit }) => {
      try {
        const rows = await listSummaries(chatsvc, { kinds: ['group', 'meeting', 'oneOnOne', 'self'], unread_only, query });
        rows.sort(byLastActivityDesc);
        const { rows: chats, truncated } = trimRows(rows, limit, 'start', MAX_CHATS, 'most recent');
        return untrustedResult({ chats, ...(truncated ? { truncated } : {}) });
      } catch (err) {
        throw wrapApiError(err, 'list chats');
      }
    },
  );
}
