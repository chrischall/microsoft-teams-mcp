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
import {
  FAVORITES_DESCRIPTION,
  favoritesFirst,
  GROUP_TITLE_DESCRIPTION,
  listSinceDaysArg,
  listSummaries,
  queryArg,
  UNREAD_HINT_DESCRIPTION,
  unreadOnlyArg,
} from './listing.js';

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
        'just the ones visible in the Teams sidebar — favorites and pinned first, then most recent activity. Each has an ' +
        'id, kind (oneOnOne / group / meeting, or self for the user\'s own 48:notes self-chat), ' +
        'title, unread hint, last-activity time (ISO 8601), and a last-message preview and sender. ' +
        GROUP_TITLE_DESCRIPTION +
        ' query also matches those participant names. ' +
        FAVORITES_DESCRIPTION +
        ' ' +
        'Pass an id (or the title) to teams_read_conversation to read the messages. Use ' +
        'unread_only to find chats with messages the user has not read, exactly as Teams counts ' +
        'them: each candidate is confirmed by reading its messages past the read marker, so a ' +
        'call recording, transcript or other system event never makes a chat unread. By default ' +
        'unread_only keeps only chats whose newest unread message is from the last 14 days ' +
        '(since_days widens it); older ones are counted in older_unread_skipped, not listed. ' +
        UNREAD_HINT_DESCRIPTION +
        ' Listing never marks anything read. Channels are listed by teams_list_teams_and_channels instead.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        unread_only: unreadOnlyArg,
        since_days: listSinceDaysArg,
        query: queryArg('chat title or an untitled chat\'s participant names'),
        limit: limitArg(MAX_CHATS, 'chats', 'favorites and pinned first, then most recent activity, kept'),
      },
    },
    async ({ unread_only, since_days, query, limit }) => {
      try {
        const { rows, retitle, errors: _untitledErrors, ...unreadInfo } = await listSummaries(chatsvc, {
          kinds: ['group', 'meeting', 'oneOnOne', 'self'],
          unread_only,
          since_days,
          query,
        });
        rows.sort(byLastActivityDesc);
        if (!unread_only) rows.sort(favoritesFirst);
        const { rows: kept, truncated } = trimRows(rows, limit, 'start', MAX_CHATS, unread_only ? 'most recent' : 'first');
        // errors names conversations too, so it is retitled alongside the rows.
        const { rows: chats, errors } = await retitle(kept);
        return untrustedResult({ chats, ...unreadInfo, ...(errors ? { errors } : {}), ...(truncated ? { truncated } : {}) });
      } catch (err) {
        throw wrapApiError(err, 'list chats');
      }
    },
  );
}
