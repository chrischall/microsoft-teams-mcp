/**
 * `teams_list_chats` — every chat the user is in (group, meeting, 1:1), read
 * from the Teams chat-service API. Unlike the old sidebar read it needs no
 * particular Teams view open, and it is not limited to what the sidebar has
 * rendered: the listing is the user's whole conversation set.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { McpToolError, toolAnnotations } from '@chrischall/mcp-utils';
import { z } from 'zod';
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
  type ListedRow,
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

/** Alphabetical by title, ignoring case and accents; a stable sort keeps recency among equals. */
function byTitle(a: ConversationSummary, b: ConversationSummary): number {
  return a.title.localeCompare(b.title, undefined, { sensitivity: 'base' });
}

/** An optional ISO 8601 date-time bound on a chat's last activity. */
function activityBoundArg(name: 'since' | 'until', side: string) {
  return z
    .string()
    .refine((s) => !Number.isNaN(Date.parse(s)), { message: `${name} must be an ISO 8601 date-time` })
    .optional()
    .describe(
      `Only return chats whose last activity is ${side} this ISO 8601 date-time (e.g. 2026-09-21T09:00:00Z). ` +
        'A chat with no last activity is left out whenever since or until is given.',
    );
}

const SORTS = ['activity', 'activity_only', 'title'] as const;
type ChatSort = (typeof SORTS)[number];

const sortArg = z
  .enum(SORTS)
  .optional()
  .describe(
    'Order of the result, and so which chats limit keeps: "activity" (default) — favorites, then pinned, ' +
      'then most recent activity (unread_only: most recent activity); "activity_only" — most recent ' +
      'activity first, ignoring favorites and pinned; "title" — alphabetical by title, case-insensitive ' +
      '(titles every untitled chat in the listing first, so it reads more).',
  );

/** Orders `rows` in place for `sort`; titles must already be final for "title". */
function sortChats(rows: ListedRow[], sort: ChatSort, unreadOnly: boolean | undefined): void {
  rows.sort(byLastActivityDesc);
  if (sort === 'title') rows.sort(byTitle);
  else if (sort === 'activity' && !unreadOnly) rows.sort(favoritesFirst);
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
        ' since / until (ISO 8601 date-times) keep chats whose last activity falls in that window; ' +
        'sort picks the order: activity (default, as above), activity_only (most recent first, ' +
        'ignoring favorites and pinned) or title (alphabetical, case-insensitive).' +
        ' Listing never marks anything read. Channels are listed by teams_list_teams_and_channels instead.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        unread_only: unreadOnlyArg,
        since_days: listSinceDaysArg,
        query: queryArg('chat title or an untitled chat\'s participant names'),
        since: activityBoundArg('since', 'at or after'),
        until: activityBoundArg('until', 'at or before'),
        sort: sortArg,
        limit: limitArg(MAX_CHATS, 'chats', 'the head of the sort order kept'),
      },
    },
    async ({ unread_only, since_days, query, since, until, sort = 'activity', limit }) => {
      const activeSince = since !== undefined ? Date.parse(since) : undefined;
      const activeUntil = until !== undefined ? Date.parse(until) : undefined;
      if (activeSince !== undefined && activeUntil !== undefined && activeSince > activeUntil) {
        throw new McpToolError(`since (${since}) is after until (${until}).`, {
          hint: 'Pass a since at or before until.',
        });
      }
      try {
        const { rows, retitle, errors: _untitledErrors, ...unreadInfo } = await listSummaries(chatsvc, {
          kinds: ['group', 'meeting', 'oneOnOne', 'self'],
          unread_only,
          since_days,
          query,
          activeSince,
          activeUntil,
        });
        // A title sort needs every untitled chat's final title before limit picks the head;
        // otherwise only the kept rows are titled. errors names conversations too, so it is
        // retitled alongside the rows.
        const titledFirst = sort === 'title' ? await retitle(rows) : undefined;
        const ordered = titledFirst?.rows ?? rows;
        sortChats(ordered, sort, unread_only);
        const which = sort === 'title' || (sort === 'activity' && !unread_only) ? 'first' : 'most recent';
        const { rows: kept, truncated } = trimRows(ordered, limit, 'start', MAX_CHATS, which);
        const { rows: chats, errors } = titledFirst ? { rows: kept, errors: titledFirst.errors } : await retitle(kept);
        return untrustedResult({ chats, ...unreadInfo, ...(errors ? { errors } : {}), ...(truncated ? { truncated } : {}) });
      } catch (err) {
        throw wrapApiError(err, 'list chats');
      }
    },
  );
}
