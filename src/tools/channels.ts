/**
 * `teams_list_teams_and_channels` — every channel the user is in, read from
 * the Teams chat-service API, so no Teams view has to be open. A channel's
 * team is named only when the team's own thread is in the listing (see
 * `conversationSummary`); otherwise `team` is absent and it sorts last.
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
  includeHiddenArg,
  listSinceDaysArg,
  listSummaries,
  queryArg,
  UNREAD_HINT_DESCRIPTION,
  unreadOnlyArg,
} from './listing.js';

/** Up to 10 pages of 200 come back; this bounds one result, not the listing. */
const MAX_CHANNELS = 500;

/**
 * Visible channels first, then those hidden in Teams; within each, by team,
 * then channel title, a channel whose team is unknown after every named team.
 */
function byTeamThenTitle(a: ConversationSummary, b: ConversationSummary): number {
  if ((a.hidden === true) !== (b.hidden === true)) return a.hidden ? 1 : -1;
  if (a.team !== b.team) {
    if (a.team === undefined) return 1;
    if (b.team === undefined) return -1;
    const byTeam = a.team.localeCompare(b.team, undefined, { sensitivity: 'base' });
    if (byTeam !== 0) return byTeam;
  }
  return a.title.localeCompare(b.title, undefined, { sensitivity: 'base' });
}

export function registerChannelTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_list_teams_and_channels',
    {
      title: 'List Teams channels',
      description:
        'List every channel the signed-in user is in, across all their teams, sorted by team ' +
        'name then channel title. Each has an id, title, team (when Teams makes it ' +
        'determinable), unread hint, last-activity time (ISO 8601), and a last-message preview ' +
        'and sender. Channels the user has hidden in Teams are listed after the rest with ' +
        'hidden: true. Pass an id (or the channel name) to teams_read_conversation to read its ' +
        'posts and replies. Use unread_only for channels with real posts the user has not read, ' +
        'each confirmed by reading past the read marker (a channel that is gone or no longer ' +
        'readable is skipped and counted in skipped_unreachable); hidden channels are left out ' +
        'of it unless include_hidden is set, as Teams shows them no badge. By default unread_only ' +
        'keeps only channels whose newest unread post is from the last 14 days (since_days widens ' +
        'it); older ones are counted in older_unread_skipped, not listed. ' +
        UNREAD_HINT_DESCRIPTION +
        ' Listing never marks anything read. query matches the channel title or team name. ' +
        FAVORITES_DESCRIPTION +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        unread_only: unreadOnlyArg,
        since_days: listSinceDaysArg,
        include_hidden: includeHiddenArg,
        query: queryArg('channel title or team name'),
        limit: limitArg(MAX_CHANNELS, 'channels', 'favorites and pinned first, then by team and title, kept'),
      },
    },
    async ({ unread_only, include_hidden, since_days, query, limit }) => {
      try {
        // Channels are never untitled group chats, so there is nothing to retitle.
        const { rows, retitle: _retitle, ...unreadInfo } = await listSummaries(chatsvc, {
          kinds: ['channel'],
          unread_only,
          include_hidden,
          since_days,
          query,
        });
        rows.sort(byTeamThenTitle);
        if (!unread_only) rows.sort(favoritesFirst);
        const { rows: channels, truncated } = trimRows(rows, limit, 'start', MAX_CHANNELS, 'first');
        return untrustedResult({ channels, ...unreadInfo, ...(truncated ? { truncated } : {}) });
      } catch (err) {
        throw wrapApiError(err, 'list channels');
      }
    },
  );
}
