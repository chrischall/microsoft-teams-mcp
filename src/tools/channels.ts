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
import { listSummaries, queryArg, unreadOnlyArg } from './listing.js';

/** Up to 10 pages of 200 come back; this bounds one result, not the listing. */
const MAX_CHANNELS = 500;

/** By team, then channel title; a channel whose team is unknown sorts after every named team. */
function byTeamThenTitle(a: ConversationSummary, b: ConversationSummary): number {
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
        'determinable), unread flag, last-activity time (ISO 8601), and a last-message preview ' +
        'and sender. Pass an id (or the channel name) to teams_read_conversation to read its ' +
        'posts and replies. Use unread_only for channels with posts the user has not read; ' +
        'listing never marks anything read. query matches the channel title or team name.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        unread_only: unreadOnlyArg,
        query: queryArg('channel title or team name'),
        limit: limitArg(MAX_CHANNELS, 'channels', 'first by team and title kept'),
      },
    },
    async ({ unread_only, query, limit }) => {
      try {
        const rows = await listSummaries(chatsvc, { kinds: ['channel'], unread_only, query });
        rows.sort(byTeamThenTitle);
        const { rows: channels, truncated } = trimRows(rows, limit, 'start', MAX_CHANNELS, 'first');
        return untrustedResult({ channels, ...(truncated ? { truncated } : {}) });
      } catch (err) {
        throw wrapApiError(err, 'list channels');
      }
    },
  );
}
