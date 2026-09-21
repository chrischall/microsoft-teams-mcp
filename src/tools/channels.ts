/**
 * Teams/channel tools — both read the user's live, signed-in
 * `teams.microsoft.com` tab through the fetchproxy bridge. See `client.ts`'s
 * module doc for why there is no way to select a specific channel: the
 * bridge can only READ the DOM, never navigate it.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { minifiedResult, toolAnnotations } from '@chrischall/mcp-utils';
import type { TeamsClient } from '../client.js';
import { wrapBridgeError } from './errors.js';

export function registerChannelTools(server: McpServer, client: TeamsClient): void {
  server.registerTool(
    'teams_list_teams_and_channels',
    {
      title: 'List Teams and their channels',
      description:
        'List the Teams the signed-in user belongs to and their channels, from the ' +
        '"Teams" nav section\'s sidebar (a DIFFERENT view from Chat). Each row is a team or ' +
        'a channel — itemType distinguishes them; a channel row\'s teamName names its parent ' +
        'team. Requires the user\'s browser to have the Teams-and-Channels view open (not ' +
        'Chat) — if it returns nothing, ask them to click "Teams" in the left nav.',
      annotations: toolAnnotations({ readOnly: true }),
      inputSchema: {},
    },
    async () => {
      try {
        const rows = await client.listTeamsAndChannels();
        return minifiedResult(rows);
      } catch (err) {
        throw wrapBridgeError(err, 'list teams and channels');
      }
    },
  );

  server.registerTool(
    'teams_get_open_channel_posts',
    {
      title: 'Read the currently open channel',
      description:
        'Read the top-level posts in whichever Teams CHANNEL is CURRENTLY OPEN in the ' +
        'user\'s signed-in browser tab — sender, time (prose, e.g. "Tuesday, August 4, 2026 ' +
        '7:50 AM" — not ISO 8601), subject (when the post has a title), and text. Threaded ' +
        'replies under a post are not included, only the top-level posts. There is no way ' +
        'to choose a different channel from here: ask the user to open the channel they ' +
        'mean first (the Teams-and-Channels view, not Chat), or use ' +
        'teams_list_teams_and_channels to show them what is available.',
      annotations: toolAnnotations({ readOnly: true }),
      inputSchema: {},
    },
    async () => {
      try {
        const rows = await client.getOpenChannelPosts();
        return minifiedResult(rows);
      } catch (err) {
        throw wrapBridgeError(err, 'read the open channel');
      }
    },
  );
}
