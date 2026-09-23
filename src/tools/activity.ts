/**
 * Activity tool — reads the user's live, signed-in `teams.cloud.microsoft`
 * tab through the fetchproxy bridge. See `client.ts`'s module doc for why
 * there is no way to select a specific activity item: the bridge can only
 * READ the DOM, never navigate it.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { toolAnnotations } from '@chrischall/mcp-utils';
import type { TeamsClient } from '../client.js';
import { wrapBridgeError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

export function registerActivityTools(server: McpServer, client: TeamsClient): void {
  server.registerTool(
    'teams_get_activity',
    {
      title: 'Read the Activity feed',
      description:
        'List the rows in the Activity feed (the bell icon in the left nav) — mentions, ' +
        'replies, reactions, and the like across every chat and channel — each with its ' +
        'title, message preview, prose time (not ISO 8601), and location (the team/channel ' +
        'or chat it happened in). Requires the user\'s browser to have the Activity view ' +
        'open — if it returns nothing, ask them to click the bell icon in the left nav.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true }),
      inputSchema: {},
    },
    async () => {
      try {
        const rows = await client.getActivity();
        return untrustedResult({ rows });
      } catch (err) {
        throw wrapBridgeError(err, 'read the Activity feed');
      }
    },
  );
}
