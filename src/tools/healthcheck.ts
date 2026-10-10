/**
 * Healthcheck. Chats and channels are read through the chat-service API, so
 * the probe is one conversations listing — which also exercises the token
 * capture on a cold start. The bridge status section still matters: the token
 * is captured through it, and the Activity feed reads the page through it.
 *
 * The bridge envelope is mcp-utils' `runBridgeHealthcheck` verbatim; this adds
 * a `chat_service` block beside it — where the token came from, when it
 * expires, and how many conversations the probe listed. A SOURCE LABEL and
 * non-secret detail only, never the token: this result is what people paste
 * into a chat when something is broken.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { toolAnnotations } from '@chrischall/mcp-utils';
import { bridgeHealthcheckDescription, runBridgeHealthcheck } from '@chrischall/mcp-utils/fetchproxy';
import type { TeamsDeps } from '../deps.js';

const HOST_LABEL = 'teams.cloud.microsoft';
const PROBE_PATH = 'chat-service conversation listing (GET /users/ME/conversations)';
const TITLE = 'Verify the Teams chat service and ContextMint Bridge end-to-end';

export function registerHealthcheckTool(server: McpServer, { teams, chatsvc }: TeamsDeps): void {
  server.registerTool(
    'teams_healthcheck',
    {
      title: TITLE,
      description:
        `${bridgeHealthcheckDescription(HOST_LABEL, PROBE_PATH)} Also reports the Teams ` +
        'chat-service token behind every chat and channel tool — its source (browser capture or ' +
        'TEAMS_ACCESS_TOKEN), its expiry, and how many conversations one listing returned — ' +
        'never the token itself. Call this when a real tool fails and you want to know which hop broke.',
      annotations: toolAnnotations({ title: TITLE, readOnly: true, idempotent: true, openWorld: true }),
      inputSchema: {},
    },
    async () => {
      // Set only by a probe that got as far as a listing.
      let conversations: number | null = null;
      const result = await runBridgeHealthcheck({
        server,
        prefix: 'teams',
        hostLabel: HOST_LABEL,
        transport: teams.transport,
        probePath: PROBE_PATH,
        probeFn: async () => {
          conversations = (await chatsvc.listConversations()).length;
          return `${conversations} conversation(s) listed`;
        },
        classifyThrown: (err) =>
          /could not capture a Teams token/i.test(err instanceof Error ? err.message : String(err))
            ? { kind: 'token_capture', hint: TOKEN_CAPTURE_HINT }
            : undefined,
      });

      // Read after the probe, which is what resolves (or captures) the token.
      const expiresAt = chatsvc.tokenExpiresAt();
      const chatService = {
        api_base: chatsvc.apiBase,
        token_source: chatsvc.tokenSource,
        token_expires_at: expiresAt === null ? null : new Date(expiresAt).toISOString(),
        token_minutes_remaining: expiresAt === null ? null : Math.round((expiresAt - Date.now()) / 60_000),
        conversations,
      };

      const first = result.content[0];
      if (first?.type !== 'text') return result;
      const body = JSON.parse(first.text) as Record<string, unknown>;
      return {
        ...result,
        content: [{ type: 'text' as const, text: JSON.stringify({ ...body, chat_service: chatService }, null, 2) }],
      };
    },
  );
}

const TOKEN_CAPTURE_HINT =
  'No Teams access token could be captured. Open a signed-in teams.cloud.microsoft tab and click ' +
  'into any chat (that is what makes Teams call its chat service, where the token is read from), ' +
  'then retry.';
