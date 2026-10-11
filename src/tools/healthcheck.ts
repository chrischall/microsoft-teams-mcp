/**
 * Healthcheck — the CREDENTIAL healthcheck, not the bridge one.
 *
 * Every tool, the Activity feed included, reads through the chat-service API
 * with plain server-side requests. The browser bridge is touched only by the
 * token capture, which builds a transport, lifts one header and closes it
 * again (`auth-fetchproxy.ts`) — so there is no long-lived bridge whose role
 * or port would mean anything here. What a broken install needs to know is
 * whether a token resolved (and from where) and whether the chat service
 * accepted it; that is mcp-utils' `runCredentialHealthcheck`, probed with one
 * conversations listing.
 *
 * A `chat_service` block sits beside it — the API base, the token's source and
 * expiry, and how many conversations the probe listed. A SOURCE LABEL and
 * non-secret detail only, never the token: this result is what people paste
 * into a chat when something is broken.
 *
 * A `sharepoint` note sits beside those two. Shared-file download
 * (`teams_get_attachment`) uses a SEPARATE SharePoint token, captured on demand
 * per SharePoint host and short-lived (~53 min) — so there is usually none held
 * until a file is opened, and this healthcheck never mints or probes one. It is
 * a static note, not a check: a healthy result does not depend on an SP token
 * existing.
 *
 * Imported from `/healthcheck`, deliberately NOT from `/fetchproxy`: that
 * module pulls in the optional `@fetchproxy/server` peer, which is absent in
 * the `.mcpb` bundle.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { toolAnnotations } from '@chrischall/mcp-utils';
import { credentialHealthcheckDescription, runCredentialHealthcheck } from '@chrischall/mcp-utils/healthcheck';
import type { TeamsDeps } from '../deps.js';

const HOST_LABEL = 'teams.cloud.microsoft';
const TITLE = 'Verify the Teams chat-service token and API end-to-end';

/**
 * Static note on the SharePoint path: a separate token, captured on demand per
 * host and short-lived, that `teams_get_attachment` uses for shared-file
 * download. Not probed here — a healthy result never depends on it existing.
 */
const SHAREPOINT_NOTE =
  'Shared-file download (teams_get_attachment) uses a separate SharePoint token, captured on ' +
  'demand per SharePoint host and short-lived (~53 min). None is held until you open a file, and ' +
  'this healthcheck neither mints nor requires one.';

export function registerHealthcheckTool(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_healthcheck',
    {
      title: TITLE,
      description:
        `${credentialHealthcheckDescription(HOST_LABEL)} The credential is the Teams chat-service ` +
        'token behind every tool; the probe is one conversations listing. Also reports the ' +
        'token\'s source (browser capture through the ContextMint Bridge, or TEAMS_ACCESS_TOKEN), ' +
        'its expiry, and how many conversations the listing returned — never the token itself. ' +
        'Call this when a real tool fails and you want to know which hop broke.',
      annotations: toolAnnotations({ title: TITLE, readOnly: true, idempotent: true, openWorld: true }),
      inputSchema: {},
    },
    async () => {
      // Set only by a probe that got as far as a listing.
      let conversations: number | null = null;
      // TEAMS_CHATSVC_BASE can point the client elsewhere; report what it really talks to.
      const base = new URL(chatsvc.apiBase);
      const result = await runCredentialHealthcheck({
        server,
        prefix: 'teams',
        hostLabel: base.host,
        probePath: `${base.pathname}/users/ME/conversations`,
        // Resolves the token the way every tool does — same cache, same capture.
        resolveCredential: async () => {
          await chatsvc.ensureToken();
          return { source: chatsvc.tokenSource };
        },
        probeFn: async () => {
          conversations = (await chatsvc.listConversations()).length;
          return `${conversations} conversation(s) listed`;
        },
        classifyThrown: (err) =>
          /could not capture a Teams token/i.test(err instanceof Error ? err.message : String(err))
            ? { kind: 'token_capture', hint: TOKEN_CAPTURE_HINT }
            : undefined,
      });

      // Read after the resolve, which is what loads (or captures) the token.
      const expiresAt = chatsvc.tokenExpiresAt();
      const chatService = {
        api_base: chatsvc.apiBase,
        token_source: chatsvc.tokenSource,
        token_expires_at: expiresAt === null ? null : new Date(expiresAt).toISOString(),
        token_minutes_remaining: expiresAt === null ? null : Math.round((expiresAt - Date.now()) / 60_000),
        conversations,
      };

      const first = result.content[0];
      let body: unknown;
      try {
        body = first === undefined ? undefined : JSON.parse(first.text);
      } catch {
        body = undefined;
      }
      if (body === null || typeof body !== 'object' || Array.isArray(body)) {
        // Not the JSON envelope this expects: keep its text as-is and report the
        // chat service beside it rather than failing the healthcheck.
        return {
          content: [
            ...result.content,
            {
              type: 'text' as const,
              text: JSON.stringify({ chat_service: chatService, sharepoint: SHAREPOINT_NOTE }, null, 2),
            },
          ],
        };
      }
      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify({ ...body, chat_service: chatService, sharepoint: SHAREPOINT_NOTE }, null, 2),
          },
        ],
      };
    },
  );
}

const TOKEN_CAPTURE_HINT =
  'No Teams access token could be captured. Open a signed-in teams.cloud.microsoft tab and click ' +
  'into any chat (that is what makes Teams call its chat service, where the token is read from), ' +
  'then retry.';
