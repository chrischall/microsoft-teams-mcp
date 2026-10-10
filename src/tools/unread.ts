/**
 * `teams_get_unread` — every message the user has not read, across all their
 * chats and channels, in one call, from the Teams chat-service API.
 *
 * One listing finds the unread conversations (the last message is past the
 * read horizon and not the user's own); each is then read newest first down
 * to its own horizon, a few at a time. Nothing writes the read marker, so the
 * user's unread badges in Teams are exactly as they were.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { mapWithConcurrency, toolAnnotations } from '@chrischall/mcp-utils';
import type { RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import {
  conversationSummary,
  isUnread,
  isUserMessage,
  matchConversations,
  normaliseMessage,
  readHorizon,
  type ConversationKind,
  type ConversationSummary,
} from '../model.js';
import { groupThreads } from './conversation.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

const MAX_PER_CONVERSATION = 200;
const DEFAULT_PER_CONVERSATION = 50;
const MAX_CONVERSATIONS = 50;
const DEFAULT_CONVERSATIONS = 20;

/** Conversations read at once. */
const CONCURRENCY = 4;

/**
 * Raw messages read per conversation before giving up. The real bound is the
 * read horizon; this only caps a backlog of nothing but system events.
 */
const MAX_RAW_MESSAGES = 4000;

/** The `kinds` filter: `chat` is every non-channel chat; the others narrow to one kind. */
const KIND_FILTERS = {
  chat: ['group', 'meeting', 'oneOnOne'],
  channel: ['channel'],
  meeting: ['meeting'],
  oneOnOne: ['oneOnOne'],
} as const satisfies Record<string, readonly ConversationKind[]>;

type KindFilter = keyof typeof KIND_FILTERS;

/** Newest activity first; a conversation with no messages at all sorts last. */
function byLastActivityDesc(a: ConversationSummary, b: ConversationSummary): number {
  return (b.lastActivity ?? '').localeCompare(a.lastActivity ?? '');
}

export function registerUnreadTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_get_unread',
    {
      title: 'Get unread Teams messages',
      description:
        'Get every unread message across all the chats and channels the signed-in user is in ' +
        '(or those matching kinds / query), grouped by conversation, most recent activity ' +
        'first. Each conversation comes with its summary (id, kind, title, team) and only the ' +
        'messages the user has not read, oldest first; a channel\'s are grouped as threads ' +
        'under their root post (a root the user already read is marked missing). Conversations ' +
        'past max_conversations are listed by id and title in more_unread_conversations; read ' +
        'one with teams_read_conversation. This does NOT mark anything read — the user\'s ' +
        'Teams unread badges are unaffected.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        kinds: z
          .array(z.enum(Object.keys(KIND_FILTERS) as [KindFilter, ...KindFilter[]]))
          .min(1)
          .optional()
          .describe(
            'Only these kinds of conversation (default all): chat = every group, meeting and 1:1 ' +
              'chat; channel; meeting = meeting chats only; oneOnOne = 1:1 chats only.',
          ),
        query: z
          .string()
          .min(1)
          .optional()
          .describe('Case-insensitive substring of the chat title, channel name or team name, or an exact conversation id.'),
        per_conversation_limit: z
          .number()
          .int()
          .min(1)
          .max(MAX_PER_CONVERSATION)
          .optional()
          .describe(
            `Maximum unread messages per conversation, newest kept (default ${DEFAULT_PER_CONVERSATION}, ` +
              `max ${MAX_PER_CONVERSATION}).`,
          ),
        max_conversations: z
          .number()
          .int()
          .min(1)
          .max(MAX_CONVERSATIONS)
          .optional()
          .describe(
            `Maximum conversations to read, most recent activity first (default ${DEFAULT_CONVERSATIONS}, ` +
              `max ${MAX_CONVERSATIONS}); the rest are listed in more_unread_conversations.`,
          ),
      },
    },
    async ({ kinds, query, per_conversation_limit, max_conversations }) => {
      try {
        const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
        const allowed = new Set<ConversationKind>((kinds ?? ['chat', 'channel']).flatMap((k) => KIND_FILTERS[k]));
        // Summaries are built against the FULL listing so a channel can find its team's name.
        let unread = raw
          .map((c) => conversationSummary(c, me, raw))
          .filter((s) => s.unread && allowed.has(s.kind));
        if (query !== undefined) unread = matchConversations(unread, query);
        unread.sort(byLastActivityDesc);

        const cap = max_conversations ?? DEFAULT_CONVERSATIONS;
        const perLimit = per_conversation_limit ?? DEFAULT_PER_CONVERSATION;
        const toRead = unread.slice(0, cap);
        const rest = unread.slice(cap);

        const read = await mapWithConcurrency(toRead, CONCURRENCY, async (summary) => {
          const horizon = readHorizon(raw.find((c) => c.id === summary.id)!);
          const wanted = (m: RawMessage): boolean => isUserMessage(m) && isUnread(m, horizon, me);

          // As in teams_read_conversation: the first wanted message past the
          // limit is the boundary, and its existence says the read was truncated.
          let kept = 0;
          let truncated = false;
          const stopAt = (m: RawMessage): boolean => {
            if (Number(m.id) <= horizon) return true;
            if (!wanted(m)) return false;
            if (kept >= perLimit) return (truncated = true);
            kept++;
            return false;
          };
          const fetched = await chatsvc.getMessages(summary.id, { stopAt, max: MAX_RAW_MESSAGES });
          const messages = fetched
            .filter(wanted)
            .reverse()
            .map((m) => normaliseMessage(m, me));
          const body = summary.kind === 'channel' ? { threads: groupThreads(messages) } : { messages };
          return {
            conversation: summary,
            unread_count: messages.length,
            ...body,
            ...(truncated
              ? {
                  truncated: {
                    returned: messages.length,
                    hint:
                      `Only the newest ${messages.length} unread messages were returned; pass a larger ` +
                      `per_conversation_limit (max ${MAX_PER_CONVERSATION}) or read the conversation ` +
                      'with teams_read_conversation (unread_only).',
                  },
                }
              : {}),
          };
        });

        // The listing can call a conversation unread when what is past the
        // horizon has since been deleted; leave those out.
        const conversations = read.filter((c) => c.unread_count > 0);
        return untrustedResult({
          total_unread: conversations.reduce((n, c) => n + c.unread_count, 0),
          conversations,
          ...(rest.length > 0
            ? { more_unread_conversations: rest.map((s) => ({ id: s.id, title: s.title })) }
            : {}),
        });
      } catch (err) {
        throw wrapApiError(err, 'get unread messages');
      }
    },
  );
}
