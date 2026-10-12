/**
 * `teams_get_unread` — every message the user has not read, across all their
 * chats and channels, in one call, from the Teams chat-service API.
 *
 * One listing nominates candidates (activity past the read horizon, or no
 * horizon at all; channels hidden and chats muted in Teams left out); each is then read newest
 * first down to its own horizon, a few at a time, and kept only when a real
 * message from someone else is there (`confirm.ts` — the same rule the list
 * tools' unread_only uses), and only when that message is within `since_days`
 * (a candidate with no activity at all inside the window is not read). Nothing
 * writes the read marker, so the user's unread badges in Teams are exactly as
 * they were.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { toolAnnotations } from '@chrischall/mcp-utils';
import type { TeamsDeps } from '../deps.js';
import {
  conversationSummary,
  matchConversations,
  normaliseMessage,
  oldestFirst,
  type ConversationKind,
  type ConversationSummary,
} from '../model.js';
import { groupThreads } from './conversation.js';
import {
  confirmedSummary,
  confirmUnreadCandidates,
  olderUnreadArg,
  olderUnreadConversations,
  olderUnreadReport,
  sinceCutoff,
  sinceDaysArg,
} from './confirm.js';
import { GROUP_TITLE_DESCRIPTION, includeHiddenArg, isUnreadCandidateFor, MATCH_REASON_DESCRIPTION } from './listing.js';
import { enrichTitles } from './titles.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

const MAX_PER_CONVERSATION = 200;
const DEFAULT_PER_CONVERSATION = 50;
const MAX_CONVERSATIONS = 50;
const DEFAULT_CONVERSATIONS = 20;

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
        'under their root post (a root the user already read is marked missing). Exact, as in ' +
        'Teams: every conversation with activity past its read marker (or never opened) is read ' +
        'past the marker and kept only if it holds a real message from someone else, so a call ' +
        'recording, transcript or other system event never makes a conversation unread. Channels ' +
        'the user has hidden in Teams are left out (include_hidden adds them); one that is gone ' +
        'or no longer readable (404/403) is skipped and counted in skipped_unreachable. Only ' +
        'conversations whose newest unread message is from the last 14 days are returned by ' +
        'default (since_days widens or narrows it); older unread conversations — typically dormant ' +
        'channels — are not read past their marker and are counted in older_unread_skipped; ' +
        'older_unread: true also lists them in older_unread_conversations (id, kind, title, team, ' +
        'lastActivity, and unread_count only for those already read — unread_count_known: false ' +
        'otherwise). Chats the user muted in Teams (muted: true) are left out too and counted in ' +
        'muted_skipped, without being read; include_muted adds them. ' +
        'Unread conversations past max_conversations are listed by id, title and unread_count in ' +
        'more_unread_conversations; read one with teams_read_conversation. A conversation that ' +
        'could not be read is listed in errors and the rest still come back. total_unread covers ' +
        'every unread conversation found; total_unread_exact is false when a per-conversation ' +
        'limit truncated or a read failed. This does NOT mark anything read — the user\'s ' +
        'Teams unread badges are unaffected. ' +
        GROUP_TITLE_DESCRIPTION +
        ' query also matches those participant names, and the last-message preview. ' +
        MATCH_REASON_DESCRIPTION +
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
          .describe(
            'Case-insensitive substring of the chat title, channel name, team name, an untitled ' +
              "chat's participant names or the last-message preview, or an exact conversation id.",
          ),
        include_hidden: includeHiddenArg,
        include_muted: z
          .boolean()
          .optional()
          .describe(
            'Also return conversations the user has muted in Teams (default false: they are left out ' +
              'unread and counted in muted_skipped).',
          ),
        since_days: sinceDaysArg(),
        older_unread: olderUnreadArg(),
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
            `Maximum unread conversations returned with their messages, most recent activity first ` +
              `(default ${DEFAULT_CONVERSATIONS}, max ${MAX_CONVERSATIONS}); the rest are listed in ` +
              'more_unread_conversations.',
          ),
      },
    },
    async ({
      kinds,
      query,
      include_hidden,
      include_muted,
      since_days,
      older_unread,
      per_conversation_limit,
      max_conversations,
    }) => {
      try {
        const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
        const allowed = new Set<ConversationKind>((kinds ?? ['chat', 'channel']).flatMap((k) => KIND_FILTERS[k]));
        // Summaries are built against the FULL listing so a channel can find its team's name.
        let candidates = raw
          .map((c) => conversationSummary(c, me, raw))
          .filter((s) => allowed.has(s.kind) && isUnreadCandidateFor(s, include_hidden));
        // A query can name a group chat or 1:1 by a participant, so every untitled candidate is titled
        // first; it also matches the listing's last-message preview (no extra reads).
        if (query !== undefined) candidates = matchConversations(await enrichTitles(chatsvc, raw, me, candidates), query, {
            messages: true,
          });
        // Muted chats are dropped after the query, so muted_skipped counts only what was asked about.
        const mutedSkipped = include_muted ? 0 : candidates.filter((s) => s.muted).length;
        if (!include_muted) candidates = candidates.filter((s) => !s.muted);
        candidates.sort(byLastActivityDesc);

        const perLimit = per_conversation_limit ?? DEFAULT_PER_CONVERSATION;
        const outcome = await confirmUnreadCandidates(
          chatsvc,
          raw,
          me,
          candidates,
          { limit: perLimit, maxRaw: MAX_RAW_MESSAGES, since: sinceCutoff(since_days) },
        );
        const { skippedUnreachable, older } = outcome;
        // Only what is returned is titled: the unread conversations, those whose read failed and,
        // with older_unread, those the window left out.
        const listedOlder = older_unread ? older : [];
        const retitled = await enrichTitles(chatsvc, raw, me, [
          ...outcome.confirmed.map((c) => c.summary),
          ...outcome.failed.map((f) => f.summary),
          ...listedOlder.map((o) => o.summary),
        ]);
        const confirmed = outcome.confirmed.map((c, i) => ({ ...c, summary: retitled[i] }));
        const failed = outcome.failed.map((f, i) => ({ ...f, summary: retitled[outcome.confirmed.length + i] }));
        const olderTitled = retitled.slice(outcome.confirmed.length + outcome.failed.length);

        const cap = max_conversations ?? DEFAULT_CONVERSATIONS;
        const conversations = confirmed.slice(0, cap).map(({ summary, read }) => {
          const messages = oldestFirst(read.messages).map((m) => normaliseMessage(m, me));
          const body = summary.kind === 'channel' ? { threads: groupThreads(messages) } : { messages };
          return {
            conversation: confirmedSummary(summary, true),
            unread_count: messages.length,
            ...body,
            ...(read.truncated
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
        const rest = confirmed.slice(cap);
        const exact = failed.length === 0 && confirmed.every((c) => !c.read.truncated);
        return untrustedResult({
          total_unread: confirmed.reduce((n, c) => n + c.read.messages.length, 0),
          total_unread_exact: exact,
          conversations,
          skipped_unreachable: skippedUnreachable,
          muted_skipped: mutedSkipped,
          ...olderUnreadReport(older.length, since_days),
          ...olderUnreadConversations(listedOlder, olderTitled),
          ...(failed.length > 0
            ? {
                errors: failed.map((f) => ({
                  conversation: f.summary,
                  error: wrapApiError(f.error, 'read this conversation').message,
                })),
              }
            : {}),
          ...(rest.length > 0
            ? {
                more_unread_conversations: rest.map(({ summary, read }) => ({
                  id: summary.id,
                  title: summary.title,
                  unread_count: read.messages.length,
                })),
              }
            : {}),
        });
      } catch (err) {
        throw wrapApiError(err, 'get unread messages');
      }
    },
  );
}
