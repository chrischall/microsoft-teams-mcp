/**
 * `teams_get_unread` — every message the user has not read, across all their
 * chats and channels, in one call, from the Teams chat-service API.
 *
 * One listing nominates candidates (activity past the read horizon, or no
 * horizon at all; channels hidden and chats muted in Teams left out); each is then read newest
 * first down to its own horizon, a few at a time, and kept only when a real
 * message from someone else is there (`confirm.ts` — the same rule the list
 * tools' unread_only uses), and only when that message is within `since_days`
 * (a candidate with no activity at all inside the window is not read). With
 * `cursors`, each read stops at the later of the read marker and that
 * conversation's cursor, so only what is new since the caller last looked
 * comes back (`cursor.ts`). By default nothing writes the read marker, so the
 * user's unread badges in Teams are exactly as they were; `mark_read`
 * (confirm-gated, as `teams_mark_read` is) moves each returned conversation's
 * marker up to its newest unread message returned — never with format
 * `summary`, and never for a conversation `per_conversation_limit` cut.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import {
  CONFIRM_INJECTION_RULE,
  confirmTokenParam,
  mapWithConcurrency,
  McpToolError,
  toolAnnotations,
} from '@chrischall/mcp-utils';
import type { TeamsDeps } from '../deps.js';
import type { ChatServiceClient, RawConversation, RawMessage } from '../chatsvc.js';
import {
  conversationSummary,
  isUnread,
  lastChangeMs,
  matchConversations,
  normaliseMessage,
  oldestFirst,
  readHorizon,
  type ConversationKind,
  type ConversationSummary,
} from '../model.js';
import {
  CONFIRM_CONCURRENCY,
  confirmedSummary,
  confirmUnreadCandidates,
  isUnreachable,
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
import { resolveTimeZone, TIME_ZONE_DESCRIPTION, tzArg } from './timezone.js';
import { CURSOR_NO_CACHE_DESCRIPTION, cursorsArg, newestId } from './cursor.js';
import { FORMAT_DESCRIPTION, formatArg, formatBody, formatConversation, formatFlatMessage, type ReadFormat } from './format.js';
import { arrivalMs } from './conversation.js';
import { DEFAULT_LIMIT } from './limit.js';
import {
  decodeFrom,
  FROM_DESCRIPTION,
  fromArg,
  fromReport,
  readFrom,
  takePage,
  type FromPosition,
  type Keyed,
} from './from.js';
import { fetchOutsideRoots, MAX_ROOT_FETCHES, rootBudget } from './roots.js';
import {
  advancesHorizon,
  applyMarks,
  confirmMarks,
  horizonMessageId,
  MARK_READ_DESCRIPTION,
  markReadArg,
  assertMarkReadFormat,
  type MarkOutcome,
  type MarkPlan,
  type MarkSkip,
} from './markread.js';

const MAX_PER_CONVERSATION = 200;
const DEFAULT_PER_CONVERSATION = 50;
const MAX_CONVERSATIONS = 50;
const DEFAULT_CONVERSATIONS = 20;
/** Messages one `from` page may return, across every conversation. */
const MAX_FROM_MESSAGES = 500;

/**
 * Raw messages read per conversation before giving up. The real bound is the
 * read horizon; this only caps a backlog of nothing but system events.
 */
const MAX_RAW_MESSAGES = 4000;

/** Channels whose outside-window roots are fetched at once (each fetch is itself bounded, `roots.ts`). */
const ROOT_CONVERSATION_CONCURRENCY = 4;

/** The `kinds` filter: `chat` is every non-channel chat; the others narrow to one kind. */
export const KIND_FILTERS = {
  chat: ['group', 'meeting', 'oneOnOne'],
  channel: ['channel'],
  meeting: ['meeting'],
  oneOnOne: ['oneOnOne'],
} as const satisfies Record<string, readonly ConversationKind[]>;

export type KindFilter = keyof typeof KIND_FILTERS;

/** What `mark_read` reports for a conversation `per_conversation_limit` cut: its marker is left alone. */
const TRUNCATED_SKIP: MarkSkip = {
  skipped: 'truncated',
  hint:
    'The read marker was left where it was: per_conversation_limit cut this read, so unread messages older ' +
    `than those returned were not returned. Read again with a larger per_conversation_limit (max ${MAX_PER_CONVERSATION}) ` +
    'or with teams_read_conversation (unread_only) to mark it read, or use teams_mark_read to mark it all read deliberately.',
};

/** Newest activity first; a conversation with no messages at all sorts last. */
function byLastActivityDesc(a: ConversationSummary, b: ConversationSummary): number {
  return (b.lastActivity ?? '').localeCompare(a.lastActivity ?? '');
}

/**
 * One `from` page of unread messages across `candidates`: each read past the
 * later of `from` and its own horizon (one with no change since `from` is not
 * read), its real unread messages pooled and taken oldest first. The position
 * never moves past the newest change the listing knew of — a message landing
 * mid-call in a conversation already read is then not skipped — nor at all
 * when a read failed, since that conversation may hold older unread messages.
 */
async function unreadFrom(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  candidates: readonly ConversationSummary[],
  pos: FromPosition,
  { limit, timeZone, format }: { limit: number; timeZone?: string; format: Exclude<ReadFormat, 'summary'> },
): Promise<Record<string, unknown>> {
  const byId = new Map(raw.map((c) => [c.id, c]));
  const toRead = candidates.filter((s) => (lastChangeMs(byId.get(s.id) ?? { id: s.id }) ?? Infinity) >= pos.t);
  type Result =
    | { kind: 'read'; messages: RawMessage[] }
    | { kind: 'unreachable' }
    | { kind: 'failed'; error: unknown };
  const results = await mapWithConcurrency(toRead, CONFIRM_CONCURRENCY, async (summary): Promise<Result> => {
    const horizon = readHorizon(byId.get(summary.id) ?? { id: summary.id });
    try {
      const fetched = await readFrom(chatsvc, summary.id, Math.max(pos.t, horizon + 1));
      return { kind: 'read', messages: fetched.filter((m) => isUnread(m, horizon, me)) };
    } catch (error) {
      return isUnreachable(error) ? { kind: 'unreachable' } : { kind: 'failed', error };
    }
  });

  let skippedUnreachable = 0;
  let anyRead = false;
  const items: Keyed<{ summary: ConversationSummary; m: RawMessage }>[] = [];
  const failed: { summary: ConversationSummary; error: unknown }[] = [];
  results.forEach((r, i) => {
    const summary = toRead[i];
    if (r.kind === 'unreachable') skippedUnreachable++;
    else if (r.kind === 'failed') failed.push({ summary, error: r.error });
    else {
      anyRead = true;
      for (const m of r.messages) items.push({ ms: arrivalMs(m), key: `${summary.id}/${m.id}`, item: { summary, m } });
    }
  });
  // Every read failed: a failure of the call (a revoked token, a rate limit), not of one conversation.
  if (failed.length > 0 && !anyRead) throw failed[0].error;

  const listed = Math.max(pos.t, ...raw.map((c) => lastChangeMs(c) ?? 0));
  const { page, has_more, remaining, next } = takePage(items, pos, limit, failed.length > 0 ? pos.t : listed);

  // Only what is returned is titled: the conversations on this page and those whose read failed.
  const shown = [...new Map(page.map(({ summary }) => [summary.id, summary])).values()];
  const titled = await enrichTitles(chatsvc, raw, me, [...shown, ...failed.map((f) => f.summary)]);
  return {
    total_unread: page.length + remaining,
    messages: page.map(({ summary, m }) => ({
      conversation_id: summary.id,
      ...formatFlatMessage(normaliseMessage(m, me, { timeZone }), format),
    })),
    conversations: titled.slice(0, shown.length).map((s) => formatConversation(confirmedSummary(s, true), format)),
    ...fromReport(pos, next, has_more),
    skipped_unreachable: skippedUnreachable,
    ...(failed.length > 0
      ? {
          errors: failed.map((f, i) => ({
            conversation: titled[shown.length + i],
            error: wrapApiError(f.error, 'read this conversation').message,
          })),
        }
      : {}),
  };
}

/** The arguments that shape a per-conversation answer, which a `from` page replaces. */
const NOT_WITH_FROM = ['since_days', 'cursors', 'per_conversation_limit', 'max_conversations', 'mark_read', 'older_unread'] as const;

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
        'under their root post. A root the user already read is fetched and marked outside_window: true ' +
        `(it is context, not unread or counted); up to ${MAX_ROOT_FETCHES} roots are fetched per call, and ` +
        'any beyond that or that cannot be fetched is { id, missing: true }. Exact, as in ' +
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
        'limit truncated or a read failed. cursors ({ conversationId: messageId }) returns only unread ' +
        'messages newer than each conversation\'s cursor (one with none is left out, and total_unread ' +
        'counts only those). Each returned conversation carries next_cursor, its newest returned message ' +
        'id; next_cursors is the cursors passed in updated with every returned conversation\'s next_cursor ' +
        '(not those only in more_unread_conversations) — pass it back as cursors to get only what is new. ' +
        CURSOR_NO_CACHE_DESCRIPTION +
        FROM_DESCRIPTION +
        'With from the result is one list across every conversation instead: messages (each with its ' +
        'conversation_id; a channel reply keeps replyTo), oldest first, at most limit; conversations (the summary ' +
        'of each one on the page); total_unread (every unread message past from, this page and later ones); ' +
        'from_time, has_more, next_from, next_from_time; skipped_unreachable, muted_skipped and errors. next_from ' +
        'does not move while a read failed (call again). from takes the place of since_days, cursors, ' +
        'per_conversation_limit, max_conversations, older_unread and mark_read, and format summary does not apply. ' +
        FORMAT_DESCRIPTION +
        MARK_READ_DESCRIPTION +
        'With mark_read each conversation returned with its messages (not more_unread_conversations) has its ' +
        'read marker moved up to its newest unread message returned (outside_window roots do not count) and ' +
        'carries marked_read (false with mark_read_warning when that move failed); one per_conversation_limit ' +
        'cut is not marked — marked_read: { skipped: "truncated", hint }; the user confirms first — a prompt where the client supports one, otherwise the ' +
        'first call returns only a preview and a confirmToken, and the repeat call with it reads and marks. ' +
        CONFIRM_INJECTION_RULE +
        ' ' +
        GROUP_TITLE_DESCRIPTION +
        ' query also matches those participant names, and the last-message preview. ' +
        MATCH_REASON_DESCRIPTION +
        'Message times are UTC (ISO 8601); ' +
        TIME_ZONE_DESCRIPTION +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      // A write only with mark_read, but a write: it moves the user's read markers (never backwards, so a repeat is a no-op).
      annotations: toolAnnotations({ readOnly: false, destructive: false, idempotent: true, openWorld: true }),
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
        cursors: cursorsArg,
        from: fromArg,
        limit: z
          .number()
          .int()
          .min(1)
          .max(MAX_FROM_MESSAGES)
          .optional()
          .describe(
            `With from only: maximum messages returned across every conversation, oldest first (default ` +
              `${DEFAULT_LIMIT}, max ${MAX_FROM_MESSAGES}).`,
          ),
        tz: tzArg,
        format: formatArg,
        mark_read: markReadArg,
        confirmToken: confirmTokenParam,
      },
    },
    async (
      {
        kinds,
        query,
        include_hidden,
        include_muted,
        since_days,
        older_unread,
        per_conversation_limit,
        max_conversations,
        cursors,
        from,
        limit,
        tz,
        format,
        mark_read,
        confirmToken,
      },
      ctx,
    ) => {
      try {
        assertMarkReadFormat(mark_read, format);
        if (from === undefined && limit !== undefined) {
          throw new McpToolError('limit applies only with from.', {
            hint: 'Without from, use per_conversation_limit and max_conversations.',
          });
        }
        if (from !== undefined) {
          const args = { since_days, cursors, per_conversation_limit, max_conversations, mark_read, older_unread };
          const clash = NOT_WITH_FROM.find((k) => args[k] !== undefined && args[k] !== false);
          if (clash !== undefined || format === 'summary') {
            throw new McpToolError(`from cannot be combined with ${clash ?? 'format summary'}.`, {
              hint:
                'from pages one oldest-first list across every conversation: pass a previous next_from to go on, ' +
                'and use teams_mark_read to mark what was paged read.',
            });
          }
        }
        const pos = from !== undefined ? decodeFrom(from) : undefined;
        const timeZone = resolveTimeZone(tz);
        const fetched_at = new Date().toISOString();
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

        if (pos !== undefined) {
          const page = await unreadFrom(chatsvc, raw, me, candidates, pos, {
            limit: limit ?? DEFAULT_LIMIT,
            timeZone,
            format: format === 'compact' ? 'compact' : 'full',
          });
          return untrustedResult({ ...page, muted_skipped: mutedSkipped, fetched_at });
        }

        const perLimit = per_conversation_limit ?? DEFAULT_PER_CONVERSATION;
        const outcome = await confirmUnreadCandidates(
          chatsvc,
          raw,
          me,
          candidates,
          { limit: perLimit, maxRaw: MAX_RAW_MESSAGES, since: sinceCutoff(since_days), cursors },
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
        const returned = confirmed.slice(0, cap);

        // Marked before roots are fetched, so a call that stops at the confirmation preview fetches none.
        let marks: Map<string, MarkOutcome> | undefined;
        if (mark_read) {
          const byId = new Map(raw.map((c) => [c.id, c]));
          const plans: MarkPlan[] = [];
          for (const { summary, read } of returned) {
            // A cut read left older unread messages out: those are not the caller's to mark (see TRUNCATED_SKIP).
            if (read.truncated) continue;
            const conv = byId.get(summary.id) ?? { id: summary.id };
            const target = newestId(read.messages);
            if (advancesHorizon(conv, target)) plans.push({ summary, previous: horizonMessageId(conv), target });
          }
          if (plans.length > 0) {
            const gate = await confirmMarks(ctx, { tool: 'teams_get_unread', me, plans, confirmToken });
            if (gate) return gate;
          }
          marks = await applyMarks(chatsvc, plans);
        }

        const budget = rootBudget();
        // Channels claim root fetches in the order returned, so the most recent activity gets them first.
        const shown = returned.map(({ summary, read }) => ({
          summary,
          read,
          messages: oldestFirst(read.messages).map((m) => normaliseMessage(m, me, { timeZone })),
        }));
        const roots = await mapWithConcurrency(shown, ROOT_CONVERSATION_CONCURRENCY, ({ summary, messages }) =>
          summary.kind === 'channel'
            ? fetchOutsideRoots(chatsvc, summary.id, messages, me, { budget, timeZone })
            : Promise.resolve(undefined),
        );
        const conversations = shown.map(({ summary, read, messages }, i) => ({
          conversation: formatConversation(confirmedSummary(summary, true), format ?? 'full'),
          unread_count: messages.length,
          ...formatBody(summary.kind, messages, format ?? 'full', roots[i]),
          // Every confirmed read holds at least one message, so there is always a newest.
          next_cursor: newestId(read.messages)!,
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
          ...(marks !== undefined
            ? (marks.get(summary.id) ?? { marked_read: read.truncated ? TRUNCATED_SKIP : false })
            : {}),
        }));
        const rest = confirmed.slice(cap);
        const next_cursors: Record<string, string> = { ...cursors };
        for (const c of conversations) next_cursors[c.conversation.id] = c.next_cursor;
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
          next_cursors,
          fetched_at,
        });
      } catch (err) {
        throw wrapApiError(err, 'get unread messages');
      }
    },
  );
}
