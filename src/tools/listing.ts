/**
 * The conversation listing shared by `teams_list_chats` and
 * `teams_list_teams_and_channels`: one `listConversations` call, normalised
 * against the signed-in user, then filtered by kind, unread state and query.
 * With `unread_only`, the listing's unread flag only nominates candidates; each
 * is confirmed by reading past its horizon (`confirm.ts`).
 */
import { z } from 'zod';
import type { ChatServiceClient } from '../chatsvc.js';
import {
  confirmedSummary,
  confirmUnreadCandidates,
  olderUnreadArg,
  olderUnreadConversations,
  olderUnreadReport,
  sinceCutoff,
  sinceDaysArg,
  type OlderUnreadRow,
} from './confirm.js';
import { wrapApiError } from './errors.js';
import { searchRecentMessages, type MessageHit } from './chat-search.js';
import { enrichTitles, participantsOf } from './titles.js';
import {
  conversationSummary,
  matchConversations,
  SELF_CHAT,
  SELF_CHAT_ALIASES,
  type ConversationKind,
  type ConversationSummary,
} from '../model.js';

export interface ListFilter {
  kinds: readonly ConversationKind[];
  unread_only?: boolean;
  /** With `unread_only`, keep channels hidden in Teams too (default false). */
  include_hidden?: boolean;
  /** With `unread_only`, the window (days) the newest unread message must fall in (default 14). */
  since_days?: number;
  /** With `unread_only`: also list what `since_days` left out (`older_unread_conversations`). */
  older_unread?: boolean;
  query?: string;
  /**
   * With `query`: also read each chat no name or preview matched and match its
   * recent messages (`search.ts`).
   */
  searchMessages?: boolean;
  /** Epoch ms: keep conversations whose last activity is at or after it; one with none is dropped. */
  activeSince?: number;
  /** Epoch ms: keep conversations whose last activity is at or before it; one with none is dropped. */
  activeUntil?: number;
}

/**
 * Whether a conversation's last activity falls in `[since, until]` (either
 * end open when undefined). With either end given, a conversation with no
 * readable last activity is out: nothing places it in the window.
 */
export function activeWithin(s: Pick<ConversationSummary, 'lastActivity'>, since?: number, until?: number): boolean {
  if (since === undefined && until === undefined) return true;
  const t = s.lastActivity === undefined ? Number.NaN : Date.parse(s.lastActivity);
  if (Number.isNaN(t)) return false;
  return (since === undefined || t >= since) && (until === undefined || t <= until);
}

/** The list tools' `since_days`: it narrows `unread_only` only. */
export const listSinceDaysArg = sinceDaysArg('With unread_only: ');

/** The list tools' `older_unread`: it needs `unread_only`. */
export const listOlderUnreadArg = olderUnreadArg('With unread_only: ');

export const unreadOnlyArg = z
  .boolean()
  .optional()
  .describe(
    'Only return conversations with a real message (not a call recording or other system event) ' +
      'the user has not read yet, each with its unread_count (default false). Exact: every candidate ' +
      'is confirmed by reading its messages past the read marker. Never marks anything read (teams_mark_read does).',
  );

/** Description text shared by the list tools on what their unread flag means. */
export const UNREAD_HINT_DESCRIPTION =
  'Without unread_only, unread is only a hint from the listing (unread_confirmed: false): it can ' +
  'flag a conversation whose only news is a call recording or other system event. For an exact ' +
  'answer use unread_only (confirmed rows carry unread_confirmed: true and unread_count) or ' +
  'teams_get_unread.';

/** Real unread messages counted per candidate before unread_count is reported as a lower bound. */
const CONFIRM_LIMIT = 50;
/** Raw messages read per candidate — one page; the horizon is usually much nearer. */
const CONFIRM_MAX_RAW = 50;

export const includeHiddenArg = z
  .boolean()
  .optional()
  .describe(
    'Also count channels the user has hidden in Teams as unread (default false: Teams shows no ' +
      'unread badge for a hidden channel, so neither does this).',
  );

/**
 * Whether a conversation is an unread candidate worth confirming: the listing
 * flags it, and it is not a channel hidden in Teams (unless the caller asked
 * for hidden ones — Teams shows a hidden channel no badge).
 */
export function isUnreadCandidateFor(s: ConversationSummary, includeHidden: boolean | undefined): boolean {
  return s.unread && (includeHidden === true || s.hidden !== true);
}

export function queryArg(what: string) {
  return z
    .string()
    .min(1)
    .optional()
    .describe(`Case-insensitive substring of the ${what}, or an exact conversation id.`);
}

/** Description text on the match_reason a query adds to each row. */
export const MATCH_REASON_DESCRIPTION =
  'With a query, each row carries match_reason — the first of "id" (exact id), "title", "team", ' +
  '"participant" (an untitled chat\'s participant) or "message" (what was said, not what the ' +
  'conversation is called) that applies.';

/** Description text on how an untitled group chat or 1:1 is titled. */
export const GROUP_TITLE_DESCRIPTION =
  'A group chat with no topic is titled "Group chat: " and the first three of its recent senders other ' +
  'than the user, sorted by name ("; " between names, "+N" for more) — so it does not change when a ' +
  'different person posts last — and carries every such recent sender, in that order, as participants; ' +
  'it is "(untitled group chat)" when no one else wrote recently. ' +
  'A 1:1 is titled by the other person, who is its one participant; it is "(1:1 chat)" when they ' +
  'wrote nothing recently.'

/**
 * Favorites, then pinned conversations, ahead of the rest — the plain list
 * tools' first sort key, as in the Teams sidebar. 0 between two of a kind, so
 * a stable sort keeps each tool's own order within them.
 */
export function favoritesFirst(a: Pick<ConversationSummary, 'favorite' | 'pinned'>, b: Pick<ConversationSummary, 'favorite' | 'pinned'>): number {
  const rank = (s: Pick<ConversationSummary, 'favorite' | 'pinned'>) => (s.favorite ? 0 : s.pinned ? 1 : 2);
  return rank(a) - rank(b);
}

/** Description text on the favorite / pinned flags and order. */
export const FAVORITES_DESCRIPTION =
  'Conversations the user marked a favorite or pinned in Teams carry favorite: true / pinned: true ' +
  'and, without unread_only, are listed first (favorites, then pinned); unread_only keeps its usual order.';

export type ListedRow = ConversationSummary & {
  unread_count?: number;
  unread_count_exact?: false;
  /** With `search_messages`: the newest recent message that matched the query. */
  matched_message?: MessageHit;
};

export interface Listing {
  rows: ListedRow[];
  /** With `unread_only`: candidates that answered 404/403 and so are not unread. */
  skipped_unreachable?: number;
  /** With `unread_only`: conversations left out because their unread activity is older than `since_days`. */
  older_unread_skipped?: number;
  older_unread_hint?: string;
  /** With `unread_only` and `older_unread`: the conversations `older_unread_skipped` counts, titled. */
  older_unread_conversations?: OlderUnreadRow[];
  /** With `unread_only`: candidates whose confirming read failed; they are left out of `rows`. */
  errors?: { conversation: { id: string; title: string }; error: string }[];
  /** With `search_messages`: chats whose recent messages could not be read, so were matched by name only. */
  messages_not_searched?: number;
}

/**
 * A listing plus `retitle`, which titles the untitled group chats and 1:1s among the
 * rows a tool is about to return — and among the conversations in `errors` —
 * by their participants (`titles.ts`). Kept apart from the rows so a tool
 * reads only the chats it returns.
 */
export type ListingWithRetitle = Listing & {
  retitle: (rows: ListedRow[]) => Promise<{ rows: ListedRow[]; errors?: Listing['errors'] }>;
};

/**
 * Every conversation of the given kinds, as summaries. Unsorted: each tool
 * orders its own result. A team's name comes from its own thread, so the
 * summaries are built against the FULL listing before any filtering. With
 * `unread_only`, only candidates a read confirms are kept.
 */
export async function listSummaries(chatsvc: ChatServiceClient, filter: ListFilter): Promise<ListingWithRetitle> {
  const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
  let rows = raw
    .map((c) => conversationSummary(c, me, raw))
    .filter((s) => filter.kinds.includes(s.kind) && activeWithin(s, filter.activeSince, filter.activeUntil));
  if (filter.unread_only) rows = rows.filter((s) => isUnreadCandidateFor(s, filter.include_hidden));
  let searched: Pick<Listing, 'messages_not_searched'> = {};
  if (filter.query !== undefined) {
    // A query can name a group chat or 1:1 by a participant, so every untitled one is titled first.
    const titled = await enrichTitles(chatsvc, raw, me, rows);
    const matched: ListedRow[] = matchConversations(titled, filter.query, { messages: true });
    if (filter.searchMessages && !namesOne(matched, filter.query)) {
      const found = await searchUnmatched(chatsvc, me, titled, matched, filter.query);
      rows = found.rows;
      if (found.failed > 0) searched = { messages_not_searched: found.failed };
    } else {
      rows = matched;
    }
  }
  if (!filter.unread_only) {
    return { rows, ...searched, retitle: async (kept) => ({ rows: await enrichTitles(chatsvc, raw, me, kept) }) };
  }

  const outcome = await confirmUnreadCandidates(chatsvc, raw, me, rows, {
    limit: CONFIRM_LIMIT,
    maxRaw: CONFIRM_MAX_RAW,
    since: sinceCutoff(filter.since_days),
  });
  const errors = (titled: readonly ConversationSummary[]): Listing['errors'] =>
    outcome.failed.length > 0
      ? outcome.failed.map((f, i) => ({
          conversation: { id: f.summary.id, title: titled[i].title, ...participantsOf(titled[i]) },
          error: wrapApiError(f.error, 'confirm this conversation is unread').message,
        }))
      : undefined;
  const failedSummaries = outcome.failed.map((f) => f.summary);
  const retitle: ListingWithRetitle['retitle'] = async (kept) => {
    const [rows, failed] = await Promise.all([
      enrichTitles(chatsvc, raw, me, kept),
      enrichTitles(chatsvc, raw, me, failedSummaries),
    ]);
    const errs = errors(failed);
    return { rows, ...(errs ? { errors: errs } : {}) };
  };
  const errs = errors(failedSummaries);
  // Listed only on request; titled now, since nothing later trims them.
  const listedOlder = filter.older_unread ? outcome.older : [];
  const olderTitled = await enrichTitles(chatsvc, raw, me, listedOlder.map((o) => o.summary));
  return {
    retitle,
    rows: outcome.confirmed.map(({ summary, read }) => ({
      ...confirmedSummary(summary, true),
      unread_count: read.messages.length,
      ...(read.truncated ? { unread_count_exact: false as const } : {}),
    })),
    skipped_unreachable: outcome.skippedUnreachable,
    ...olderUnreadReport(outcome.older.length, filter.since_days),
    ...olderUnreadConversations(listedOlder, olderTitled),
    ...(errs ? { errors: errs } : {}),
    ...searched,
  };
}

/** Whether the match is an exact id or a self-chat alias — a named chat, not a search. */
function namesOne(matched: readonly ListedRow[], query: string): boolean {
  if (matched.length !== 1) return false;
  const [m] = matched;
  return m.match_reason === 'id' || (m.id === SELF_CHAT && SELF_CHAT_ALIASES.has(query.trim().toLowerCase()));
}

/**
 * `matched` plus every other row in `all` whose recent messages contain the
 * query (`match_reason: "message"`, with `matched_message`), in `all`'s order.
 */
async function searchUnmatched(
  chatsvc: ChatServiceClient,
  me: string,
  all: readonly ConversationSummary[],
  matched: readonly ListedRow[],
  query: string,
): Promise<{ rows: ListedRow[]; failed: number }> {
  const byId = new Map(matched.map((m) => [m.id, m]));
  const rest = all.filter((s) => !byId.has(s.id));
  const { hits, failed } = await searchRecentMessages(chatsvc, me, rest.map((s) => s.id), query);
  const rows = all.flatMap((s): ListedRow[] => {
    const named = byId.get(s.id);
    if (named !== undefined) return [named];
    const hit = hits.get(s.id);
    return hit !== undefined ? [{ ...s, match_reason: 'message', matched_message: hit }] : [];
  });
  return { rows, failed };
}
