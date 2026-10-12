/**
 * The conversation listing shared by `teams_list_chats` and
 * `teams_list_teams_and_channels`: one `listConversations` call, normalised
 * against the signed-in user, then filtered by kind, unread state and query.
 * With `unread_only`, the listing's unread flag only nominates candidates; each
 * is confirmed by reading past its horizon (`confirm.ts`).
 */
import { z } from 'zod';
import type { ChatServiceClient } from '../chatsvc.js';
import { confirmedSummary, confirmUnreadCandidates, olderUnreadReport, sinceCutoff, sinceDaysArg } from './confirm.js';
import { wrapApiError } from './errors.js';
import { enrichTitles } from './titles.js';
import {
  conversationSummary,
  matchConversations,
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
  query?: string;
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

export const unreadOnlyArg = z
  .boolean()
  .optional()
  .describe(
    'Only return conversations with a real message (not a call recording or other system event) ' +
      'the user has not read yet, each with its unread_count (default false). Exact: every candidate ' +
      'is confirmed by reading its messages past the read marker. Never marks anything read.',
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

/** Description text on how an untitled group chat or 1:1 is titled. */
export const GROUP_TITLE_DESCRIPTION =
  'A group chat with no topic is titled "Group chat: " and up to three of its recent senders other ' +
  'than the user, most recent first ("; " between names, "+N" for more), and carries every such ' +
  'recent sender as participants; it is "(untitled group chat)" when no one else wrote recently. ' +
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

export type ListedRow = ConversationSummary & { unread_count?: number; unread_count_exact?: false };

export interface Listing {
  rows: ListedRow[];
  /** With `unread_only`: candidates that answered 404/403 and so are not unread. */
  skipped_unreachable?: number;
  /** With `unread_only`: conversations left out because their unread activity is older than `since_days`. */
  older_unread_skipped?: number;
  older_unread_hint?: string;
  /** With `unread_only`: candidates whose confirming read failed; they are left out of `rows`. */
  errors?: { conversation: { id: string; title: string }; error: string }[];
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
  // A query can name a group chat or 1:1 by a participant, so every untitled one is titled first.
  if (filter.query !== undefined) rows = matchConversations(await enrichTitles(chatsvc, raw, me, rows), filter.query);
  if (!filter.unread_only) return { rows, retitle: async (kept) => ({ rows: await enrichTitles(chatsvc, raw, me, kept) }) };

  const outcome = await confirmUnreadCandidates(chatsvc, raw, me, rows, {
    limit: CONFIRM_LIMIT,
    maxRaw: CONFIRM_MAX_RAW,
    since: sinceCutoff(filter.since_days),
  });
  const errors = (titled: readonly ConversationSummary[]): Listing['errors'] =>
    outcome.failed.length > 0
      ? outcome.failed.map((f, i) => ({
          conversation: { id: f.summary.id, title: titled[i].title },
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
  return {
    retitle,
    rows: outcome.confirmed.map(({ summary, read }) => ({
      ...confirmedSummary(summary, true),
      unread_count: read.messages.length,
      ...(read.truncated ? { unread_count_exact: false as const } : {}),
    })),
    skipped_unreachable: outcome.skippedUnreachable,
    ...olderUnreadReport(outcome.olderSkipped, filter.since_days),
    ...(errs ? { errors: errs } : {}),
  };
}
