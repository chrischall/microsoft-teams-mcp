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
 * Every conversation of the given kinds, as summaries. Unsorted: each tool
 * orders its own result. A team's name comes from its own thread, so the
 * summaries are built against the FULL listing before any filtering. With
 * `unread_only`, only candidates a read confirms are kept.
 */
export async function listSummaries(chatsvc: ChatServiceClient, filter: ListFilter): Promise<Listing> {
  const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
  let rows = raw.map((c) => conversationSummary(c, me, raw)).filter((s) => filter.kinds.includes(s.kind));
  if (filter.unread_only) rows = rows.filter((s) => isUnreadCandidateFor(s, filter.include_hidden));
  if (filter.query !== undefined) rows = matchConversations(rows, filter.query);
  if (!filter.unread_only) return { rows };

  const outcome = await confirmUnreadCandidates(chatsvc, raw, me, rows, {
    limit: CONFIRM_LIMIT,
    maxRaw: CONFIRM_MAX_RAW,
    since: sinceCutoff(filter.since_days),
  });
  return {
    rows: outcome.confirmed.map(({ summary, read }) => ({
      ...confirmedSummary(summary, true),
      unread_count: read.messages.length,
      ...(read.truncated ? { unread_count_exact: false as const } : {}),
    })),
    skipped_unreachable: outcome.skippedUnreachable,
    ...olderUnreadReport(outcome.olderSkipped, filter.since_days),
    ...(outcome.failed.length > 0
      ? {
          errors: outcome.failed.map((f) => ({
            conversation: { id: f.summary.id, title: f.summary.title },
            error: wrapApiError(f.error, 'confirm this conversation is unread').message,
          })),
        }
      : {}),
  };
}
