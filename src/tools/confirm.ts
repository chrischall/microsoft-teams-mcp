/**
 * Confirm-by-fetch: the one rule every unread answer goes through.
 *
 * The conversation listing cannot say whether a conversation is unread. Live
 * (2026-10-10, a real tenant): `lastimreceivedtime` equals the last message's
 * id even when that is a call recording; `lastMessage` can be stale, showing an
 * older real message while only `Event/Call` and `ThreadActivity/*` came after
 * the horizon; and a meeting chat never opened has no horizon at all. Reading
 * the messages past the horizon matched Teams exactly (0 of 32 meeting-chat
 * candidates were really unread). So the listing only nominates candidates
 * (`isUnreadCandidate`), and a candidate counts as unread only when a read past
 * its horizon finds a real message from someone else (`isUnread`). A candidate
 * that answers 404 or 403 — a channel that is gone, or no longer readable — is
 * not unread; it is skipped and counted, never reported as an error.
 *
 * A `since` window (`since_days`, default {@link DEFAULT_SINCE_DAYS}) keeps
 * only conversations whose newest real unread message is inside it. Live
 * (2026-10-10): 14 of 16 confirmed-unread channels had last moved months or
 * years ago, in teams the user never looks at. A candidate whose newest
 * activity of any kind is already older than the window is skipped without a
 * read; both kinds of skip are kept in `older` (counted in `older_unread_skipped`,
 * listed with `older_unread`), never dropped silently.
 */
import { errorStatusOf, mapWithConcurrency } from '@chrischall/mcp-utils';
import type { ChatServiceClient, RawConversation, RawMessage } from '../chatsvc.js';
import { z } from 'zod';
import { isUnread, newestActivityMs, readHorizon, type ConversationSummary } from '../model.js';
import { makeLimitedStopAt } from './limit.js';

/** Candidates read at once. */
export const CONFIRM_CONCURRENCY = 4;

export const DEFAULT_SINCE_DAYS = 14;
export const MAX_SINCE_DAYS = 3650;
const DAY_MS = 86_400_000;

/** The optional `since_days` argument shared by every unread answer. */
export function sinceDaysArg(scope = '') {
  return z
    .number()
    .int()
    .min(1)
    .max(MAX_SINCE_DAYS)
    .optional()
    .describe(
      `${scope}Only count a conversation as unread when its newest unread message arrived in the last ` +
        `since_days days (default ${DEFAULT_SINCE_DAYS}, max ${MAX_SINCE_DAYS}). Older unread ` +
        'conversations are left out and counted in older_unread_skipped; pass a larger since_days ' +
        'to include them.',
    );
}

/** The start of the `since_days` window (epoch ms), measured from now. */
export function sinceCutoff(sinceDays: number | undefined): number {
  return Date.now() - (sinceDays ?? DEFAULT_SINCE_DAYS) * DAY_MS;
}

/**
 * What a result says about conversations the window left out: always the
 * count, and a hint only when something was left out.
 */
export function olderUnreadReport(
  olderSkipped: number,
  sinceDays: number | undefined,
): { older_unread_skipped: number; older_unread_hint?: string } {
  const days = sinceDays ?? DEFAULT_SINCE_DAYS;
  return {
    older_unread_skipped: olderSkipped,
    ...(olderSkipped > 0
      ? {
          older_unread_hint:
            `${olderSkipped} conversation${olderSkipped === 1 ? '' : 's'} with unread activity older than ` +
            `${days} days ${olderSkipped === 1 ? 'was' : 'were'} left out; pass a larger since_days ` +
            `(max ${MAX_SINCE_DAYS}) to include ${olderSkipped === 1 ? 'it' : 'them'}, or older_unread: true ` +
            `to list ${olderSkipped === 1 ? 'it' : 'them'}.`,
        }
      : {}),
  };
}

/** The optional `older_unread` argument: list what the `since_days` window left out. */
export function olderUnreadArg(scope = '') {
  return z
    .boolean()
    .optional()
    .describe(
      `${scope}Also list the conversations since_days left out, as older_unread_conversations: id, ` +
        'kind, title, team (a channel\'s), lastActivity and, where its messages were read, ' +
        'unread_count. One never read (nothing at all happened inside the window) carries ' +
        'unread_count_known: false instead and is not read now — its news may be only system ' +
        'events. Default false.',
    );
}

/** One conversation the `since` window left out, as `older_unread_conversations` lists it. */
export interface OlderUnreadRow {
  id: string;
  kind: ConversationSummary['kind'];
  title: string;
  team?: string;
  lastActivity?: string;
  unread_count?: number;
  /** The read was cut short, so `unread_count` is a lower bound. */
  unread_count_exact?: false;
  /** Never read: whether it holds a real unread message, and how many, is not known. */
  unread_count_known?: false;
}

/**
 * `older_unread_conversations` from the left-out conversations, given their
 * (re)titled summaries in the same order: most recent activity first, and
 * absent when nothing was left out.
 */
export function olderUnreadConversations(
  older: readonly OlderSkip[],
  titled: readonly ConversationSummary[],
): { older_unread_conversations?: OlderUnreadRow[] } {
  if (older.length === 0) return {};
  const rows = older.map(({ read }, i): OlderUnreadRow => {
    const s = titled[i];
    return {
      id: s.id,
      kind: s.kind,
      title: s.title,
      ...(s.team !== undefined ? { team: s.team } : {}),
      ...(s.lastActivity !== undefined ? { lastActivity: s.lastActivity } : {}),
      ...(read === undefined
        ? { unread_count_known: false as const }
        : { unread_count: read.messages.length, ...(read.truncated ? { unread_count_exact: false as const } : {}) }),
    };
  });
  rows.sort((a, b) => (b.lastActivity ?? '').localeCompare(a.lastActivity ?? ''));
  return { older_unread_conversations: rows };
}

/** A 404 or 403: the conversation is gone or not readable, so it is not unread. */
export function isUnreachable(err: unknown): boolean {
  const status = errorStatusOf(err);
  return status === 404 || status === 403;
}

export interface UnreadRead {
  /** The real unread messages read, newest first (at most `limit`). */
  messages: RawMessage[];
  /** More unread messages may exist than were read (the limit or `maxRaw` cut the read). */
  truncated: boolean;
}

/**
 * Read a conversation's messages past `horizon`, newest first, keeping only
 * the real unread ones (`isUnread`). Stops at the horizon, at the first real
 * unread message past `limit`, or after `maxRaw` raw messages.
 */
export async function readUnread(
  chatsvc: ChatServiceClient,
  id: string,
  horizon: number,
  me: string,
  { limit, maxRaw }: { limit: number; maxRaw: number },
): Promise<UnreadRead> {
  const wanted = (m: RawMessage): boolean => isUnread(m, horizon, me);
  const limited = makeLimitedStopAt(horizon, wanted, limit);
  let capped = false;
  const fetched = await chatsvc.getMessages(id, {
    stopAt: limited.stopAt,
    max: maxRaw,
    onCapped: () => (capped = true),
  });
  return { messages: fetched.filter(wanted), truncated: limited.truncated || capped };
}

export interface Confirmed {
  summary: ConversationSummary;
  read: UnreadRead;
}

/** A candidate the `since` window left out, with its read when there was one. */
export interface OlderSkip {
  summary: ConversationSummary;
  /** Absent when nothing inside the window happened, so it was never read. */
  read?: UnreadRead;
}

export interface ConfirmOutcome {
  /** Candidates holding at least one real unread message, in candidate order. */
  confirmed: Confirmed[];
  /** Candidates that answered 404/403 — not unread. */
  skippedUnreachable: number;
  /**
   * Candidates left out by the `since` window, in candidate order: newest
   * activity older than it (never read), or newest real unread message older
   * than it (read).
   */
  older: OlderSkip[];
  /** Candidates whose read failed some other way, in candidate order. */
  failed: { summary: ConversationSummary; error: unknown }[];
}

/**
 * Read each candidate past its own horizon, {@link CONFIRM_CONCURRENCY} at a
 * time, and keep those with a real unread message. One read failing never
 * costs the others; when every read failed (none confirmed, none came back
 * empty), the first failure is thrown — it is a failure of the call (a revoked
 * token, a rate limit), not of one conversation.
 *
 * With `since` (epoch ms), a candidate whose newest activity predates it is
 * not read at all, and one whose newest real unread message predates it is
 * not kept; both go in `older`.
 */
export async function confirmUnreadCandidates(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  candidates: readonly ConversationSummary[],
  opts: { limit: number; maxRaw: number; since?: number },
): Promise<ConfirmOutcome> {
  const byId = new Map(raw.map((c) => [c.id, c]));
  const since = opts.since ?? 0;
  type Result =
    | { kind: 'read'; read: UnreadRead }
    | { kind: 'older' }
    | { kind: 'unreachable' }
    | { kind: 'failed'; error: unknown };
  const results = await mapWithConcurrency(candidates, CONFIRM_CONCURRENCY, async (summary): Promise<Result> => {
    const conv = byId.get(summary.id) ?? { id: summary.id };
    const newest = newestActivityMs(conv);
    // Nothing at all happened inside the window, so no unread message can be in it.
    if (newest !== undefined && newest < since) return { kind: 'older' };
    const horizon = readHorizon(conv);
    try {
      return { kind: 'read', read: await readUnread(chatsvc, summary.id, horizon, me, opts) };
    } catch (error) {
      return isUnreachable(error) ? { kind: 'unreachable' } : { kind: 'failed', error };
    }
  });

  const outcome: ConfirmOutcome = { confirmed: [], skippedUnreachable: 0, older: [], failed: [] };
  let anyRead = false;
  results.forEach((r, i) => {
    const summary = candidates[i];
    if (r.kind === 'older') outcome.older.push({ summary });
    else if (r.kind === 'unreachable') outcome.skippedUnreachable++;
    else if (r.kind === 'failed') outcome.failed.push({ summary, error: r.error });
    else {
      anyRead = true;
      // Read newest first, so the first message is the newest real unread one.
      const newestUnread = r.read.messages[0];
      if (newestUnread === undefined) return;
      if (Number(newestUnread.id) < since) outcome.older.push({ summary, read: r.read });
      else outcome.confirmed.push({ summary, read: r.read });
    }
  });
  const first = outcome.failed[0];
  if (first !== undefined && !anyRead) throw first.error;
  return outcome;
}

/** A summary whose unread state a read has settled. */
export function confirmedSummary(summary: ConversationSummary, unread: boolean): ConversationSummary {
  return { ...summary, unread, unread_confirmed: true };
}
