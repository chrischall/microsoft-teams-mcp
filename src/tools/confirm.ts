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
 * read; both kinds of skip are counted in `olderSkipped`, never dropped silently.
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
            `(max ${MAX_SINCE_DAYS}) to include ${olderSkipped === 1 ? 'it' : 'them'}.`,
        }
      : {}),
  };
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

export interface ConfirmOutcome {
  /** Candidates holding at least one real unread message, in candidate order. */
  confirmed: Confirmed[];
  /** Candidates that answered 404/403 — not unread. */
  skippedUnreachable: number;
  /**
   * Candidates left out by the `since` window: newest activity older than it
   * (never read), or newest real unread message older than it.
   */
  olderSkipped: number;
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
 * not kept; both are counted in `olderSkipped`.
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

  const outcome: ConfirmOutcome = { confirmed: [], skippedUnreachable: 0, olderSkipped: 0, failed: [] };
  let anyRead = false;
  results.forEach((r, i) => {
    const summary = candidates[i];
    if (r.kind === 'older') outcome.olderSkipped++;
    else if (r.kind === 'unreachable') outcome.skippedUnreachable++;
    else if (r.kind === 'failed') outcome.failed.push({ summary, error: r.error });
    else {
      anyRead = true;
      // Read newest first, so the first message is the newest real unread one.
      const newestUnread = r.read.messages[0];
      if (newestUnread === undefined) return;
      if (Number(newestUnread.id) < since) outcome.olderSkipped++;
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
