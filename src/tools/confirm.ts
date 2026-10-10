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
 */
import { errorStatusOf, mapWithConcurrency } from '@chrischall/mcp-utils';
import type { ChatServiceClient, RawConversation, RawMessage } from '../chatsvc.js';
import { isUnread, readHorizon, type ConversationSummary } from '../model.js';
import { makeLimitedStopAt } from './limit.js';

/** Candidates read at once. */
export const CONFIRM_CONCURRENCY = 4;

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
  /** Candidates whose read failed some other way, in candidate order. */
  failed: { summary: ConversationSummary; error: unknown }[];
}

/**
 * Read each candidate past its own horizon, {@link CONFIRM_CONCURRENCY} at a
 * time, and keep those with a real unread message. One read failing never
 * costs the others; when every read failed (none confirmed, none came back
 * empty), the first failure is thrown — it is a failure of the call (a revoked
 * token, a rate limit), not of one conversation.
 */
export async function confirmUnreadCandidates(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  candidates: readonly ConversationSummary[],
  opts: { limit: number; maxRaw: number },
): Promise<ConfirmOutcome> {
  const byId = new Map(raw.map((c) => [c.id, c]));
  type Result = { kind: 'read'; read: UnreadRead } | { kind: 'unreachable' } | { kind: 'failed'; error: unknown };
  const results = await mapWithConcurrency(candidates, CONFIRM_CONCURRENCY, async (summary): Promise<Result> => {
    const horizon = readHorizon(byId.get(summary.id) ?? { id: summary.id });
    try {
      return { kind: 'read', read: await readUnread(chatsvc, summary.id, horizon, me, opts) };
    } catch (error) {
      return isUnreachable(error) ? { kind: 'unreachable' } : { kind: 'failed', error };
    }
  });

  const outcome: ConfirmOutcome = { confirmed: [], skippedUnreachable: 0, failed: [] };
  let anyRead = false;
  results.forEach((r, i) => {
    const summary = candidates[i];
    if (r.kind === 'unreachable') outcome.skippedUnreachable++;
    else if (r.kind === 'failed') outcome.failed.push({ summary, error: r.error });
    else {
      anyRead = true;
      if (r.read.messages.length > 0) outcome.confirmed.push({ summary, read: r.read });
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
