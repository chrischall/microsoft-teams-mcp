/**
 * Channel thread roots that lie outside a read's window.
 *
 * A channel read groups replies under their root post (`format.ts`), but the
 * window — `since`, `limit`, a cursor, or the read horizon for unread — often
 * ends after a thread's root while its newer replies are inside. Each such
 * root is fetched by id (`GET …/conversations/{channelId}/messages/{rootId}`,
 * verified live 2026-10-12: it answers for a root weeks older than the window,
 * in ~250 ms) so the thread keeps its subject, time and sender. A fetched root
 * is context, not part of the read: it is shown marked `outside_window` and
 * counts toward no unread count, message count or cursor.
 *
 * Fetches are bounded per CALL — {@link MAX_ROOT_FETCHES}, shared by every
 * conversation the call reads, newest threads first — and a few run at once.
 * A root past the budget, gone, or unreadable (the service answers 403, not
 * 404, for an id it does not hold) stays the `{ id, missing: true }`
 * placeholder.
 */
import { mapWithConcurrency } from '@chrischall/mcp-utils';
import type { ChatServiceClient } from '../chatsvc.js';
import { normaliseMessage, type NormalisedMessage } from '../model.js';

/** Root fetches one tool call may make, across every conversation it reads. */
export const MAX_ROOT_FETCHES = 25;

/** Root fetches in flight at once. */
export const ROOT_FETCH_CONCURRENCY = 5;

/** What is left of one call's {@link MAX_ROOT_FETCHES}. */
export interface RootBudget {
  remaining: number;
}

export function rootBudget(): RootBudget {
  return { remaining: MAX_ROOT_FETCHES };
}

/**
 * The roots of `messages` (oldest first, as normalised) that they do not
 * themselves hold, newest thread first: the ones worth fetching first when the
 * budget runs short.
 */
function outsideRootIds(messages: readonly NormalisedMessage[]): string[] {
  const held = new Set(messages.map((m) => m.id));
  const ids = new Set<string>();
  for (let i = messages.length - 1; i >= 0; i--) {
    const root = messages[i].replyTo;
    if (root !== undefined && !held.has(root)) ids.add(root);
  }
  return [...ids];
}

/**
 * Fetch the roots a channel read's replies point at but its window does not
 * hold, spending from `budget`. Returns those found, by id, normalised as the
 * read's own messages are; a root that failed, came back as some other
 * message, or is itself a reply is left out (it stays missing).
 */
export async function fetchOutsideRoots(
  chatsvc: ChatServiceClient,
  conversationId: string,
  messages: readonly NormalisedMessage[],
  me: string,
  { budget, timeZone }: { budget: RootBudget; timeZone?: string },
): Promise<Map<string, NormalisedMessage>> {
  const wanted = outsideRootIds(messages);
  // Taken synchronously, so conversations read concurrently cannot overspend it.
  const ids = wanted.slice(0, Math.max(0, budget.remaining));
  budget.remaining -= ids.length;
  const roots = new Map<string, NormalisedMessage>();
  if (ids.length === 0) return roots;
  const fetched = await mapWithConcurrency(ids, ROOT_FETCH_CONCURRENCY, async (id) => {
    try {
      return await chatsvc.getMessage(conversationId, id);
    } catch {
      return undefined;
    }
  });
  fetched.forEach((raw, i) => {
    if (raw === undefined || raw.id !== ids[i]) return;
    const root = normaliseMessage(raw, me, { timeZone });
    if (root.replyTo === undefined) roots.set(root.id, root);
  });
  return roots;
}
