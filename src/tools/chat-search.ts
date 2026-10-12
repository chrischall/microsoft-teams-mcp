/**
 * `teams_list_chats`' opt-in `search_messages`: a query matched against what
 * was said in each chat, not only what the chat is called.
 *
 * The chat service has no search call this server can reach, so each chat no
 * name or last-message preview matched is read here — one page of its
 * {@link SEARCH_FETCH_MESSAGES} most recent messages, {@link SEARCH_CONCURRENCY}
 * chats at a time, the same bounded fan-out `titles.ts` uses. Only messages a
 * person wrote count (not deleted ones or system events), matched on their
 * plain text. A read that fails leaves its chat out and is counted — it never
 * fails the call. Nothing is cached: a search wants the newest messages.
 */
import { mapWithConcurrency } from '@chrischall/mcp-utils';
import type { ChatServiceClient } from '../chatsvc.js';
import { isUserMessage, normaliseMessage } from '../model.js';

/** Recent messages read per chat — one page. */
export const SEARCH_FETCH_MESSAGES = 50;
/** Chats read at once. */
export const SEARCH_CONCURRENCY = 4;
/** Longest matched-message text returned, centred on the match. */
const SNIPPET_CHARS = 200;
/** Context kept before the match when the text has to be cut. */
const SNIPPET_LEAD = 60;

/** The newest message in a chat that matched, as `matched_message`. */
export interface MessageHit {
  id: string;
  time: string;
  sender: string;
  /** Its plain text, cut to a window around the match when long. */
  text: string;
}

export interface MessageSearch {
  /** The chats with a matching recent message, by id, each with its newest match. */
  hits: Map<string, MessageHit>;
  /** Chats whose read failed, so only their names and preview were matched. */
  failed: number;
}

/**
 * Read the recent messages of each chat in `ids` and keep, per chat, the
 * newest whose text contains `query` (case-insensitive).
 */
export async function searchRecentMessages(
  chatsvc: ChatServiceClient,
  me: string,
  ids: readonly string[],
  query: string,
): Promise<MessageSearch> {
  const q = query.trim().toLowerCase();
  const hits = new Map<string, MessageHit>();
  let failed = 0;
  if (q === '') return { hits, failed };
  await mapWithConcurrency([...new Set(ids)], SEARCH_CONCURRENCY, async (id) => {
    let messages;
    try {
      messages = await chatsvc.getMessages(id, { max: SEARCH_FETCH_MESSAGES, pageSize: SEARCH_FETCH_MESSAGES });
    } catch {
      failed++;
      return;
    }
    // Served newest first, so the first match is the newest.
    for (const m of messages) {
      if (!isUserMessage(m)) continue;
      const n = normaliseMessage(m, me);
      if (n.deleted) continue;
      const at = n.text.toLowerCase().indexOf(q);
      if (at === -1) continue;
      hits.set(id, { id: n.id, time: n.time, sender: n.sender, text: snippet(n.text, at) });
      return;
    }
  });
  return { hits, failed };
}

/** `text` whole when short, else a window of it that starts a little before index `at`. */
function snippet(text: string, at: number): string {
  if (text.length <= SNIPPET_CHARS) return text;
  const start = Math.max(0, Math.min(at - SNIPPET_LEAD, text.length - SNIPPET_CHARS));
  const end = start + SNIPPET_CHARS;
  return `${start > 0 ? '…' : ''}${text.slice(start, end)}${end < text.length ? '…' : ''}`;
}
