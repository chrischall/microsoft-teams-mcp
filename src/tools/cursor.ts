/**
 * Message-id cursors: "only what is new since I last looked".
 *
 * A chat-service message id IS its arrival time in epoch ms (`arrivalMs`), so
 * ids order messages, and "newer than the last one I saw" is a plain numeric
 * comparison — a stop boundary for the newest-first read, exactly like the read
 * horizon. The cursor handed back (`next_cursor`) is just the newest returned
 * message's id; the caller passes it back as `after_message_id` (or, for
 * `teams_get_unread`, inside `cursors`). The server keeps no state and no
 * cache: every call reads live, and says when in `fetched_at`.
 */
import { z } from 'zod';
import type { RawMessage } from '../chatsvc.js';

/** A message id: decimal digits (its arrival epoch ms). */
const messageId = z.string().regex(/^\d+$/, { message: 'a message id is decimal digits, e.g. "1760000000000"' });

export const afterMessageIdArg = messageId
  .optional()
  .describe(
    'Only return messages newer than this message id (a previous result\'s next_cursor, or any message id); ' +
      'the read stops at the first message at or before it.',
  );

export const cursorsArg = z
  .record(z.string(), messageId)
  .optional()
  .describe(
    'Per-conversation cursors, { conversationId: messageId } (a previous result\'s next_cursors): each ' +
      "conversation's read stops at the later of its read marker and its cursor, so only unread messages " +
      'newer than the cursor come back, and a conversation with none past it is left out.',
  );

/** Shared description text: what next_cursor is and that nothing is cached. */
export const CURSOR_NO_CACHE_DESCRIPTION =
  'fetched_at is when this call read Teams: the server holds no cache, so every call reads live ' +
  '(there is nothing to refresh). ';

/** The id of the newest of `messages` (ids are arrival ms), or undefined when there are none. */
export function newestId(messages: readonly Pick<RawMessage, 'id'>[]): string | undefined {
  let newest: string | undefined;
  for (const m of messages) if (newest === undefined || Number(m.id) > Number(newest)) newest = m.id;
  return newest;
}
