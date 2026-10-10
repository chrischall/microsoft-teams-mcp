/**
 * What every tool registrar receives: the one long-lived client, built ONCE in
 * `index.ts` (see the note there on why never inside a registrar).
 *
 * - `chatsvc` — the chat-service REST client every tool reads and writes
 *   through, the Activity feed included (the bridge is touched only to capture
 *   its token).
 *
 * Registrars take a `Pick` of what they actually use, so tests stub only that.
 */
import type { ChatServiceClient } from './chatsvc.js';

export interface TeamsDeps {
  chatsvc: ChatServiceClient;
}
