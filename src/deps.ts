/**
 * What every tool registrar receives: the two long-lived clients, built ONCE
 * in `index.ts` (see the note there on why never inside a registrar).
 *
 * - `chatsvc` — the chat-service REST client every chat and channel tool reads
 *   through (the bridge is touched only to capture its token).
 * - `teams` — the DOM-reading bridge client, kept for the Activity feed alone.
 *
 * Registrars take a `Pick` of what they actually use, so tests stub only that.
 */
import type { ChatServiceClient } from './chatsvc.js';
import type { TeamsClient } from './client.js';

export interface TeamsDeps {
  teams: TeamsClient;
  chatsvc: ChatServiceClient;
}
