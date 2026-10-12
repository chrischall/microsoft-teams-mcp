/**
 * What every tool registrar receives: the one long-lived client, built ONCE in
 * `index.ts` (see the note there on why never inside a registrar).
 *
 * - `chatsvc` — the chat-service REST client every tool reads and writes
 *   through, the Activity feed included (the bridge is touched only to capture
 *   its token).
 * - `sharepoint` — the SharePoint download client `teams_get_attachment` uses
 *   to fetch a shared file's contents (a separate, host-scoped, short-lived
 *   token from chatsvc's; the bridge is touched only to capture it).
 * - `substrate` — the Microsoft Substrate search client `teams_search_messages`
 *   uses first (a third token, captured on demand off the Teams tab's own search
 *   request; the tool falls back to scanning the chat service without it).
 *
 * Registrars take a `Pick` of what they actually use, so tests stub only that.
 */
import type { ChatServiceClient } from './chatsvc.js';
import type { SharePointClient } from './sharepoint.js';
import type { SubstrateSearch } from './substrate.js';

export interface TeamsDeps {
  chatsvc: ChatServiceClient;
  sharepoint: SharePointClient;
  substrate: SubstrateSearch;
}
