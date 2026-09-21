/**
 * Full-bridge Teams client.
 *
 * Unlike Outlook's bootstrap-capture archetype, this one touches the bridge
 * on every call. Teams Web's message data lives only in a client-side sync
 * cache fed by a WebSocket/registrar channel — there is no REST or GraphQL
 * endpoint that returns it, so the only way to read it is to read the DOM the
 * page has already rendered. That rules out a captured-token-then-plain-fetch
 * design: every read here goes through `read_dom_list` against the user's
 * live, signed-in `teams.microsoft.com` tab.
 *
 * Consequence: this MCP can only read whichever chat is CURRENTLY DISPLAYED
 * in that tab. `@fetchproxy/server` has no capability to navigate a tab to a
 * different conversation — every declared capability is a fetch or a read,
 * never a page interaction — so there is no way to ask for "chat X's
 * messages" on demand. `listChats()` reads the chat-list sidebar (always
 * visible); `getOpenChatMessages()` reads whatever conversation the user has
 * open. Verified live 2026-09-21 against a real Teams tenant.
 */
import {
  createFetchproxyTransport,
  type FetchproxyTransport,
} from '@chrischall/mcp-utils/fetchproxy';
import { readPortEnv } from '@chrischall/mcp-utils';
import type { DomListSelectorDecl } from '@fetchproxy/protocol';
import { PACKAGE_NAME, VERSION } from './version.js';

/**
 * The two hosts Teams Web actually serves from. A tenant's tab may be on
 * either; `readDomListAnyDomain` tries both in order and falls through to
 * the next on a "no tab matching" failure, since only one will ever have a
 * live tab.
 */
export const DOMAINS = ['teams.microsoft.com', 'teams.cloud.microsoft'] as const;

/**
 * fetchproxy concentrator port — one for the whole fleet. `TEAMS_WS_PORT`
 * overrides it for local development and for a hosted bridged registration
 * (where mcp-host names this variable in `bridgePortEnv`).
 */
const DEFAULT_WS_PORT = 37_149;
export function getWsPort(): number {
  return readPortEnv('TEAMS_WS_PORT', DEFAULT_WS_PORT);
}

/**
 * The currently-open chat's message thread. `.fui-ChatMessage` (others'
 * messages) and `.fui-ChatMyMessage` (the signed-in user's own) are Teams
 * v2's two message-row classes; both carry the same `message-author-name` /
 * `time` / `chat-pane-message` structure. `sender` reads back even for a
 * grouped/consecutive message Teams visually hides the name on — it stays
 * in the DOM, just visually collapsed. A quoted-reply message's `text`
 * includes the quoted preview concatenated in (a known limitation of bulk
 * `.textContent` extraction, not fixable without a second, narrower field).
 */
export const CHAT_MESSAGES_SELECTOR: DomListSelectorDecl = {
  name: 'chatMessages',
  itemSelector: '.fui-ChatMessage, .fui-ChatMyMessage',
  fields: [
    { name: 'sender', selector: '[data-tid="message-author-name"]' },
    { name: 'time', selector: 'time', attribute: 'datetime' },
    { name: 'messageId', selector: 'time', attribute: 'id' },
    { name: 'text', selector: '[data-tid="chat-pane-message"]' },
  ],
  maxItems: 200,
};

/**
 * The chat-list sidebar (always rendered, regardless of which chat is
 * open). `conversationKey` is Teams' own compound tree-item value — it
 * embeds the real thread id (e.g. `19:<id>@thread.v2`) but is NOT usable to
 * open that chat from here; there is no navigate capability. It is exposed
 * so a caller can at least correlate a list entry with the currently-open
 * chat's messages.
 */
export const CHAT_LIST_SELECTOR: DomListSelectorDecl = {
  name: 'chatList',
  itemSelector: '[data-testid="list-item"]',
  fields: [
    { name: 'title', selector: '[id^="title-chat-list-item_"]' },
    { name: 'preview', selector: '[data-testid="comfy-message"]' },
    { name: 'time', selector: '[id^="time-chat-list-item_"]' },
    { name: 'conversationKey', attribute: 'data-fui-tree-item-value' },
    { name: 'itemType', attribute: 'data-item-type' },
  ],
  maxItems: 200,
};

export interface ChatMessageRow {
  sender?: string;
  /** ISO 8601 (the `<time datetime>` attribute). */
  time?: string;
  messageId?: string;
  text?: string;
}

export interface ChatListRow {
  title?: string;
  preview?: string;
  time?: string;
  conversationKey?: string;
  itemType?: string;
}

export interface TeamsClientOptions {
  /** Injected for tests so no suite ever touches the real bridge. */
  transport?: FetchproxyTransport;
}

export class TeamsClient {
  readonly #transport: FetchproxyTransport;
  #started = false;

  constructor(opts: TeamsClientOptions = {}) {
    this.#transport =
      opts.transport ??
      createFetchproxyTransport({
        serverName: PACKAGE_NAME,
        version: VERSION,
        domains: [...DOMAINS],
        capabilities: ['read_dom_list'],
        domListSelectors: [CHAT_MESSAGES_SELECTOR, CHAT_LIST_SELECTOR],
        port: getWsPort(),
        // stderr only — stdout is the JSON-RPC channel. Without this the
        // first-ever pairing (or a scope widening) leaves a tool call
        // hanging with no way for the user to know a Transporter approval
        // is what it's waiting on.
        onPairCode: (code) => {
          console.error(
            `[microsoft-teams-mcp] fetchproxy pair code: ${code} — approve in the Transporter extension popup`,
          );
        },
      });
  }

  /**
   * `start()` only loads the identity keypair — it does not bind the port or
   * dial the extension, which stays lazy until the first verb. Called once,
   * lazily, so a server that never gets a tool call never touches the
   * bridge at all.
   */
  async #ensureStarted(): Promise<void> {
    if (this.#started) return;
    await this.#transport.start();
    this.#started = true;
  }

  /**
   * Try each declared domain in turn, falling through to the next on a
   * "no tab matching" failure. Only one domain will ever have a live tab —
   * `readDomList` (unlike `captureRequestHeader`) is a point-in-time DOM
   * read with no open capture window to race, so sequential try-then-fall-
   * through is correct here (not `Promise.any`).
   */
  async #readDomListAnyDomain(name: string): Promise<Record<string, string>[]> {
    await this.#ensureStarted();
    const errors: string[] = [];
    for (const domain of DOMAINS) {
      try {
        return await this.#transport.server.readDomList({ name, domain });
      } catch (e) {
        const err = e as Error & { name?: string; hint?: string };
        errors.push(`${domain}: ${err.message}`);
        // Only a "no tab on this domain" failure is worth trying the next
        // domain for. Anything else (scope rejection, bridge down) is the
        // same for every domain and should surface immediately.
        if (!/no tab matching/i.test(err.message)) throw e;
      }
    }
    throw new Error(
      `could not reach a signed-in Teams tab on any of [${DOMAINS.join(', ')}] (${errors.join('; ')}). ` +
        'Open teams.microsoft.com in a signed-in browser tab and retry.',
    );
  }

  /** The chat-list sidebar, regardless of which chat is currently open. */
  async listChats(): Promise<ChatListRow[]> {
    return (await this.#readDomListAnyDomain('chatList')) as ChatListRow[];
  }

  /**
   * Messages from whichever chat is CURRENTLY DISPLAYED in the signed-in
   * tab. There is no way to select a different chat from here — see the
   * module doc.
   */
  async getOpenChatMessages(): Promise<ChatMessageRow[]> {
    return (await this.#readDomListAnyDomain('chatMessages')) as ChatMessageRow[];
  }

  async close(): Promise<void> {
    await this.#transport.close();
    this.#started = false;
  }

  /** Non-secret bridge status for the healthcheck tool. */
  status(): unknown {
    return this.#transport.status();
  }

  /** The underlying bridge transport, for `registerBridgeHealthcheckTool`. */
  get transport(): FetchproxyTransport {
    return this.#transport;
  }
}
