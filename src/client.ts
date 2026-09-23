/**
 * Full-bridge Teams client.
 *
 * Unlike Outlook's bootstrap-capture archetype, this one touches the bridge
 * on every call. Teams Web's message data lives only in a client-side sync
 * cache fed by a WebSocket/registrar channel — there is no REST or GraphQL
 * endpoint that returns it, so the only way to read it is to read the DOM the
 * page has already rendered. That rules out a captured-token-then-plain-fetch
 * design: every read here goes through `read_dom_list` against the user's
 * live, signed-in `teams.cloud.microsoft` tab.
 *
 * Consequence: this MCP can only read whichever chat or channel is
 * CURRENTLY DISPLAYED in that tab. `@fetchproxy/server` has no capability to
 * navigate a tab to a different conversation — every declared capability is
 * a fetch or a read, never a page interaction — so there is no way to ask
 * for "chat X's messages" or "channel Y's posts" on demand.
 * `listChats()`/`listTeamsAndChannels()` read their sidebars (always
 * visible once that nav section is open); `getOpenChatMessages()`/
 * `getOpenChannelPosts()` read whatever conversation the user has open.
 * Verified live 2026-09-21 against a real Teams tenant.
 */
import {
  createFetchproxyTransport,
  type FetchproxyTransport,
} from '@chrischall/mcp-utils/fetchproxy';
import { readPortEnv } from '@chrischall/mcp-utils';
import type { DomListSelectorDecl } from '@fetchproxy/protocol';
import { PACKAGE_NAME, VERSION } from './version.js';

/**
 * The New Teams host this MCP reads from. `teams.microsoft.com` (the old
 * host) is deliberately not supported — only `teams.cloud.microsoft` tabs
 * are read.
 */
export const DOMAINS = ['teams.cloud.microsoft'] as const;

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

/**
 * The Teams-and-Channels sidebar (the left rail under the "Teams" nav icon,
 * a DIFFERENT view from Chat — always rendered once that view is open,
 * regardless of which channel is open). Team and channel rows share one
 * item selector; `title-*`/`time-*` use a compound attribute selector
 * (`[id^="title-"][id*="-list-item-"]`) rather than a comma-joined pair of
 * prefixes, because a channel and a team item use different id prefixes
 * (`title-channel-list-item-*` / `title-team-list-item-*`) and this reads
 * either without needing two alternatives. `teamName` is the parent team's
 * display name shown under a channel row (absent on a team row itself).
 */
export const TEAMS_AND_CHANNELS_SELECTOR: DomListSelectorDecl = {
  name: 'teamsAndChannels',
  itemSelector: '[data-item-type="team"], [data-item-type="channel"]',
  fields: [
    { name: 'title', selector: '[id^="title-"][id*="-list-item-"]' },
    { name: 'teamName', selector: '[id^="preview-channel-list-item-"]' },
    { name: 'time', selector: '[id^="time-"][id*="-list-item-"]' },
    { name: 'conversationKey', attribute: 'data-fui-tree-item-value' },
    { name: 'itemType', attribute: 'data-item-type' },
  ],
  maxItems: 300,
};

/**
 * Top-level posts in whichever CHANNEL is currently open (same "reads only
 * what's displayed" constraint as chat messages — see module doc). Channel
 * posts use a different renderer than chat messages (`channel-pane-message`,
 * not `.fui-ChatMessage`), and their `<time>` element carries no ISO
 * `datetime` attribute — only a human-readable `aria-label` — so `time` here
 * is prose ("Tuesday, August 4, 2026 7:50 AM"), not ISO 8601 like
 * `ChatMessageRow.time`. `subject` is present only on a post that has a
 * title; a reply-shaped post has none. Threaded REPLIES under a post are not
 * captured — only the top-level posts a channel's "Posts" tab lists.
 */
export const CHANNEL_POSTS_SELECTOR: DomListSelectorDecl = {
  name: 'channelPosts',
  itemSelector: '[data-tid="channel-pane-message"]',
  fields: [
    { name: 'sender', selector: '[id^="author-"]' },
    { name: 'subject', selector: '[id^="subject-line-"]' },
    { name: 'time', selector: 'time[data-tid="timestamp"]', attribute: 'aria-label' },
    { name: 'messageId', selector: 'time[data-tid="timestamp"]', attribute: 'id' },
    { name: 'text', selector: '[id^="content-"]' },
  ],
  maxItems: 200,
};

/**
 * Top-level rows in the Activity feed (the bell icon in the left nav) —
 * mentions, replies, reactions, and the like across every chat and channel.
 * Always rendered once that nav section is open, regardless of which item
 * is selected — same "sidebar, not detail pane" shape as `CHAT_LIST_SELECTOR`
 * and `TEAMS_AND_CHANNELS_SELECTOR`. `location` is the team/channel or chat
 * the activity happened in, prose-joined by Teams itself (e.g. "NS_MS
 * Product CE > MS DevOps"). `time` is prose (e.g. "1:19 PM"), not ISO 8601 —
 * Teams renders it that way here, unlike `chatMessages.time`.
 */
export const ACTIVITY_FEED_SELECTOR: DomListSelectorDecl = {
  name: 'activityFeed',
  itemSelector: '[data-tid="activity-feed-list-item"]',
  fields: [
    { name: 'title', selector: '[id^="activity-feed-item-title-"]' },
    { name: 'preview', selector: '[id^="activity-feed-item-message-preview-"]' },
    { name: 'time', selector: '[id^="activity-feed-item-timestamp-"]' },
    { name: 'location', selector: '[id^="activity-feed-item-location-"]' },
  ],
  maxItems: 200,
};

/**
 * WHICH conversation is open — the sidebar row Teams marks as selected. Read
 * alongside `chatMessages`/`channelPosts` so the "open" tools can say which
 * chat or channel their rows came from; without it the model gets bare rows
 * and can misattribute another conversation's messages to the one the user
 * asked about. The chat-list and channel-list rows are Fluent UI tree items
 * (`data-fui-tree-item-value`, the same attribute `CHAT_LIST_SELECTOR` /
 * `TEAMS_AND_CHANNELS_SELECTOR` read `conversationKey` from), and a Fluent
 * tree item carries `aria-selected="true"` (or `aria-current`) when it is the
 * open one. `title` matches both the chat (`title-chat-list-item_*`) and
 * channel (`title-channel-list-item-*`) title ids. `conversationKey` embeds
 * the real thread id. Not yet re-verified against a live tenant — the tools
 * treat an empty read as "could not identify" rather than failing.
 */
export const OPEN_CONVERSATION_SELECTOR: DomListSelectorDecl = {
  name: 'openConversation',
  itemSelector:
    '[data-fui-tree-item-value][aria-selected="true"], ' +
    '[data-fui-tree-item-value][aria-current="true"], ' +
    '[data-fui-tree-item-value][aria-current="page"]',
  fields: [
    { name: 'title', selector: '[id^="title-"][id*="list-item"]' },
    { name: 'conversationKey', attribute: 'data-fui-tree-item-value' },
    { name: 'itemType', attribute: 'data-item-type' },
  ],
  maxItems: 5,
};

/** Every selector the client declares to the bridge (the pair-popup scope). */
export const DOM_LIST_SELECTORS: readonly DomListSelectorDecl[] = [
  CHAT_MESSAGES_SELECTOR,
  CHAT_LIST_SELECTOR,
  TEAMS_AND_CHANNELS_SELECTOR,
  CHANNEL_POSTS_SELECTOR,
  ACTIVITY_FEED_SELECTOR,
  OPEN_CONVERSATION_SELECTOR,
];

export interface OpenConversationRow {
  title?: string;
  /** Teams' compound tree-item value; embeds the thread id. */
  conversationKey?: string;
  /** `'chat'`, `'channel'`, … as Teams labels the sidebar row. */
  itemType?: string;
}

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

export interface TeamOrChannelRow {
  title?: string;
  /** The parent team's display name. Present on a channel row, absent on a team row. */
  teamName?: string;
  time?: string;
  conversationKey?: string;
  /** `'team'` or `'channel'`. */
  itemType?: string;
}

export interface ActivityFeedRow {
  title?: string;
  preview?: string;
  /** Prose, e.g. "1:19 PM" — NOT ISO 8601. */
  time?: string;
  /** The team/channel or chat the activity happened in, prose-joined by Teams. */
  location?: string;
}

export interface ChannelPostRow {
  sender?: string;
  subject?: string;
  /** Prose, e.g. "Tuesday, August 4, 2026 7:50 AM" — NOT ISO 8601. */
  time?: string;
  messageId?: string;
  text?: string;
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
        domListSelectors: [...DOM_LIST_SELECTORS],
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

  /** Reads a declared selector off the single supported domain's live tab. */
  async #readDomList(name: string): Promise<Record<string, string>[]> {
    await this.#ensureStarted();
    try {
      return await this.#transport.server.readDomList({ name, domain: DOMAINS[0] });
    } catch (e) {
      const err = e as Error;
      if (/no tab matching/i.test(err.message)) {
        throw new Error(
          `could not reach a signed-in Teams tab on ${DOMAINS[0]} (${err.message}). ` +
            `Open ${DOMAINS[0]} in a signed-in browser tab and retry.`,
        );
      }
      throw e;
    }
  }

  /** The chat-list sidebar, regardless of which chat is currently open. */
  async listChats(): Promise<ChatListRow[]> {
    return (await this.#readDomList('chatList')) as ChatListRow[];
  }

  /**
   * Messages from whichever chat is CURRENTLY DISPLAYED in the signed-in
   * tab. There is no way to select a different chat from here — see the
   * module doc.
   */
  async getOpenChatMessages(): Promise<ChatMessageRow[]> {
    return (await this.#readDomList('chatMessages')) as ChatMessageRow[];
  }

  /** The Teams-and-Channels sidebar, regardless of which channel is open. */
  async listTeamsAndChannels(): Promise<TeamOrChannelRow[]> {
    return (await this.#readDomList('teamsAndChannels')) as TeamOrChannelRow[];
  }

  /**
   * Top-level posts from whichever channel is CURRENTLY DISPLAYED in the
   * signed-in tab (the Teams-and-Channels view must be open, not Chat).
   * There is no way to select a different channel from here — see the
   * module doc.
   */
  async getOpenChannelPosts(): Promise<ChannelPostRow[]> {
    return (await this.#readDomList('channelPosts')) as ChannelPostRow[];
  }

  /**
   * The sidebar row Teams marks as selected — i.e. which chat or channel is
   * currently open — or `null` when none can be identified.
   */
  async getOpenConversation(): Promise<OpenConversationRow | null> {
    const rows = (await this.#readDomList('openConversation')) as OpenConversationRow[];
    return rows[0] ?? null;
  }

  /** The Activity feed (the bell icon in the left nav), regardless of which item is selected. */
  async getActivity(): Promise<ActivityFeedRow[]> {
    return (await this.#readDomList('activityFeed')) as ActivityFeedRow[];
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
