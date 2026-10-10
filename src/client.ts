/**
 * DOM-reading Teams client — now used for the Activity feed ONLY.
 *
 * Chats and channels are read through the Teams chat-service REST API
 * (`chatsvc.ts`), authenticated by a token captured once from the signed-in
 * tab. The Activity feed has no chat-service equivalent we have verified, so
 * it is still read off the page: every call here goes through `read_dom_list`
 * against the user's live, signed-in `teams.cloud.microsoft` tab, and returns
 * only what that tab has rendered (the Activity view must be open).
 * `@fetchproxy/server` has no capability to navigate a tab — every declared
 * capability is a fetch or a read, never a page interaction.
 * Verified live 2026-09-21 against a real Teams tenant.
 */
import {
  createFetchproxyTransport,
  type FetchproxyTransport,
} from '@chrischall/mcp-utils/fetchproxy';
import { readPortEnv } from '@chrischall/mcp-utils';
import type { DomListSelectorDecl } from '@fetchproxy/protocol';
import { bridgeDeadlineMs, captureWindowMs } from './capture-window.js';
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
 * Top-level rows in the Activity feed (the bell icon in the left nav) —
 * mentions, replies, reactions, and the like across every chat and channel.
 * Always rendered once that nav section is open, regardless of which item
 * is selected (a list, not a detail pane). `location` is the team/channel or
 * chat the activity happened in, prose-joined by Teams itself (e.g. "Example
 * Team > Example Channel"). `time` is prose (e.g. "1:19 PM"), not ISO 8601 —
 * Teams renders it that way here.
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

/** Every selector the client declares to the bridge (the pair-popup scope). */
export const DOM_LIST_SELECTORS: readonly DomListSelectorDecl[] = [ACTIVITY_FEED_SELECTOR];

export interface ActivityFeedRow {
  title?: string;
  preview?: string;
  /** Prose, e.g. "1:19 PM" — NOT ISO 8601. */
  time?: string;
  /** The team/channel or chat the activity happened in, prose-joined by Teams. */
  location?: string;
}

export interface TeamsClientOptions {
  /** Injected for tests so no suite ever touches the real bridge. */
  transport?: FetchproxyTransport;
}

export class TeamsClient {
  readonly #transport: FetchproxyTransport;
  /**
   * The in-flight (or settled) start, memoized so parallel first calls share
   * ONE `transport.start()` — two concurrent starts would each create and
   * persist a fresh fetchproxy identity keypair, leaving the paired key and
   * the on-disk key out of step. Cleared on failure so the next call retries.
   */
  #starting: Promise<void> | undefined;

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
        // The same deadline the token capture runs under. This transport may
        // be the concentrator a capture relays through, and it is what
        // teams_healthcheck reports as the bridge's fetch timeout — so it
        // must not sit at the library's 30s while TEAMS_CAPTURE_TIMEOUT asks
        // for longer.
        fetchTimeoutMs: bridgeDeadlineMs(captureWindowMs()),
        // stderr only — stdout is the JSON-RPC channel. Without this the
        // first-ever pairing (or a scope widening) leaves a tool call
        // hanging with no way for the user to know a ContextMint Bridge approval
        // is what it's waiting on.
        onPairCode: (code) => {
          console.error(
            `[microsoft-teams-mcp] fetchproxy pair code: ${code} — approve in the ContextMint Bridge extension popup`,
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
  #ensureStarted(): Promise<void> {
    this.#starting ??= this.#transport.start().catch((e: unknown) => {
      this.#starting = undefined;
      throw e;
    });
    return this.#starting;
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

  /** The Activity feed (the bell icon in the left nav), regardless of which item is selected. */
  async getActivity(): Promise<ActivityFeedRow[]> {
    return (await this.#readDomList('activityFeed')) as ActivityFeedRow[];
  }

  async close(): Promise<void> {
    await this.#transport.close();
    this.#starting = undefined;
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
