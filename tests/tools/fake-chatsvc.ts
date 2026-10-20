// A stub `ChatServiceClient` for tool tests: canned conversations, a fixed
// signed-in user, and no token or network anywhere. Invented names only.
import { vi } from 'vitest';
import type { ChatServiceClient, GetMessagesOptions, RawConversation, RawMessage } from '../../src/chatsvc.js';

export const ME = '8:orgid:00000000-0000-0000-0000-0000000000a1';
/** One day in ms. */
export const DAY = 86_400_000;

/**
 * Pin the wall clock (only `Date`; timers stay real) so the unread tools'
 * since_days window is measured from a fixed "now". Undo with `vi.useRealTimers()`.
 */
export function pinClock(nowMs: number): void {
  vi.useFakeTimers({ toFake: ['Date'] });
  vi.setSystemTime(nowMs);
}

export const OTHER = '8:orgid:00000000-0000-0000-0000-0000000000b2';
const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;

/** A conversation whose last message arrived at `at` (epoch ms), read up to `horizon`. */
export function rawConv(
  id: string,
  opts: {
    topic?: string;
    at?: number;
    horizon?: number;
    fromMe?: boolean;
    sender?: string;
    spaceId?: string;
    text?: string;
    /** The last message's type (default `Text`). */
    type?: string;
    /** `properties.lastimreceivedtime`, as epoch ms (written as ISO, as Teams sends it). */
    lastIm?: number;
    /** `memberProperties.interest` — `NotInterested` is a channel hidden in Teams. */
    interest?: string;
    /** `properties.favorite` as Teams sends it (`"true"`, `"True"`, `"false"`). */
    favorite?: string;
    /** `properties.ispinned`, likewise. */
    pinned?: string;
    /** `properties.alerts` — `"false"` is a conversation muted in Teams. */
    alerts?: string;
  } = {},
): RawConversation {
  const at = opts.at ?? 1_760_000_000_000;
  return {
    id,
    properties: {
      consumptionhorizon: `${opts.horizon ?? at};${at};${opts.horizon ?? at}`,
      ...(opts.lastIm !== undefined ? { lastimreceivedtime: new Date(opts.lastIm).toISOString() } : {}),
      ...(opts.favorite !== undefined ? { favorite: opts.favorite } : {}),
      ...(opts.pinned !== undefined ? { ispinned: opts.pinned } : {}),
      ...(opts.alerts !== undefined ? { alerts: opts.alerts } : {}),
    },
    ...(opts.interest !== undefined ? { memberProperties: { interest: opts.interest } } : {}),
    threadProperties: {
      ...(opts.topic !== undefined ? { topic: opts.topic } : {}),
      ...(opts.spaceId !== undefined ? { spaceId: opts.spaceId } : {}),
    },
    lastMessage: {
      id: String(at),
      originalarrivaltime: new Date(at).toISOString(),
      content: opts.text ?? 'Example text',
      from: FROM(opts.fromMe ? ME : OTHER),
      imdisplayname: opts.fromMe ? 'Me Myself' : (opts.sender ?? 'Example Person'),
      messagetype: opts.type ?? 'Text',
    },
  };
}

/** One message arriving at `at` (epoch ms — which is also its id). */
export function rawMsg(
  at: number,
  opts: {
    text?: string;
    fromMe?: boolean;
    sender?: string;
    type?: string;
    rootId?: number;
    conv?: string;
    deleted?: boolean;
    /** Its last change (epoch ms) — an edit or reaction after it arrived. */
    version?: number;
  } = {},
): RawMessage {
  const conv = opts.conv ?? '19:group0001@thread.v2';
  return {
    id: String(at),
    originalarrivaltime: new Date(at).toISOString(),
    content: opts.text ?? `message ${at}`,
    messagetype: opts.type ?? 'Text',
    from: FROM(opts.fromMe ? ME : OTHER),
    imdisplayname: opts.fromMe ? 'Me Myself' : (opts.sender ?? 'Example Person'),
    conversationLink:
      `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/conversations/${conv}` +
      (opts.rootId !== undefined ? `;messageid=${opts.rootId}` : ''),
    ...(opts.deleted ? { properties: { deletetime: String(at + 1) } } : {}),
    ...(opts.version !== undefined ? { version: String(opts.version) } : {}),
  };
}

/**
 * A `getMessages` stub that pages the way the real client does over the given
 * messages (in any order — served newest CHANGE first, as live: a message's
 * `version` when it has one, else its arrival): it stops BEFORE the first
 * message `stopAt` accepts, or once `max` are collected (reporting `onCapped`).
 */
export function messagesFrom(messages: RawMessage[], pageSize = Infinity) {
  const changed = (m: RawMessage) => Math.max(Number(m.version ?? 0) || 0, Number(m.id));
  const newestFirst = [...messages].sort((a, b) => changed(b) - changed(a));
  return vi.fn(async (_id: string, opts: GetMessagesOptions) => {
    const out: RawMessage[] = [];
    for (const [i, m] of newestFirst.entries()) {
      // Like the real client: onPage runs between pages, before the next one is requested.
      if (i > 0 && i % pageSize === 0) await opts.onPage?.(i / pageSize, out.length);
      if (opts.stopAt?.(m)) return out;
      out.push(m);
      if (out.length >= opts.max) {
        opts.onCapped?.();
        return out;
      }
    }
    return out;
  });
}

export function fakeChatsvc(
  conversations: RawConversation[] = [],
  overrides: Partial<Record<keyof ChatServiceClient, unknown>> = {},
): ChatServiceClient {
  return {
    listConversations: vi.fn().mockResolvedValue(conversations),
    meMri: vi.fn().mockResolvedValue(ME),
    getMessages: vi.fn().mockResolvedValue([]),
    // No message by id unless a test says so: a channel root outside the window stays missing.
    getMessage: vi.fn().mockResolvedValue(undefined),
    ...overrides,
  } as unknown as ChatServiceClient;
}
