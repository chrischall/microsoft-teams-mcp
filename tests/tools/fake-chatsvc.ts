// A stub `ChatServiceClient` for tool tests: canned conversations, a fixed
// signed-in user, and no token or network anywhere. Invented names only.
import { vi } from 'vitest';
import type { ChatServiceClient, GetMessagesOptions, RawConversation, RawMessage } from '../../src/chatsvc.js';

export const ME = '8:orgid:00000000-0000-0000-0000-0000000000a1';
export const OTHER = '8:orgid:00000000-0000-0000-0000-0000000000b2';
const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;

/** A conversation whose last message arrived at `at` (epoch ms), read up to `horizon`. */
export function rawConv(
  id: string,
  opts: { topic?: string; at?: number; horizon?: number; fromMe?: boolean; sender?: string; spaceId?: string; text?: string } = {},
): RawConversation {
  const at = opts.at ?? 1_760_000_000_000;
  return {
    id,
    properties: { consumptionhorizon: `${opts.horizon ?? at};${at};${opts.horizon ?? at}` },
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
      messagetype: 'Text',
    },
  };
}

/** One message arriving at `at` (epoch ms — which is also its id). */
export function rawMsg(
  at: number,
  opts: { text?: string; fromMe?: boolean; sender?: string; type?: string; rootId?: number; conv?: string; deleted?: boolean } = {},
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
  };
}

/**
 * A `getMessages` stub that pages the way the real client does over the given
 * messages (in any order — served newest first): it stops BEFORE the first
 * message `stopAt` accepts, or once `max` are collected.
 */
export function messagesFrom(messages: RawMessage[]) {
  const newestFirst = [...messages].sort((a, b) => Number(b.id) - Number(a.id));
  return vi.fn(async (_id: string, opts: GetMessagesOptions) => {
    const out: RawMessage[] = [];
    for (const m of newestFirst) {
      if (opts.stopAt?.(m)) return out;
      out.push(m);
      if (out.length >= opts.max) return out;
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
    ...overrides,
  } as unknown as ChatServiceClient;
}
