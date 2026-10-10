// A stub `ChatServiceClient` for tool tests: canned conversations, a fixed
// signed-in user, and no token or network anywhere. Invented names only.
import { vi } from 'vitest';
import type { ChatServiceClient, RawConversation } from '../../src/chatsvc.js';

export const ME = '8:orgid:00000000-0000-0000-0000-0000000000a1';
const OTHER = '8:orgid:00000000-0000-0000-0000-0000000000b2';
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
