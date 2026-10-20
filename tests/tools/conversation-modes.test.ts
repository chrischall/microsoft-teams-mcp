// teams_list_messages's output-size modes (format) and multi-conversation reads.
import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerConversationTools } from '../../src/tools/conversation.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import { fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;

const GROUP = '19:group0001@thread.v2';
const OTHER_GROUP = '19:group0002@thread.v2';
const CHANNEL = '19:chan0001@thread.tacv2';
const TEAM = '19:team0001@thread.tacv2';

const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * MIN, horizon: T0 + 2 * MIN }),
  rawConv(OTHER_GROUP, { topic: 'Falcon launch party', at: T0 + 2 * MIN }),
  rawConv(CHANNEL, { topic: 'Design', at: T0 + 5 * MIN, horizon: T0 + 1 * MIN, spaceId: TEAM }),
  rawConv(TEAM, { topic: 'Platform Team', at: T0 }),
];

const MESSAGES: Record<string, RawMessage[]> = {
  [GROUP]: [
    rawMsg(T0 + 1 * MIN, { text: 'first' }),
    rawMsg(T0 + 2 * MIN, { text: 'second', fromMe: true }),
    rawMsg(T0 + 3 * MIN, { text: 'third' }),
    rawMsg(T0 + 4 * MIN, { text: 'fourth' }),
  ],
  [OTHER_GROUP]: [rawMsg(T0 + 2 * MIN, { conv: OTHER_GROUP, text: 'party time' })],
  [CHANNEL]: [
    rawMsg(T0 - 10 * MIN, { conv: CHANNEL, text: 'old root' }),
    rawMsg(T0 + 2 * MIN, { conv: CHANNEL, rootId: T0 - 10 * MIN, text: 'reply to old root' }),
    rawMsg(T0 + 3 * MIN, { conv: CHANNEL, rootId: T0 + 3 * MIN, text: 'new post' }),
    rawMsg(T0 + 5 * MIN, { conv: CHANNEL, rootId: T0 + 3 * MIN, text: 'reply to new post' }),
  ],
};

function perConversation(byId: Record<string, RawMessage[]> = MESSAGES) {
  const pagers = new Map(Object.entries(byId).map(([id, msgs]) => [id, messagesFrom(msgs)]));
  return vi.fn(async (id: string, opts: GetMessagesOptions) => (pagers.get(id) ?? messagesFrom([]))(id, opts));
}

function chatsvcWith(overrides: Record<string, unknown> = {}): ChatServiceClient {
  return fakeChatsvc(CONVS, { getMessages: perConversation(), ...overrides });
}

async function call(chatsvc: ChatServiceClient, args: Record<string, unknown>) {
  const harness = await createTestHarness((server) => registerConversationTools(server, { chatsvc }));
  try {
    const result = await harness.callTool('teams_list_messages', args);
    return { result, data: result.isError ? undefined : (parseToolResult(result) as Record<string, unknown>) };
  } finally {
    await harness.close();
  }
}

type Entry = Record<string, unknown> & { conversation: { id: string } };

describe('teams_list_messages format', () => {
  it('defaults to full: every message field and the whole conversation summary', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP });
    const full = await call(chatsvcWith(), { conversation: GROUP, format: 'full' });

    expect(data?.conversation).toHaveProperty('lastMessagePreview');
    expect((data?.messages as Record<string, unknown>[])[0]).toHaveProperty('attachments');
    expect((data?.messages as Record<string, unknown>[])[0]).toHaveProperty('edited');
    expect({ ...full.data, fetched_at: 0 }).toEqual({ ...data, fetched_at: 0 });
  });

  it('compact: messages carry only id, time, sender and text, and the conversation only id, kind and title', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, format: 'compact' });

    expect(data?.conversation).toEqual({ id: GROUP, kind: 'group', title: 'Project Falcon' });
    expect(data?.messages).toEqual([
      { id: String(T0 + 1 * MIN), time: new Date(T0 + 1 * MIN).toISOString(), sender: 'Example Person', text: 'first' },
      { id: String(T0 + 2 * MIN), time: new Date(T0 + 2 * MIN).toISOString(), sender: 'Me Myself', text: 'second' },
      { id: String(T0 + 3 * MIN), time: new Date(T0 + 3 * MIN).toISOString(), sender: 'Example Person', text: 'third' },
      { id: String(T0 + 4 * MIN), time: new Date(T0 + 4 * MIN).toISOString(), sender: 'Example Person', text: 'fourth' },
    ]);
    // Everything else a read reports is unchanged.
    expect(data?.unread_count).toBe(2);
    expect(data?.next_cursor).toBe(String(T0 + 4 * MIN));
    expect(data).toHaveProperty('fetched_at');
  });

  it('compact keeps time_local when tz is asked for', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, format: 'compact', tz: 'America/New_York' });

    expect((data?.messages as Record<string, unknown>[])[0]).toHaveProperty('time_local');
  });

  it('compact groups a channel into threads of compact posts', async () => {
    const { data } = await call(chatsvcWith(), { conversation: CHANNEL, format: 'compact' });

    expect(data?.conversation).toEqual({ id: CHANNEL, kind: 'channel', title: 'Design', team: 'Platform Team' });
    const threads = data?.threads as { root: Record<string, unknown>; replies: Record<string, unknown>[] }[];
    expect(threads.map((t) => t.root.text ?? 'missing')).toEqual(['old root', 'new post']);
    expect(Object.keys(threads[1].replies[0]).sort()).toEqual(['id', 'sender', 'text', 'time']);
  });

  it('summary: no messages, just counts and the newest time', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, format: 'summary' });

    expect(data).not.toHaveProperty('messages');
    expect(data).not.toHaveProperty('threads');
    expect(data?.conversation).toEqual({ id: GROUP, kind: 'group', title: 'Project Falcon' });
    expect(data).toMatchObject({
      unread_count: 2,
      message_count: 4,
      newest_time: new Date(T0 + 4 * MIN).toISOString(),
      next_cursor: String(T0 + 4 * MIN),
    });
  });

  it('summary of a channel lists each thread: root id, subject, reply count and newest time', async () => {
    const { data } = await call(chatsvcWith(), { conversation: CHANNEL, format: 'summary' });

    expect(data).toMatchObject({ message_count: 4, unread_count: 3 });
    expect(data?.threads).toEqual([
      { root_id: String(T0 - 10 * MIN), reply_count: 1, newest_time: new Date(T0 + 2 * MIN).toISOString() },
      { root_id: String(T0 + 3 * MIN), reply_count: 1, newest_time: new Date(T0 + 5 * MIN).toISOString() },
    ]);
  });

  it('rejects an unknown format', async () => {
    const { result } = await call(chatsvcWith(), { conversation: GROUP, format: 'tiny' });

    expect(result.isError).toBe(true);
  });

  it('documents format', async () => {
    const harness = await createTestHarness((server) => registerConversationTools(server, { chatsvc: chatsvcWith() }));
    try {
      const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_list_messages')!;
      expect(tool.description).toMatch(/format/);
      expect(tool.description).toMatch(/compact/);
      expect(tool.description).toMatch(/summary/);
    } finally {
      await harness.close();
    }
  });
});

describe('teams_list_messages with several conversations', () => {
  it('reads each one, in the order asked, as {conversations: [...]} of the single-read result', async () => {
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc, { conversation: [GROUP, 'design'] });

    const conversations = data?.conversations as Entry[];
    expect(conversations.map((c) => c.conversation.id)).toEqual([GROUP, CHANNEL]);
    expect((conversations[0].messages as { text: string }[]).map((m) => m.text)).toEqual([
      'first',
      'second',
      'third',
      'fourth',
    ]);
    expect(conversations[0]).toMatchObject({ unread_count: 2, next_cursor: String(T0 + 4 * MIN) });
    expect(conversations[1]).toHaveProperty('threads');
    expect(conversations[0]).not.toHaveProperty('fetched_at');
    expect(data).toHaveProperty('fetched_at');
    expect(data).not.toHaveProperty('errors');
    expect(chatsvc.listConversations).toHaveBeenCalledTimes(1);
  });

  it('answers a one-element array with conversations, not the single-read shape', async () => {
    const { data } = await call(chatsvcWith(), { conversation: [GROUP] });

    expect((data?.conversations as Entry[]).map((c) => c.conversation.id)).toEqual([GROUP]);
    expect(data).not.toHaveProperty('messages');
  });

  it('applies unread_only, limit and format to each conversation', async () => {
    const { data } = await call(chatsvcWith(), {
      conversation: [GROUP, CHANNEL],
      unread_only: true,
      format: 'summary',
    });

    const conversations = data?.conversations as Entry[];
    expect(conversations[0]).toMatchObject({ conversation: { id: GROUP }, message_count: 2, unread_count: 2 });
    expect(conversations[1]).toMatchObject({ conversation: { id: CHANNEL }, message_count: 3, unread_count: 3 });
    expect(conversations[1].conversation).not.toHaveProperty('lastMessagePreview');
  });

  it('puts a conversation that fails into errors and still returns the rest', async () => {
    const pager = perConversation();
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => {
      if (id === CHANNEL) throw new McpToolError('Teams chat service is rate-limiting requests (429).', { status: 429 });
      return pager(id, opts);
    });
    const { result, data } = await call(chatsvcWith({ getMessages }), {
      conversation: [GROUP, CHANNEL, 'no such chat'],
    });

    expect(result.isError).toBeFalsy();
    expect((data?.conversations as Entry[]).map((c) => c.conversation.id)).toEqual([GROUP]);
    const errors = data?.errors as { conversation: string; id?: string; error: string; hint?: string }[];
    expect(errors.map((e) => e.conversation)).toEqual([CHANNEL, 'no such chat']);
    expect(errors[0]).toMatchObject({ id: CHANNEL, title: 'Design', kind: 'channel', team: 'Platform Team' });
    expect(errors[1]).not.toHaveProperty('title');
    expect(errors[0].error).toMatch(/rate-limiting/);
    expect(errors[1].error).toMatch(/no chat or channel matches/i);
    expect(errors[1].hint).toMatch(/teams_list_conversations/);
  });

  it('returns an ambiguous name among the conversations, in the order asked, as the single-read candidates object', async () => {
    const chatsvc = chatsvcWith();
    const multi = await call(chatsvc, { conversation: [CHANNEL, 'falcon', GROUP] });
    const single = await call(chatsvc, { conversation: 'falcon' });

    const conversations = multi.data?.conversations as Record<string, unknown>[];
    expect(conversations).toHaveLength(3);
    expect((conversations[0] as Entry).conversation.id).toBe(CHANNEL);
    expect((conversations[2] as Entry).conversation.id).toBe(GROUP);
    const { untrusted_content: _u, note: _n, ...singleBody } = single.data!;
    expect(conversations[1]).toEqual(singleBody);
    expect(conversations[1]).toMatchObject({ ambiguous: true, query: 'falcon', hint: expect.any(String) });
    expect((conversations[1].candidates as { id: string }[]).map((c) => c.id).sort()).toEqual([GROUP, OTHER_GROUP].sort());
    expect(multi.data).not.toHaveProperty('errors');
  });

  it('returns errors only, not a tool error, when every conversation fails', async () => {
    const { result, data } = await call(chatsvcWith(), { conversation: ['nothing here', 'nor here'] });

    expect(result.isError).toBeFalsy();
    expect(data?.conversations).toEqual([]);
    expect(data?.errors).toHaveLength(2);
  });

  it('still fails the call when the listing itself cannot be read', async () => {
    const listConversations = vi.fn().mockRejectedValue(new Error('Could not capture a Teams token'));
    const { result } = await call(chatsvcWith({ listConversations }), { conversation: [GROUP, CHANNEL] });

    expect(result.isError).toBe(true);
  });

  it('reads a conversation named twice only once', async () => {
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc, { conversation: [GROUP, 'project falcon', GROUP] });

    expect((data?.conversations as Entry[]).map((c) => c.conversation.id)).toEqual([GROUP]);
    expect(chatsvc.getMessages).toHaveBeenCalledTimes(1);
  });

  it('reads at most 4 conversations at once', async () => {
    const many = Array.from({ length: 10 }, (_, i) => rawConv(`19:many${i}@thread.v2`, { topic: `Many ${i}` }));
    let inFlight = 0;
    let peak = 0;
    const getMessages = vi.fn(async () => {
      inFlight++;
      peak = Math.max(peak, inFlight);
      await new Promise((r) => setTimeout(r, 5));
      inFlight--;
      return [];
    });
    const { data } = await call(fakeChatsvc(many, { getMessages }), { conversation: many.map((c) => c.id) });

    expect(data?.conversations).toHaveLength(10);
    expect(peak).toBeLessThanOrEqual(4);
    expect(peak).toBeGreaterThan(1);
  });

  it('accepts up to 20 conversations and rejects more, or none', async () => {
    const ids = (n: number) => Array.from({ length: n }, (_, i) => `19:many${i}@thread.v2`);
    const twenty = Array.from({ length: 20 }, (_, i) => rawConv(`19:many${i}@thread.v2`, { topic: `Many ${i}` }));

    expect((await call(fakeChatsvc(twenty), { conversation: ids(20) })).result.isError).toBeFalsy();
    expect((await call(fakeChatsvc(twenty), { conversation: ids(21) })).result.isError).toBe(true);
    expect((await call(fakeChatsvc(twenty), { conversation: [] })).result.isError).toBe(true);
  });

  it('wraps the multi-read in the untrusted-data envelope', async () => {
    const { result } = await call(chatsvcWith(), { conversation: [GROUP, CHANNEL] });

    expect((result.content?.[0] as { text: string }).text).toMatch(/untrusted/i);
  });

  it('documents passing several conversations', async () => {
    const harness = await createTestHarness((server) => registerConversationTools(server, { chatsvc: chatsvcWith() }));
    try {
      const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_list_messages')!;
      expect(tool.description).toMatch(/array/);
      expect(tool.description).toMatch(/errors/);
    } finally {
      await harness.close();
    }
  });
});
