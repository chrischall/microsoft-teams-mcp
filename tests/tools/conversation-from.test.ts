// teams_list_messages's `from` paging: oldest first from a UTC time, resumable with next_from.
import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { registerConversationTools } from '../../src/tools/conversation.js';
import { decodeFrom, encodeFrom, FROM_MAX_RAW } from '../../src/tools/from.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import { fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
const iso = (ms: number) => new Date(ms).toISOString();

const GROUP = '19:group0001@thread.v2';
const OTHER_GROUP = '19:group0002@thread.v2';
const CHANNEL = '19:chan0001@thread.tacv2';
const TEAM = '19:team0001@thread.tacv2';

const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 6 * MIN, horizon: T0 + 4 * MIN }),
  rawConv(OTHER_GROUP, { topic: 'Launch party', at: T0 + 2 * MIN }),
  rawConv(CHANNEL, { topic: 'Design', at: T0 + 5 * MIN, horizon: T0 + 1 * MIN, spaceId: TEAM }),
  rawConv(TEAM, { topic: 'Platform Team', at: T0 }),
];

const MESSAGES: Record<string, RawMessage[]> = {
  [GROUP]: [
    rawMsg(T0 + 1 * MIN, { text: 'one' }),
    rawMsg(T0 + 2 * MIN, { text: 'two', fromMe: true }),
    rawMsg(T0 + 2.5 * MIN, { type: 'ThreadActivity/AddMember', text: '<addmember/>' }),
    rawMsg(T0 + 3 * MIN, { text: 'three' }),
    rawMsg(T0 + 4 * MIN, { text: 'four' }),
    rawMsg(T0 + 5 * MIN, { text: 'five' }),
    rawMsg(T0 + 6 * MIN, { text: 'six' }),
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

const errorText = (result: { content?: unknown[] }) => (result.content?.[0] as { text?: string })?.text ?? '';
const texts = (data: Record<string, unknown> | undefined) => (data?.messages as { text: string }[]).map((m) => m.text);

describe('teams_list_messages from', () => {
  it('returns what arrived at or after from, oldest first, up to limit, with next_from and has_more', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, from: iso(T0 + 2 * MIN), limit: 2 });
    expect(texts(data)).toEqual(['two', 'three']);
    expect(data?.has_more).toBe(true);
    expect(data?.from_time).toBe(iso(T0 + 2 * MIN));
    expect(data?.next_from_time).toBe(iso(T0 + 3 * MIN));
    expect(decodeFrom(data?.next_from as string)).toEqual({ t: T0 + 3 * MIN, seen: [String(T0 + 3 * MIN)] });
  });

  it('resumes from next_from with no repeats and no gaps, then keeps returning a position when caught up', async () => {
    const chatsvc = chatsvcWith();
    const seen: string[] = [];
    let from = iso(T0);
    for (let i = 0; i < 10; i++) {
      const { data } = await call(chatsvc, { conversation: GROUP, from, limit: 2 });
      seen.push(...texts(data));
      from = data?.next_from as string;
      if (data?.has_more === false) break;
    }
    expect(seen).toEqual(['one', 'two', 'three', 'four', 'five', 'six']);
    const { data: idle } = await call(chatsvc, { conversation: GROUP, from });
    expect(texts(idle)).toEqual([]);
    expect(idle?.has_more).toBe(false);
    expect(idle?.next_from).toBe(from);
  });

  it('defaults to 50 and stops reading at from, not the start of the history', async () => {
    const getMessages = perConversation();
    const { data } = await call(chatsvcWith({ getMessages }), { conversation: GROUP, from: iso(T0 + 5 * MIN) });
    expect(texts(data)).toEqual(['five', 'six']);
    expect(data?.has_more).toBe(false);
    const opts = getMessages.mock.calls[0][1] as GetMessagesOptions;
    expect(opts.max).toBe(FROM_MAX_RAW);
    expect(opts.maxPages).toBeGreaterThan(20);
  });

  it('finds a message that arrived after from even when an older edit is paged ahead of it', async () => {
    // Pages come newest CHANGE first: "one" edited at +10 is served before "six".
    const edited = { [GROUP]: [rawMsg(T0 + 1 * MIN, { text: 'one', version: T0 + 10 * MIN }), ...MESSAGES[GROUP].slice(1)] };
    const { data } = await call(chatsvcWith({ getMessages: perConversation(edited) }), {
      conversation: GROUP,
      from: iso(T0 + 3 * MIN),
    });
    expect(texts(data)).toEqual(['three', 'four', 'five', 'six']);
  });

  it('with unread_only, pages only the unread messages past from', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, from: iso(T0), unread_only: true });
    expect(texts(data)).toEqual(['five', 'six']);
  });

  it('groups a channel page under its roots, oldest first', async () => {
    const { data } = await call(chatsvcWith(), { conversation: CHANNEL, from: iso(T0 + 3 * MIN), format: 'compact' });
    const threads = data?.threads as { root: { text?: string }; replies: { text: string }[] }[];
    expect(threads.map((t) => [t.root.text, ...t.replies.map((r) => r.text)])).toEqual([['new post', 'reply to new post']]);
    expect(data?.has_more).toBe(false);
  });

  it('pages each conversation of an array on its own', async () => {
    const { data } = await call(chatsvcWith(), { conversation: [GROUP, OTHER_GROUP], from: iso(T0 + 2 * MIN), limit: 1 });
    const reads = data?.conversations as Record<string, unknown>[];
    expect(reads.map((r) => (r.messages as { text: string }[]).map((m) => m.text))).toEqual([['two'], ['party time']]);
    expect(reads.map((r) => r.has_more)).toEqual([true, false]);
  });

  it('refuses a conversation busier since from than a read can reach, rather than skip its oldest', async () => {
    const getMessages = vi.fn(async (_id: string, opts: GetMessagesOptions) => {
      opts.onCapped?.();
      return [rawMsg(T0 + 6 * MIN, { text: 'six' })];
    });
    const { result } = await call(chatsvcWith({ getMessages }), { conversation: GROUP, from: iso(T0) });
    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/later from/);
  });

  it('refuses from beside since, after_message_id or mark_read', async () => {
    for (const extra of [{ since: iso(T0) }, { after_message_id: String(T0) }, { mark_read: true }]) {
      const { result } = await call(chatsvcWith(), { conversation: GROUP, from: iso(T0), ...extra });
      expect(result.isError, JSON.stringify(extra)).toBe(true);
      expect(errorText(result)).toMatch(/from/);
    }
  });

  it('refuses a from that is not a time or a token', async () => {
    const { result } = await call(chatsvcWith(), { conversation: GROUP, from: 'last tuesday' });
    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/next_from/);
  });

  it('accepts its own token unchanged', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, from: encodeFrom({ t: T0 + 6 * MIN, seen: [] }) });
    expect(texts(data)).toEqual(['six']);
  });
});
