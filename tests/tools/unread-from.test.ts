// teams_list_unread_messages's `from` paging: every unread message since a UTC time, across conversations, oldest first.
import { afterEach, beforeEach, describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerUnreadTools } from '../../src/tools/unread.js';
import { decodeFrom } from '../../src/tools/from.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import { fakeChatsvc, messagesFrom, pinClock, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
const NOW = T0 + 60 * MIN;
const iso = (ms: number) => new Date(ms).toISOString();

beforeEach(() => pinClock(NOW));
afterEach(() => vi.useRealTimers());

const GROUP = '19:group0001@thread.v2';
const DM = '19:aaaa_bbbb@unq.gbl.spaces';
const CHANNEL = '19:chan0001@thread.tacv2';
const TEAM = '19:team0001@thread.tacv2';
const MUTED = '19:group0009@thread.v2';
const GONE = '19:group0404@thread.v2';

const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 6 * MIN, horizon: T0 + 1 * MIN }),
  rawConv(DM, { at: T0 + 5 * MIN, horizon: T0, sender: 'Example Friend' }),
  rawConv(CHANNEL, { topic: 'Design', at: T0 + 4 * MIN, horizon: T0, spaceId: TEAM }),
  rawConv(MUTED, { topic: 'Noisy', at: T0 + 3 * MIN, horizon: T0, alerts: 'false' }),
  rawConv(TEAM, { topic: 'Platform Team', at: T0 }),
];

const MESSAGES: Record<string, RawMessage[]> = {
  [GROUP]: [
    rawMsg(T0 + 1 * MIN, { text: 'group read' }),
    rawMsg(T0 + 2 * MIN, { text: 'group a' }),
    rawMsg(T0 + 3 * MIN, { text: 'mine', fromMe: true }),
    rawMsg(T0 + 3.5 * MIN, { type: 'Event/Call', text: '<call/>' }),
    rawMsg(T0 + 6 * MIN, { text: 'group b' }),
  ],
  [DM]: [rawMsg(T0 + 2 * MIN, { conv: DM, text: 'dm a' }), rawMsg(T0 + 5 * MIN, { conv: DM, text: 'dm b' })],
  [CHANNEL]: [
    rawMsg(T0 + 3 * MIN, { conv: CHANNEL, rootId: T0 + 3 * MIN, text: 'post' }),
    rawMsg(T0 + 4 * MIN, { conv: CHANNEL, rootId: T0 + 3 * MIN, text: 'reply' }),
  ],
  [MUTED]: [rawMsg(T0 + 3 * MIN, { conv: MUTED, text: 'muted' })],
};

function perConversation(byId: Record<string, RawMessage[]> = MESSAGES) {
  const pagers = new Map(Object.entries(byId).map(([id, msgs]) => [id, messagesFrom(msgs)]));
  return vi.fn(async (id: string, opts: GetMessagesOptions) => (pagers.get(id) ?? messagesFrom([]))(id, opts));
}

function chatsvcWith(overrides: Record<string, unknown> = {}, convs = CONVS): ChatServiceClient {
  return fakeChatsvc(convs, { getMessages: perConversation(), ...overrides });
}

async function call(chatsvc: ChatServiceClient, args: Record<string, unknown> = {}) {
  const harness = await createTestHarness((server) => registerUnreadTools(server, { chatsvc }));
  try {
    const result = await harness.callTool('teams_list_unread_messages', args);
    return { result, data: result.isError ? undefined : (parseToolResult(result) as Record<string, unknown>) };
  } finally {
    await harness.close();
  }
}

const errorText = (result: { content?: unknown[] }) => (result.content?.[0] as { text?: string })?.text ?? '';
type Flat = { conversation_id: string; text: string; replyTo?: string };
const texts = (data: Record<string, unknown> | undefined) => (data?.messages as Flat[]).map((m) => m.text);

describe('teams_list_unread_messages from', () => {
  it('returns every unread message since from across conversations, oldest first, as one list', async () => {
    const { data } = await call(chatsvcWith(), { from: iso(T0) });
    // group a and dm a share a millisecond: ties go by conversation id.
    expect(texts(data)).toEqual(['dm a', 'group a', 'post', 'reply', 'dm b', 'group b']);
    expect((data?.messages as Flat[]).map((m) => m.conversation_id)).toEqual([DM, GROUP, CHANNEL, CHANNEL, DM, GROUP]);
    expect((data?.messages as Flat[])[3].replyTo).toBe(String(T0 + 3 * MIN));
    expect((data?.conversations as { id: string; title: string }[]).map((c) => c.id).sort()).toEqual([CHANNEL, DM, GROUP].sort());
    expect(data?.total_unread).toBe(6);
    expect(data?.has_more).toBe(false);
    expect(data?.muted_skipped).toBe(1);
  });

  it('leaves out what arrived before from', async () => {
    const { data } = await call(chatsvcWith(), { from: iso(T0 + 4 * MIN) });
    expect(texts(data)).toEqual(['reply', 'dm b', 'group b']);
  });

  it('pages by limit across conversations and resumes from next_from with no gaps or repeats', async () => {
    const chatsvc = chatsvcWith();
    const first = await call(chatsvc, { from: iso(T0), limit: 4 });
    expect(texts(first.data)).toEqual(['dm a', 'group a', 'post', 'reply']);
    expect(first.data?.has_more).toBe(true);
    expect(first.data?.total_unread).toBe(6);
    expect(first.data?.next_from_time).toBe(iso(T0 + 4 * MIN));
    const second = await call(chatsvc, { from: first.data?.next_from, limit: 4 });
    expect(texts(second.data)).toEqual(['dm b', 'group b']);
    expect(second.data?.has_more).toBe(false);
    const idle = await call(chatsvc, { from: second.data?.next_from });
    expect(texts(idle.data)).toEqual([]);
    expect(idle.data?.next_from).toBe(second.data?.next_from);
  });

  it('splits messages sharing a millisecond across pages without losing either', async () => {
    const tie = {
      ...MESSAGES,
      [GROUP]: [rawMsg(T0 + 5 * MIN, { text: 'group tie' })],
      [DM]: [rawMsg(T0 + 5 * MIN, { conv: DM, text: 'dm tie' })],
      [CHANNEL]: [],
    };
    const chatsvc = chatsvcWith({ getMessages: perConversation(tie) });
    const first = await call(chatsvc, { from: iso(T0), limit: 1 });
    const second = await call(chatsvc, { from: first.data?.next_from, limit: 1 });
    const third = await call(chatsvc, { from: second.data?.next_from, limit: 1 });
    expect([...texts(first.data), ...texts(second.data), ...texts(third.data)].sort()).toEqual(['dm tie', 'group tie']);
    expect(second.data?.has_more).toBe(false);
  });

  it('never moves next_from past the newest change the listing knew of', async () => {
    // A message that landed after the listing was read: returned now, and again next time.
    const late = { ...MESSAGES, [DM]: [...MESSAGES[DM], rawMsg(T0 + 9 * MIN, { conv: DM, text: 'late' })] };
    const { data } = await call(chatsvcWith({ getMessages: perConversation(late) }), { from: iso(T0) });
    expect(texts(data).at(-1)).toBe('late');
    expect(decodeFrom(data?.next_from as string).t).toBe(T0 + 6 * MIN);
  });

  it('does not read a conversation with no activity since from', async () => {
    const getMessages = perConversation();
    await call(chatsvcWith({ getMessages }), { from: iso(T0 + 5.5 * MIN) });
    expect(getMessages.mock.calls.map((c) => c[0])).toEqual([GROUP]);
  });

  it('includes muted chats with include_muted, and reports a read that failed beside the rest', async () => {
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => {
      if (id === GROUP) throw new Error('boom');
      return perConversation()(id, opts);
    });
    const { data } = await call(chatsvcWith({ getMessages }), { from: iso(T0), include_muted: true });
    expect(texts(data)).toContain('muted');
    expect(texts(data)).not.toContain('group a');
    expect((data?.errors as { conversation: { id: string } }[]).map((e) => e.conversation.id)).toEqual([GROUP]);
    // A failed read may hold older unread messages, so the position does not move past them.
    expect(decodeFrom(data?.next_from as string).t).toBe(T0);
  });

  it('returns what it reached of a conversation too busy to read back to from, names the gap, and keeps moving', async () => {
    // GROUP's read runs out of depth having reached back only to +5: everything after +5 is known.
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => {
      if (id !== GROUP) return perConversation()(id, opts);
      opts.onCapped?.();
      return [rawMsg(T0 + 6 * MIN, { text: 'group b' }), rawMsg(T0 + 5 * MIN, { text: 'group reached' })];
    });
    const { data } = await call(chatsvcWith({ getMessages }), { from: iso(T0) });
    expect(texts(data)).toEqual(['dm a', 'post', 'reply', 'dm b', 'group b']);
    expect(data).not.toHaveProperty('errors');
    expect(data?.incomplete).toEqual([
      expect.objectContaining({
        conversation: expect.objectContaining({ id: GROUP }),
        // Its read marker is at +1: nothing before it can be unread, so the gap starts just past it.
        unread_not_read_from: iso(T0 + 1 * MIN + 1),
        unread_not_read_until: iso(T0 + 5 * MIN),
        hint: expect.stringMatching(/later from/),
      }),
    ]);
    // Not pinned like a failed read: the stream moves on.
    expect(decodeFrom(data?.next_from as string).t).toBe(T0 + 6 * MIN);
  });

  it('skips a conversation that is gone (404) and counts it', async () => {
    const convs = [...CONVS, rawConv(GONE, { topic: 'Gone', at: T0 + 2 * MIN, horizon: T0 })];
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => {
      if (id === GONE) throw new McpToolError('Teams chat service returned HTTP 404.', { status: 404, kind: 'http' });
      return perConversation()(id, opts);
    });
    const { data } = await call(chatsvcWith({ getMessages }, convs), { from: iso(T0) });
    expect(data?.skipped_unreachable).toBe(1);
    expect(texts(data)).toHaveLength(6);
  });

  it('trims each message in format compact, keeping its conversation and thread', async () => {
    const { data } = await call(chatsvcWith(), { from: iso(T0), format: 'compact' });
    const reply = (data?.messages as Record<string, unknown>[]).find((m) => m.text === 'reply')!;
    expect(reply).toMatchObject({ conversation_id: CHANNEL, replyTo: String(T0 + 3 * MIN) });
    expect(reply).not.toHaveProperty('attachments');
  });

  it('refuses from beside the per-conversation arguments, and limit without from', async () => {
    for (const extra of [
      { since_days: 3 },
      { cursors: { [GROUP]: String(T0) } },
      { per_conversation_limit: 5 },
      { max_conversations: 5 },
      { mark_read: true },
      { older_unread: true },
      { format: 'summary' },
    ]) {
      const { result } = await call(chatsvcWith(), { from: iso(T0), ...extra });
      expect(result.isError, JSON.stringify(extra)).toBe(true);
      expect(errorText(result)).toMatch(/from/);
    }
    const { result } = await call(chatsvcWith(), { limit: 5 });
    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/limit/);
  });
});
