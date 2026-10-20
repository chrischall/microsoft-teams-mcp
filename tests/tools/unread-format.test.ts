// teams_list_unread_messages's output-size modes (format).
import { afterEach, beforeEach, describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { registerUnreadTools } from '../../src/tools/unread.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import { fakeChatsvc, messagesFrom, pinClock, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;

beforeEach(() => pinClock(T0 + 60 * MIN));
afterEach(() => vi.useRealTimers());

const GROUP = '19:group0001@thread.v2';
const CHANNEL = '19:chan0001@thread.tacv2';
const TEAM = '19:team0001@thread.tacv2';

const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * MIN, horizon: T0 + 2 * MIN }),
  rawConv(CHANNEL, { topic: 'Design', at: T0 + 5 * MIN, horizon: T0 + 1 * MIN, spaceId: TEAM }),
  rawConv(TEAM, { topic: 'Platform Team', at: T0 }),
];

const MESSAGES: Record<string, RawMessage[]> = {
  [GROUP]: [
    rawMsg(T0 + 2 * MIN, { text: 'read' }),
    rawMsg(T0 + 3 * MIN, { text: 'group unread 1' }),
    rawMsg(T0 + 4 * MIN, { text: 'group unread 2' }),
  ],
  [CHANNEL]: [
    rawMsg(T0 - 10 * MIN, { conv: CHANNEL, text: 'old root' }),
    rawMsg(T0 + 2 * MIN, { conv: CHANNEL, rootId: T0 - 10 * MIN, text: 'reply to old root' }),
    rawMsg(T0 + 3 * MIN, { conv: CHANNEL, rootId: T0 + 3 * MIN, text: 'new post' }),
    rawMsg(T0 + 5 * MIN, { conv: CHANNEL, rootId: T0 + 3 * MIN, text: 'reply to new post' }),
  ],
};

function chatsvcWith(): ChatServiceClient {
  const pagers = new Map(Object.entries(MESSAGES).map(([id, msgs]) => [id, messagesFrom(msgs)]));
  const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => (pagers.get(id) ?? messagesFrom([]))(id, opts));
  return fakeChatsvc(CONVS, { getMessages });
}

async function call(args: Record<string, unknown> = {}) {
  const harness = await createTestHarness((server) => registerUnreadTools(server, { chatsvc: chatsvcWith() }));
  try {
    const result = await harness.callTool('teams_list_unread_messages', args);
    return { result, data: result.isError ? undefined : (parseToolResult(result) as Record<string, unknown>) };
  } finally {
    await harness.close();
  }
}

type Entry = Record<string, unknown> & { conversation: Record<string, unknown> & { id: string } };
const entry = (data: Record<string, unknown> | undefined, id: string) =>
  (data?.conversations as Entry[]).find((e) => e.conversation.id === id)!;

describe('teams_list_unread_messages format', () => {
  it('defaults to full, the same as format "full"', async () => {
    const { data } = await call();
    const full = await call({ format: 'full' });

    expect(entry(data, GROUP).conversation).toHaveProperty('unread_confirmed', true);
    expect((entry(data, GROUP).messages as Record<string, unknown>[])[0]).toHaveProperty('attachments');
    expect({ ...full.data, fetched_at: 0 }).toEqual({ ...data, fetched_at: 0 });
  });

  it('compact: each conversation keeps its id, each message only id, time, sender and text', async () => {
    const { data } = await call({ format: 'compact' });

    const group = entry(data, GROUP);
    expect(group.conversation).toEqual({ id: GROUP, kind: 'group', title: 'Project Falcon' });
    expect(group.messages).toEqual([
      { id: String(T0 + 3 * MIN), time: new Date(T0 + 3 * MIN).toISOString(), sender: 'Example Person', text: 'group unread 1' },
      { id: String(T0 + 4 * MIN), time: new Date(T0 + 4 * MIN).toISOString(), sender: 'Example Person', text: 'group unread 2' },
    ]);
    expect(group).toMatchObject({ unread_count: 2, next_cursor: String(T0 + 4 * MIN) });
    const channel = entry(data, CHANNEL);
    expect(channel.conversation).toEqual({ id: CHANNEL, kind: 'channel', title: 'Design', team: 'Platform Team' });
    const threads = channel.threads as { root: Record<string, unknown>; replies: Record<string, unknown>[] }[];
    expect(threads[0].root).toEqual({ id: String(T0 - 10 * MIN), missing: true });
    expect(Object.keys(threads[1].replies[0]).sort()).toEqual(['id', 'sender', 'text', 'time']);
    expect(data?.next_cursors).toEqual({ [GROUP]: String(T0 + 4 * MIN), [CHANNEL]: String(T0 + 5 * MIN) });
  });

  it('summary: per conversation id, title, kind, counts, newest time and channel threads; no messages', async () => {
    const { data } = await call({ format: 'summary', tz: 'UTC' });

    const group = entry(data, GROUP);
    expect(group).not.toHaveProperty('messages');
    expect(group).toMatchObject({
      conversation: { id: GROUP, kind: 'group', title: 'Project Falcon' },
      unread_count: 2,
      message_count: 2,
      newest_time: new Date(T0 + 4 * MIN).toISOString(),
      next_cursor: String(T0 + 4 * MIN),
    });
    expect(group).toHaveProperty('newest_time_local');
    const channel = entry(data, CHANNEL);
    expect(channel).toMatchObject({ unread_count: 3, message_count: 3 });
    expect((channel.threads as Record<string, unknown>[]).map(({ newest_time_local: _l, ...t }) => t)).toEqual([
      { root_id: String(T0 - 10 * MIN), root_missing: true, reply_count: 1, newest_time: new Date(T0 + 2 * MIN).toISOString() },
      { root_id: String(T0 + 3 * MIN), reply_count: 1, newest_time: new Date(T0 + 5 * MIN).toISOString() },
    ]);
    expect(data?.total_unread).toBe(5);
  });

  it('rejects an unknown format', async () => {
    const { result } = await call({ format: 'tiny' });

    expect(result.isError).toBe(true);
  });

  it('documents format', async () => {
    const harness = await createTestHarness((server) => registerUnreadTools(server, { chatsvc: chatsvcWith() }));
    try {
      const tool = (await harness.client.listTools()).tools[0];
      expect(tool.description).toMatch(/compact/);
      expect(tool.description).toMatch(/summary/);
    } finally {
      await harness.close();
    }
  });
});
