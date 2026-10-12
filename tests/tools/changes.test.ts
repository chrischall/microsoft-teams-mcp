import { afterEach, beforeEach, describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { decodeChangesCursor, encodeChangesCursor, registerChangesTools } from '../../src/tools/changes.js';
import { UNTRUSTED_DESCRIPTION_SUFFIX } from '../../src/tools/untrusted.js';
import { lastChangeMs } from '../../src/model.js';
import type { ChatServiceClient, GetMessagesOptions, RawConversation, RawMessage } from '../../src/chatsvc.js';
import { DAY, fakeChatsvc, pinClock, rawConv, rawMsg } from './fake-chatsvc.js';

const NOW = 1_760_000_000_000;
const MIN = 60_000;
const HOUR = 60 * MIN;
/** The default cursor: 24 hours before now. */
const T = NOW - DAY;

beforeEach(() => pinClock(NOW));
afterEach(() => vi.useRealTimers());

const GROUP = '19:group0001@thread.v2';
const QUIET = '19:group0002@thread.v2';
const DM = '19:aaaa_bbbb@unq.gbl.spaces';
const MEETING = '19:meeting_abc0001@thread.v2';
const CHANNEL = '19:chan0001@thread.tacv2';
const HIDDEN = '19:chan0002@thread.tacv2';
const TEAM = '19:team0001@thread.tacv2';

/** A message as the service sends it, with `version` (its last change, epoch ms) set. */
function versioned(m: RawMessage, version: number): RawMessage {
  return { ...m, version: String(version) };
}

/**
 * A `getMessages` that pages the way the live service does: newest CHANGE
 * first (`version`, else the id), so an old message edited recently comes
 * ahead of newer messages. Stops before the first message `stopAt` accepts.
 */
function servedByVersion(byId: Record<string, RawMessage[]>) {
  const changed = (m: RawMessage) => Math.max(Number(m.version) || 0, Number(m.id));
  return vi.fn(async (id: string, opts: GetMessagesOptions) => {
    const all = [...(byId[id] ?? [])].sort((a, b) => changed(b) - changed(a));
    const out: RawMessage[] = [];
    for (const m of all) {
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

/** A channel listing entry whose lastMessage is an old root while a reply arrived at `updatedAt`. */
function channelConv(id: string, opts: { at: number; updatedAt?: number; interest?: string; topic?: string }): RawConversation {
  return {
    ...rawConv(id, { topic: opts.topic ?? 'Design', at: opts.at, spaceId: TEAM, interest: opts.interest }),
    ...(opts.updatedAt !== undefined
      ? { lastUpdatedMessageId: opts.updatedAt, lastUpdatedMessageVersion: opts.updatedAt }
      : {}),
  };
}

const OLD_ROOT = T - 3 * DAY;

const CONVS: RawConversation[] = [
  rawConv(GROUP, { topic: 'Project Falcon', at: NOW - 2 * HOUR }),
  rawConv(QUIET, { topic: 'Nothing new', at: T - HOUR }),
  rawConv(DM, { at: NOW - 3 * HOUR, sender: 'Example Friend' }),
  rawConv(MEETING, { topic: 'Weekly sync', at: NOW - 5 * HOUR }),
  // The channel's lastMessage is a root from days ago; only lastUpdatedMessageId shows the new replies.
  channelConv(CHANNEL, { at: OLD_ROOT, updatedAt: NOW - HOUR }),
  channelConv(HIDDEN, { at: NOW - 4 * HOUR, interest: 'NotInterested', topic: 'Hidden one' }),
  rawConv(TEAM, { topic: 'Platform Team', at: T - 10 * DAY }),
];

const MESSAGES: Record<string, RawMessage[]> = {
  [GROUP]: [
    rawMsg(T - HOUR, { text: 'before the cursor' }),
    // Edited after the cursor: served first, but not new, and must not end the read.
    versioned(rawMsg(T - 2 * HOUR, { text: 'old, edited' }), NOW - 30 * MIN),
    rawMsg(NOW - 4 * HOUR, { text: 'new 1' }),
    rawMsg(NOW - 3.5 * HOUR, { type: 'ThreadActivity/AddMember', text: '<addmember/>' }),
    rawMsg(NOW - 3 * HOUR, { text: 'mine', fromMe: true }),
    rawMsg(NOW - 2.5 * HOUR, { text: 'deleted', deleted: true }),
    rawMsg(NOW - 2 * HOUR, { text: 'new 2' }),
  ],
  [QUIET]: [rawMsg(T - HOUR)],
  [DM]: [rawMsg(T - 2 * HOUR, { conv: DM }), rawMsg(NOW - 3 * HOUR, { conv: DM, text: 'dm new' })],
  [MEETING]: [rawMsg(NOW - 5 * HOUR, { conv: MEETING, type: 'Event/Call', text: '<call/>' })],
  [CHANNEL]: [
    rawMsg(OLD_ROOT, { conv: CHANNEL, text: 'old root' }),
    rawMsg(OLD_ROOT + MIN, { conv: CHANNEL, rootId: OLD_ROOT, text: 'old reply' }),
    rawMsg(NOW - 6 * HOUR, { conv: CHANNEL, rootId: OLD_ROOT, text: 'new reply to old root' }),
    rawMsg(NOW - 5 * HOUR, { conv: CHANNEL, rootId: NOW - 5 * HOUR, text: 'new post' }),
    rawMsg(NOW - 2 * HOUR, { conv: CHANNEL, rootId: NOW - 5 * HOUR, text: 'reply to new post' }),
    rawMsg(NOW - HOUR, { conv: CHANNEL, rootId: OLD_ROOT, text: 'another reply to old root' }),
  ],
  [HIDDEN]: [rawMsg(NOW - 4 * HOUR, { conv: HIDDEN, rootId: NOW - 4 * HOUR })],
};

function chatsvcWith(overrides: Record<string, unknown> = {}, convs = CONVS, messages = MESSAGES): ChatServiceClient {
  return fakeChatsvc(convs, { getMessages: servedByVersion(messages), ...overrides });
}

async function call(chatsvc: ChatServiceClient, args: Record<string, unknown> = {}) {
  const harness = await createTestHarness((server) => registerChangesTools(server, { chatsvc }));
  try {
    const result = await harness.callTool('teams_get_changes', args);
    return { result, data: result.isError ? undefined : (parseToolResult(result) as Record<string, unknown>) };
  } finally {
    await harness.close();
  }
}

type Row = {
  conversation: { id: string; title: string; kind: string };
  new_message_count: number;
  new_from_others_count: number;
  newest_time: string;
  threads?: { root_id: string; new_root: boolean; new_reply_count: number }[];
  truncated?: true;
};

const rows = (data: Record<string, unknown> | undefined) => (data?.conversations ?? []) as Row[];
const ids = (data: Record<string, unknown> | undefined) => rows(data).map((r) => r.conversation.id);
const row = (data: Record<string, unknown> | undefined, id: string) => rows(data).find((r) => r.conversation.id === id);

function errorText(result: { content?: unknown[] }): string {
  return (result.content?.[0] as { text?: string })?.text ?? '';
}

describe('changes cursor', () => {
  it('round-trips an epoch-ms time through an opaque base64url token', () => {
    const token = encodeChangesCursor(1_760_000_123_456);
    expect(token).toMatch(/^[A-Za-z0-9_-]+$/);
    expect(JSON.parse(Buffer.from(token, 'base64url').toString('utf8'))).toEqual({ t: 1_760_000_123_456 });
    expect(decodeChangesCursor(token)).toBe(1_760_000_123_456);
  });

  it('also accepts a plain ISO 8601 time', () => {
    expect(decodeChangesCursor('2026-10-12T09:00:00Z')).toBe(Date.parse('2026-10-12T09:00:00Z'));
  });

  it('rejects anything else with an McpToolError-style error', () => {
    for (const bad of ['not a cursor', Buffer.from('{"x":1}').toString('base64url'), Buffer.from('{"t":-5}').toString('base64url')]) {
      expect(() => decodeChangesCursor(bad), bad).toThrow(/cursor/i);
    }
  });
});

describe('lastChangeMs', () => {
  it('is the latest of the last message, lastimreceivedtime and the last updated message id/version', () => {
    const base = rawConv(CHANNEL, { at: 1000 });
    expect(lastChangeMs(base)).toBe(1000);
    expect(lastChangeMs({ ...base, lastUpdatedMessageId: 5000 })).toBe(5000);
    expect(lastChangeMs({ ...base, lastUpdatedMessageVersion: '7000' })).toBe(7000);
    expect(lastChangeMs({ id: 'x' })).toBeUndefined();
  });
});

describe('teams_get_changes', () => {
  it('is a read-only open-world tool whose description carries the untrusted suffix', async () => {
    const harness = await createTestHarness((server) => registerChangesTools(server, { chatsvc: chatsvcWith() }));
    const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_get_changes');
    await harness.close();
    expect(tool?.annotations).toMatchObject({ readOnlyHint: true, openWorldHint: true });
    expect(tool?.description).toContain(UNTRUSTED_DESCRIPTION_SUFFIX);
  });

  it('returns its result inside the untrusted-data envelope', async () => {
    const { data } = await call(chatsvcWith());
    expect(data?.untrusted_content).toBe(true);
  });

  it('defaults to the last 24 hours and lists only conversations with a real new message, most recent first', async () => {
    const { data } = await call(chatsvcWith());
    expect(data?.since).toBe(new Date(T).toISOString());
    // CHANNEL's newest real message is an hour ago, GROUP's two, DM's three. MEETING's
    // only news is a call (a system event), QUIET had nothing, HIDDEN is hidden.
    expect(ids(data)).toEqual([CHANNEL, GROUP, DM]);
  });

  it('does not read a conversation whose listing shows nothing after the cursor', async () => {
    const getMessages = servedByVersion(MESSAGES);
    await call(chatsvcWith({ getMessages }));
    const read = getMessages.mock.calls.map(([id]) => id);
    expect(read).not.toContain(QUIET);
    expect(read).not.toContain(TEAM);
    expect(read).toContain(CHANNEL);
  });

  it('counts real new messages only — no system events, no deleted ones, none before the cursor', async () => {
    const { data } = await call(chatsvcWith());
    const group = row(data, GROUP);
    // new 1, mine, new 2 — not the add-member event, the deleted one, or the edited old one.
    expect(group?.new_message_count).toBe(3);
    expect(group?.new_from_others_count).toBe(2);
    expect(group?.newest_time).toBe(new Date(NOW - 2 * HOUR).toISOString());
    expect(group?.conversation).toMatchObject({ id: GROUP, title: 'Project Falcon', kind: 'group' });
  });

  it('keeps reading past an old message that was edited after the cursor', async () => {
    // The edited message is served first; a read that stopped at it would see nothing new.
    const { data } = await call(chatsvcWith(), { kinds: ['chat'] });
    expect(row(data, GROUP)?.new_message_count).toBe(3);
  });

  it('nominates a channel by its last updated message when lastMessage is an old root', async () => {
    const { data } = await call(chatsvcWith());
    const channel = row(data, CHANNEL);
    expect(channel?.new_message_count).toBe(4);
  });

  it("names a channel's thread roots that got new replies, including roots older than the cursor", async () => {
    const { data } = await call(chatsvcWith());
    expect(row(data, CHANNEL)?.threads).toEqual([
      { root_id: String(OLD_ROOT), new_root: false, new_reply_count: 2 },
      { root_id: String(NOW - 5 * HOUR), new_root: true, new_reply_count: 1 },
    ]);
    expect(row(data, GROUP)?.threads).toBeUndefined();
  });

  it('reads from a cursor: only what arrived after it', async () => {
    const cursor = encodeChangesCursor(NOW - 2.5 * HOUR);
    const { data } = await call(chatsvcWith(), { cursor });
    expect(ids(data)).toEqual([CHANNEL, GROUP]);
    expect(row(data, GROUP)?.new_message_count).toBe(1);
    expect(row(data, CHANNEL)?.threads).toEqual([
      { root_id: String(OLD_ROOT), new_root: false, new_reply_count: 1 },
      { root_id: String(NOW - 5 * HOUR), new_root: false, new_reply_count: 1 },
    ]);
  });

  it('accepts an ISO time as the cursor', async () => {
    const { data } = await call(chatsvcWith(), { cursor: new Date(NOW - 2.5 * HOUR).toISOString() });
    expect(ids(data)).toEqual([CHANNEL, GROUP]);
  });

  it('refuses an unreadable cursor', async () => {
    const { result } = await call(chatsvcWith(), { cursor: 'garbage' });
    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/cursor/i);
  });

  it('returns next_cursor as the newest arrival seen, which reads nothing new when passed back', async () => {
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc);
    expect(decodeChangesCursor(data?.next_cursor as string)).toBe(NOW - HOUR);
    expect(data?.next_cursor_time).toBe(new Date(NOW - HOUR).toISOString());
    const again = await call(chatsvc, { cursor: data?.next_cursor });
    expect(ids(again.data)).toEqual([]);
    expect(decodeChangesCursor(again.data?.next_cursor as string)).toBe(NOW - HOUR);
  });

  it('keeps the input time as next_cursor when nothing is new', async () => {
    const cursor = encodeChangesCursor(NOW - MIN);
    const { data } = await call(chatsvcWith(), { cursor });
    expect(ids(data)).toEqual([]);
    expect(data?.next_cursor).toBe(cursor);
    expect(data?.total_new_messages).toBe(0);
  });

  it('advances next_cursor past system events it saw, so they are not re-read', async () => {
    const { data } = await call(chatsvcWith(), { kinds: ['meeting'] });
    expect(ids(data)).toEqual([]);
    expect(decodeChangesCursor(data?.next_cursor as string)).toBe(NOW - 5 * HOUR);
  });

  it('never moves next_cursor past what the listing had seen', async () => {
    // A message that arrived after the listing was taken (here: newer than any listed activity).
    const late = { ...MESSAGES, [DM]: [...MESSAGES[DM], rawMsg(NOW + 5 * MIN, { conv: DM })] };
    const { data } = await call(chatsvcWith({}, CONVS, late));
    expect(row(data, DM)?.new_message_count).toBe(2);
    expect(decodeChangesCursor(data?.next_cursor as string)).toBe(NOW - HOUR);
  });

  it('filters by kinds', async () => {
    expect(ids((await call(chatsvcWith(), { kinds: ['channel'] })).data)).toEqual([CHANNEL]);
    expect(ids((await call(chatsvcWith(), { kinds: ['oneOnOne'] })).data)).toEqual([DM]);
    expect(ids((await call(chatsvcWith(), { kinds: ['chat'] })).data)).toEqual([GROUP, DM]);
  });

  it('includes channels hidden in Teams only with include_hidden', async () => {
    const { data } = await call(chatsvcWith(), { include_hidden: true });
    expect(ids(data)).toEqual([CHANNEL, GROUP, DM, HIDDEN]);
    expect(row(data, HIDDEN)?.conversation).toMatchObject({ hidden: true });
  });

  it("with include_hidden, reads every hidden channel: Teams stops moving a hidden channel's listing", async () => {
    // Live: a hidden channel's lastMessage / lastUpdatedMessageId stayed days old while new posts arrived.
    const stale = CONVS.map((c) => (c.id === HIDDEN ? channelConv(HIDDEN, { at: T - 2 * DAY, interest: 'NotInterested', topic: 'Hidden one' }) : c));
    const { data } = await call(chatsvcWith({}, stale), { include_hidden: true });
    expect(row(data, HIDDEN)?.new_message_count).toBe(1);
  });

  it('returns at most limit conversations and names the rest, counting all of them', async () => {
    const { data } = await call(chatsvcWith(), { limit: 2 });
    expect(ids(data)).toEqual([CHANNEL, GROUP]);
    expect(data?.more_conversations).toEqual([
      expect.objectContaining({ id: DM, new_message_count: 1 }),
    ]);
    expect(data?.total_new_messages).toBe(4 + 3 + 1);
    // Every changed conversation was read, so the cursor still covers them all.
    expect(decodeChangesCursor(data?.next_cursor as string)).toBe(NOW - HOUR);
  });

  it('skips and counts a conversation that is gone (404/403)', async () => {
    const base = servedByVersion(MESSAGES);
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => {
      if (id === DM) throw new McpToolError('Teams chat service returned HTTP 403.', { status: 403, kind: 'http' });
      return base(id, opts);
    });
    const { data } = await call(chatsvcWith({ getMessages }));
    expect(ids(data)).toEqual([CHANNEL, GROUP]);
    expect(data?.skipped_unreachable).toBe(1);
  });

  it('lists a failed read in errors and holds next_cursor at the input time so nothing is lost', async () => {
    const base = servedByVersion(MESSAGES);
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => {
      if (id === DM) throw new Error('boom');
      return base(id, opts);
    });
    const { data } = await call(chatsvcWith({ getMessages }));
    expect(ids(data)).toEqual([CHANNEL, GROUP]);
    expect(data?.errors).toEqual([expect.objectContaining({ conversation: expect.objectContaining({ id: DM }) })]);
    expect(decodeChangesCursor(data?.next_cursor as string)).toBe(T);
  });

  it('fails the call when every read failed', async () => {
    const getMessages = vi.fn().mockRejectedValue(new Error('boom'));
    const { result } = await call(chatsvcWith({ getMessages }));
    expect(result.isError).toBe(true);
  });

  it('marks a conversation truncated when its read hit the raw-message cap', async () => {
    const flood = Array.from({ length: 1100 }, (_, i) => rawMsg(NOW - 20 * HOUR + i * MIN));
    const convs = [rawConv(GROUP, { topic: 'Busy', at: NOW - 20 * HOUR + 1099 * MIN })];
    const { data } = await call(chatsvcWith({}, convs, { [GROUP]: flood }));
    expect(row(data, GROUP)?.truncated).toBe(true);
    expect(row(data, GROUP)?.new_message_count).toBe(1000);
  });
});
