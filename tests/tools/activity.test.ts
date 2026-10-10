// teams_get_activity over the chat service's `48:notifications` conversation.
// Every fixture is invented: names, ids and text are placeholders.
import { afterEach, describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { activityKind, ACTIVITY_FEED_ID, registerActivityTools } from '../../src/tools/activity.js';
import type { RawMessage } from '../../src/chatsvc.js';
import { DAY, fakeChatsvc, messagesFrom, pinClock } from './fake-chatsvc.js';

const NOW = 1_760_000_000_000;

interface ActivityOpts {
  type: string;
  subtype?: string;
  read?: boolean;
  /** Arrival offset before NOW, in ms (default one hour). */
  ago?: number;
  from?: string;
  thread?: string;
  topic?: string;
  threadType?: string;
  messageId?: string;
  replyChainId?: string;
  preview?: string;
  /** Send `properties.activity` as an object rather than the usual JSON string. */
  asObject?: boolean;
  deleted?: boolean;
  /** Leave `properties.activity` out entirely. */
  noActivity?: boolean;
}

/** One entry of the `48:notifications` conversation, in the shape seen live. */
function entry(opts: ActivityOpts): RawMessage {
  const at = NOW - (opts.ago ?? 3_600_000);
  const activity = {
    activityType: opts.type,
    activitySubtype: opts.subtype ?? '',
    activityReason: 'example',
    activityTimestamp: new Date(at).toISOString(),
    activityId: 0,
    sourceThreadId: opts.thread ?? '19:group0001@thread.v2',
    sourceMessageId: opts.messageId ?? String(at - 1000),
    ...(opts.replyChainId !== undefined ? { sourceReplyChainId: opts.replyChainId } : {}),
    sourceUserId: '8:orgid:00000000-0000-0000-0000-0000000000b2',
    sourceUserImDisplayName: opts.from ?? 'Example Person',
    targetUserId: '8:orgid:00000000-0000-0000-0000-0000000000a1',
    messagePreview: opts.preview ?? 'Example preview',
    activityContext: {},
    sourceThreadTopic: opts.topic ?? 'Example Chat',
    sourceThreadType: opts.threadType ?? 'chat',
    sourceThreadIsPrivateChannel: false,
    threadParticipants: [],
    sourceThreadRosterNonBotMemberCount: 3,
  };
  return {
    id: String(at),
    originalarrivaltime: new Date(at).toISOString(),
    messagetype: 'Text',
    content: '',
    conversationid: ACTIVITY_FEED_ID,
    properties: {
      isread: opts.read ? 'true' : 'false',
      ...(opts.noActivity ? {} : { activity: opts.asObject ? activity : JSON.stringify(activity) }),
      ...(opts.deleted ? { deletetime: String(at + 1) } : {}),
    },
  };
}

async function call(entries: RawMessage[], args: Record<string, unknown> = {}) {
  const getMessages = messagesFrom(entries);
  const chatsvc = fakeChatsvc([], { getMessages });
  const harness = await createTestHarness((server) => registerActivityTools(server, { chatsvc }));
  try {
    const result = await harness.callTool('teams_get_activity', args);
    return { result, data: parseToolResult(result) as Record<string, unknown>, getMessages, chatsvc };
  } finally {
    await harness.close();
  }
}

type Row = Record<string, unknown> & { where?: Record<string, unknown> };
const rowsOf = (data: Record<string, unknown>) => data.rows as Row[];

describe('activityKind', () => {
  it.each([
    ['mention', 'mention'],
    ['mentionInChat', 'mention'],
    ['reaction', 'reaction'],
    ['reactionInChat', 'reaction'],
    ['reply', 'reply'],
    ['replyToReply', 'reply'],
    ['msGraph', 'meeting'],
    ['call', 'call'],
    ['somethingNew', 'other'],
    ['', 'other'],
  ])('%s is bucketed as %s', (type, kind) => {
    expect(activityKind(type)).toBe(kind);
  });
});

describe('teams_get_activity', () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it('reads the 48:notifications conversation directly, never the conversation listing', async () => {
    pinClock(NOW);
    const { getMessages, chatsvc } = await call([entry({ type: 'mention', subtype: 'person' })]);

    expect(getMessages).toHaveBeenCalledWith(ACTIVITY_FEED_ID, expect.anything());
    expect(chatsvc.listConversations).not.toHaveBeenCalled();
    expect(ACTIVITY_FEED_ID).toBe('48:notifications');
  });

  it('normalises an entry into a row whose ids feed teams_read_conversation', async () => {
    pinClock(NOW);
    const { data } = await call([
      entry({
        type: 'mention',
        subtype: 'channel',
        ago: 60_000,
        from: 'Example Person',
        thread: '19:channel0001@thread.tacv2',
        topic: 'Release Planning',
        threadType: 'topic',
        messageId: '1759999990000',
        replyChainId: '1759999900000',
        preview: '<p>Can <b>you</b> look at this?</p>',
      }),
    ]);

    expect(rowsOf(data)).toEqual([
      {
        id: String(NOW - 60_000),
        time: new Date(NOW - 60_000).toISOString(),
        kind: 'mention',
        type: 'mention',
        subtype: 'channel',
        read: false,
        from: 'Example Person',
        where: { conversation_id: '19:channel0001@thread.tacv2', title: 'Release Planning', type: 'topic' },
        message_id: '1759999990000',
        reply_chain_id: '1759999900000',
        preview: 'Can you look at this?',
      },
    ]);
  });

  it('accepts properties.activity as an object as well as a JSON string', async () => {
    pinClock(NOW);
    const { data } = await call([entry({ type: 'reply', asObject: true, preview: 'object form' })]);

    expect(rowsOf(data)).toHaveLength(1);
    expect(rowsOf(data)[0]).toMatchObject({ kind: 'reply', preview: 'object form' });
  });

  it('falls back to originalarrivaltime when activityTimestamp is missing', async () => {
    pinClock(NOW);
    const e = entry({ type: 'call', subtype: 'missedMeetingNudge', ago: 120_000 });
    const act = JSON.parse(e.properties!.activity as string) as Record<string, unknown>;
    delete act.activityTimestamp;
    e.properties!.activity = JSON.stringify(act);

    const { data } = await call([e]);

    expect(rowsOf(data)[0].time).toBe(new Date(NOW - 120_000).toISOString());
  });

  it('buckets every kind seen live', async () => {
    pinClock(NOW);
    const kinds: [string, string, string][] = [
      ['mention', 'person', 'mention'],
      ['mention', 'team', 'mention'],
      ['mentionInChat', 'everyone', 'mention'],
      ['reactionInChat', '2705_whiteheavycheckmark', 'reaction'],
      ['reaction', 'like', 'reaction'],
      ['reply', '', 'reply'],
      ['replyToReply', '', 'reply'],
      ['msGraph', 'privateMeetingCreated', 'meeting'],
      ['msGraph', 'privateMeetingForwardedNotificationWithAttendee', 'meeting'],
      ['call', 'missedMeetingNudge', 'call'],
      ['followedChannel', 'post', 'other'],
    ];
    const { data } = await call(kinds.map(([type, subtype], i) => entry({ type, subtype, ago: (i + 1) * 60_000 })));

    expect(rowsOf(data).map((r) => [r.type, r.subtype, r.kind])).toEqual(kinds);
  });

  it('returns rows newest first, whatever order the feed served them in', async () => {
    pinClock(NOW);
    const older = entry({ type: 'reply', ago: 3 * 3_600_000, preview: 'older' });
    const newer = entry({ type: 'reply', ago: 3_600_000, preview: 'newer' });
    const getMessages = vi.fn().mockResolvedValue([older, newer]);
    const harness = await createTestHarness((server) =>
      registerActivityTools(server, { chatsvc: fakeChatsvc([], { getMessages }) }),
    );

    const data = parseToolResult(await harness.callTool('teams_get_activity')) as Record<string, unknown>;

    expect(rowsOf(data).map((r) => r.preview)).toEqual(['newer', 'older']);
    await harness.close();
  });

  it('filters by kinds', async () => {
    pinClock(NOW);
    const { data } = await call([
      entry({ type: 'mention', subtype: 'person', ago: 60_000 }),
      entry({ type: 'reactionInChat', subtype: 'like', ago: 120_000 }),
      entry({ type: 'reply', ago: 180_000 }),
      entry({ type: 'msGraph', subtype: 'privateMeetingUpdated', ago: 240_000 }),
      entry({ type: 'call', subtype: 'missedMeetingNudge', ago: 300_000 }),
      entry({ type: 'unknownThing', ago: 360_000 }),
    ], { kinds: ['mention', 'call', 'other'] });

    expect(rowsOf(data).map((r) => r.kind)).toEqual(['mention', 'call', 'other']);
  });

  it('keeps only unread entries with unread_only', async () => {
    pinClock(NOW);
    const { data } = await call(
      [
        entry({ type: 'mention', read: true, ago: 60_000, preview: 'seen' }),
        entry({ type: 'mention', read: false, ago: 120_000, preview: 'not seen' }),
      ],
      { unread_only: true },
    );

    expect(rowsOf(data).map((r) => [r.preview, r.read])).toEqual([['not seen', false]]);
  });

  it('reports read as a boolean from properties.isread', async () => {
    pinClock(NOW);
    const { data } = await call([
      entry({ type: 'reply', read: true, ago: 60_000 }),
      entry({ type: 'reply', read: false, ago: 120_000 }),
    ]);

    expect(rowsOf(data).map((r) => r.read)).toEqual([true, false]);
  });

  it('only covers the last 14 days by default, and stops paging at the cutoff', async () => {
    pinClock(NOW);
    const { data, getMessages } = await call([
      entry({ type: 'reply', ago: 13 * DAY, preview: 'inside' }),
      entry({ type: 'reply', ago: 15 * DAY, preview: 'outside' }),
    ]);

    expect(rowsOf(data).map((r) => r.preview)).toEqual(['inside']);
    const opts = getMessages.mock.calls[0][1];
    expect(opts.stopAt?.(entry({ type: 'reply', ago: 15 * DAY }))).toBe(true);
    expect(opts.stopAt?.(entry({ type: 'reply', ago: 13 * DAY }))).toBe(false);
  });

  it('widens or narrows the window with since_days', async () => {
    pinClock(NOW);
    const entries = [
      entry({ type: 'reply', ago: 12 * 3_600_000, preview: 'today' }),
      entry({ type: 'reply', ago: 3 * DAY, preview: 'three days' }),
      entry({ type: 'reply', ago: 30 * DAY, preview: 'a month' }),
    ];

    const narrow = await call(entries, { since_days: 1 });
    const wide = await call(entries, { since_days: 60 });

    expect(rowsOf(narrow.data).map((r) => r.preview)).toEqual(['today']);
    expect(rowsOf(wide.data).map((r) => r.preview)).toEqual(['today', 'three days', 'a month']);
  });

  it('skips deleted entries', async () => {
    pinClock(NOW);
    const { data } = await call([
      entry({ type: 'mention', ago: 60_000, preview: 'kept' }),
      entry({ type: 'mention', ago: 120_000, preview: 'gone', deleted: true }),
    ]);

    expect(rowsOf(data).map((r) => r.preview)).toEqual(['kept']);
  });

  it('skips entries without an activity payload, or with one that does not parse', async () => {
    pinClock(NOW);
    const garbled = entry({ type: 'mention', ago: 180_000 });
    garbled.properties!.activity = '{not json';
    const { data } = await call([
      entry({ type: 'mention', ago: 60_000, preview: 'kept' }),
      entry({ type: 'mention', ago: 120_000, noActivity: true }),
      garbled,
    ]);

    expect(rowsOf(data).map((r) => r.preview)).toEqual(['kept']);
  });

  it('caps a preview at 200 characters of plain text', async () => {
    pinClock(NOW);
    const { data } = await call([entry({ type: 'reply', preview: `<div>${'x'.repeat(500)}</div>` })]);

    const preview = String(rowsOf(data)[0].preview);
    expect(preview.length).toBeLessThanOrEqual(200);
    expect(preview.endsWith('…')).toBe(true);
    expect(preview).not.toContain('<');
  });

  it('omits reply_chain_id when the entry has none', async () => {
    pinClock(NOW);
    const { data } = await call([entry({ type: 'mentionInChat', subtype: 'person' })]);

    expect(rowsOf(data)[0]).not.toHaveProperty('reply_chain_id');
  });

  it('keeps the newest rows up to limit and says how many were left out', async () => {
    pinClock(NOW);
    const entries = Array.from({ length: 60 }, (_, i) => entry({ type: 'reply', ago: (i + 1) * 60_000, preview: `item ${i}` }));

    const all = await call(entries);
    const one = await call(entries, { limit: 1 });

    expect(rowsOf(all.data)).toHaveLength(50);
    expect(all.data.truncated).toMatchObject({ total: 60, returned: 50 });
    expect(rowsOf(one.data).map((r) => r.preview)).toEqual(['item 0']);
  });

  it('wraps its result as untrusted third-party content', async () => {
    pinClock(NOW);
    const { data } = await call([entry({ type: 'mention' })]);

    expect(data.untrusted_content).toBe(true);
  });

  it('surfaces a token-capture failure with the click-into-a-chat hint', async () => {
    const chatsvc = fakeChatsvc([], {
      getMessages: vi.fn().mockRejectedValue(new Error('could not capture a Teams token from the browser (timed out).')),
    });
    const harness = await createTestHarness((server) => registerActivityTools(server, { chatsvc }));

    const result = await harness.callTool('teams_get_activity');

    expect(result.isError).toBe(true);
    const text = (result.content?.[0] as { text?: string })?.text ?? '';
    expect(text).toMatch(/Activity feed/);
    expect(text).toMatch(/click into any chat/i);
    await harness.close();
  });

  it('takes only optional arguments: unread_only, kinds, since_days and limit', async () => {
    const harness = await createTestHarness((server) => registerActivityTools(server, { chatsvc: fakeChatsvc() }));

    const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_get_activity');
    const schema = tool?.inputSchema as { properties?: Record<string, unknown>; required?: string[] };

    expect(Object.keys(schema.properties ?? {}).sort()).toEqual(['kinds', 'limit', 'since_days', 'unread_only']);
    expect(schema.required ?? []).toEqual([]);
    await harness.close();
  });

  it('is read-only and says it needs no Activity view and marks nothing read', async () => {
    const harness = await createTestHarness((server) => registerActivityTools(server, { chatsvc: fakeChatsvc() }));
    const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_get_activity');

    expect(tool?.annotations?.readOnlyHint).toBe(true);
    expect(tool?.description).toMatch(/teams_read_conversation/);
    expect(tool?.description).toMatch(/does NOT mark anything read/);
    expect(tool?.description).not.toMatch(/bell icon in the left nav\) open|requires the user's browser/i);
    await harness.close();
  });
});
