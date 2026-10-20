import { afterEach, beforeEach, describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerUnreadTools } from '../../src/tools/unread.js';
import { localIsoTime } from '../../src/model.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import { DAY, fakeChatsvc, messagesFrom, pinClock, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
/** "Now" for the since_days window: an hour after every fixture's activity. */
const NOW = T0 + 60 * MIN;

beforeEach(() => pinClock(NOW));
afterEach(() => vi.useRealTimers());

const GROUP = '19:group0001@thread.v2';
const READ_GROUP = '19:group0002@thread.v2';
const MINE_LAST = '19:group0003@thread.v2';
const MEETING = '19:meeting_abc0001@thread.v2';
const DM = '19:aaaa_bbbb@unq.gbl.spaces';
const CHANNEL = '19:chan0001@thread.tacv2';
const TEAM = '19:team0001@thread.tacv2';

/**
 * Unread: GROUP (last at +4, read to +2), DM (+6, read to +5), CHANNEL (+5,
 * read to +1), MEETING (+3, read to +2). Not unread: READ_GROUP (read to its
 * last message), MINE_LAST (the last message is the user's own), TEAM.
 */
const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * MIN, horizon: T0 + 2 * MIN }),
  rawConv(READ_GROUP, { topic: 'All caught up', at: T0 + 9 * MIN }),
  rawConv(MINE_LAST, { topic: 'I spoke last', at: T0 + 8 * MIN, horizon: T0, fromMe: true }),
  rawConv(MEETING, { topic: 'Weekly sync', at: T0 + 3 * MIN, horizon: T0 + 2 * MIN }),
  rawConv(DM, { at: T0 + 6 * MIN, horizon: T0 + 5 * MIN, sender: 'Example Friend' }),
  rawConv(CHANNEL, { topic: 'Design', at: T0 + 5 * MIN, horizon: T0 + 1 * MIN, spaceId: TEAM }),
  rawConv(TEAM, { topic: 'Platform Team', at: T0 }),
];

const MESSAGES: Record<string, RawMessage[]> = {
  [GROUP]: [
    rawMsg(T0 + 1 * MIN, { text: 'already read' }),
    rawMsg(T0 + 2 * MIN, { text: 'the horizon message' }),
    rawMsg(T0 + 2.5 * MIN, { type: 'ThreadActivity/AddMember', text: '<addmember/>' }),
    rawMsg(T0 + 3 * MIN, { text: 'group unread 1' }),
    rawMsg(T0 + 3.5 * MIN, { text: 'mine, after the horizon', fromMe: true }),
    rawMsg(T0 + 4 * MIN, { text: 'group unread 2' }),
  ],
  [MEETING]: [rawMsg(T0 + 2 * MIN, { conv: MEETING }), rawMsg(T0 + 3 * MIN, { conv: MEETING, text: 'meeting unread' })],
  [DM]: [rawMsg(T0 + 5 * MIN, { conv: DM }), rawMsg(T0 + 6 * MIN, { conv: DM, text: 'dm unread' })],
  [CHANNEL]: [
    rawMsg(T0 - 10 * MIN, { conv: CHANNEL, text: 'old root' }),
    rawMsg(T0 + 2 * MIN, { conv: CHANNEL, rootId: T0 - 10 * MIN, text: 'reply to old root' }),
    rawMsg(T0 + 3 * MIN, { conv: CHANNEL, rootId: T0 + 3 * MIN, text: 'new post' }),
    rawMsg(T0 + 5 * MIN, { conv: CHANNEL, rootId: T0 + 3 * MIN, text: 'reply to new post' }),
  ],
};

/** A `getMessages` that serves each conversation's own messages, paging like the real client. */
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

function errorText(result: { content?: unknown[] }): string {
  return (result.content?.[0] as { text?: string })?.text ?? '';
}

type Msg = { id: string; text: string; fromMe: boolean };
type Thread = { root: Msg | { id: string; missing: true }; replies: Msg[] };
type Entry = {
  conversation: { id: string; kind: string; title: string; team?: string };
  unread_count: number;
  messages?: Msg[];
  threads?: Thread[];
  truncated?: { returned: number; hint: string };
};

const ids = (data: Record<string, unknown> | undefined) => (data?.conversations as Entry[]).map((e) => e.conversation.id);
const entry = (data: Record<string, unknown> | undefined, id: string) =>
  (data?.conversations as Entry[]).find((e) => e.conversation.id === id)!;

describe('teams_list_unread_messages', () => {
  it('returns every unread conversation, most recent first, with only its unread messages', async () => {
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc);

    expect(ids(data)).toEqual([DM, CHANNEL, GROUP, MEETING]);
    expect(entry(data, GROUP).messages?.map((m) => m.text)).toEqual(['group unread 1', 'group unread 2']);
    expect(entry(data, GROUP).unread_count).toBe(2);
    expect(entry(data, DM).messages?.map((m) => m.text)).toEqual(['dm unread']);
    expect(entry(data, DM).conversation).toMatchObject({ kind: 'oneOnOne', title: 'Example Friend' });
    expect(data?.total_unread).toBe(2 + 1 + 3 + 1);
    expect(data).not.toHaveProperty('more_unread_conversations');
  });

  it('adds time_local in tz (or TEAMS_TIME_ZONE) to every unread message, chat and channel alike', async () => {
    type Timed = { time: string; time_local?: string };
    const timed = (data: Record<string, unknown> | undefined): Timed[] =>
      (data?.conversations as { messages?: Timed[]; threads?: { root: Timed | { missing: true }; replies: Timed[] }[] }[]).flatMap(
        (e) => e.messages ?? e.threads!.flatMap((t) => ['time' in t.root ? [t.root as Timed] : [], t.replies].flat()),
      );
    try {
      vi.stubEnv('TEAMS_TIME_ZONE', '');
      const plain = timed((await call(chatsvcWith())).data);
      expect(plain.length).toBeGreaterThan(0);
      expect(plain.every((m) => !('time_local' in m))).toBe(true);

      const ny = timed((await call(chatsvcWith(), { tz: 'America/New_York' })).data);
      expect(ny.map((m) => m.time_local)).toEqual(ny.map((m) => localIsoTime(m.time, 'America/New_York')));
      expect(ny[0].time_local).toMatch(/-04:00$/);

      vi.stubEnv('TEAMS_TIME_ZONE', 'Asia/Kolkata');
      const env = timed((await call(chatsvcWith())).data);
      expect(env.every((m) => m.time_local?.endsWith('+05:30'))).toBe(true);
    } finally {
      vi.unstubAllEnvs();
    }
  });

  it('rejects a tz Intl does not know', async () => {
    const { result } = await call(chatsvcWith(), { tz: 'Mars/Olympus_Mons' });
    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/IANA time zone/);
  });

  it('titles an unread untitled group chat by its participants', async () => {
    const UNTITLED = '19:untitled0001@thread.v2';
    const convs = [rawConv(UNTITLED, { at: T0 + 4 * MIN, horizon: T0, sender: 'King, Greg' })];
    const msgs = [
      { ...rawMsg(T0 + MIN, { conv: UNTITLED, sender: 'Alpha, Ann' }), from: 'https://x/contacts/8:orgid:ann' },
      rawMsg(T0 + 4 * MIN, { conv: UNTITLED, sender: 'King, Greg' }),
    ];
    const { data } = await call(fakeChatsvc(convs, { getMessages: perConversation({ [UNTITLED]: msgs }) }));
    const [c] = data?.conversations as { conversation: { title: string; participants?: string[] } }[];
    expect(c.conversation).toMatchObject({ title: 'Group chat: Alpha, Ann; King, Greg', participants: ['Alpha, Ann', 'King, Greg'] });
  });

  it('never fetches a conversation that is read up to date, or whose last message is the user\'s', async () => {
    const chatsvc = chatsvcWith();
    await call(chatsvc);

    const fetched = vi.mocked(chatsvc.getMessages).mock.calls.map((c) => c[0]);
    expect(fetched.sort()).toEqual([CHANNEL, DM, GROUP, MEETING].sort());
  });

  it('stops each read at that conversation\'s own read horizon', async () => {
    const chatsvc = chatsvcWith();
    await call(chatsvc);

    const optsFor = (id: string) =>
      vi.mocked(chatsvc.getMessages).mock.calls.find((c) => c[0] === id)![1] as GetMessagesOptions;
    expect(optsFor(GROUP).stopAt?.(rawMsg(T0 + 2 * MIN))).toBe(true);
    expect(optsFor(GROUP).stopAt?.(rawMsg(T0 + 3 * MIN))).toBe(false);
    expect(optsFor(DM).stopAt?.(rawMsg(T0 + 5 * MIN, { conv: DM }))).toBe(true);
  });

  it('groups a channel\'s unread posts into threads, marking an already-read root it cannot fetch as missing', async () => {
    const { data } = await call(chatsvcWith());

    const ch = entry(data, CHANNEL);
    expect(ch).not.toHaveProperty('messages');
    expect(ch.conversation).toMatchObject({ kind: 'channel', title: 'Design', team: 'Platform Team' });
    expect(ch.threads).toHaveLength(2);
    expect(ch.threads![0].root).toEqual({ id: String(T0 - 10 * MIN), missing: true });
    expect(ch.threads![0].replies.map((r) => r.text)).toEqual(['reply to old root']);
    expect((ch.threads![1].root as Msg).text).toBe('new post');
    expect(ch.threads![1].replies.map((r) => r.text)).toEqual(['reply to new post']);
    expect(ch.unread_count).toBe(3);
  });

  it('fills an already-read root from the service, marked outside_window', async () => {
    const getMessage = vi.fn(async (_c: string, id: string) =>
      id === String(T0 - 10 * MIN) ? rawMsg(T0 - 10 * MIN, { conv: CHANNEL, text: 'old root' }) : undefined,
    );
    const { data } = await call(chatsvcWith({ getMessage }));

    expect(getMessage).toHaveBeenCalledTimes(1);
    expect(getMessage).toHaveBeenCalledWith(CHANNEL, String(T0 - 10 * MIN));
    const ch = entry(data, CHANNEL);
    expect(ch.threads![0].root).toMatchObject({ id: String(T0 - 10 * MIN), text: 'old root', outside_window: true });
    expect(ch.threads![0].replies.map((r) => r.text)).toEqual(['reply to old root']);
    // The read root is not unread and is not counted.
    expect(ch.unread_count).toBe(3);
    expect(data?.total_unread).toBe(2 + 1 + 3 + 1);
  });

  it('filters by kind: chat covers group, meeting and 1:1 chats but not channels', async () => {
    const { data } = await call(chatsvcWith(), { kinds: ['chat'] });

    expect(ids(data)).toEqual([DM, GROUP, MEETING]);
  });

  it('filters by the narrower kinds', async () => {
    expect(ids((await call(chatsvcWith(), { kinds: ['channel'] })).data)).toEqual([CHANNEL]);
    expect(ids((await call(chatsvcWith(), { kinds: ['meeting', 'oneOnOne'] })).data)).toEqual([DM, MEETING]);
  });

  it('rejects an unknown kind', async () => {
    const { result } = await call(chatsvcWith(), { kinds: ['group'] });

    expect(result.isError).toBe(true);
  });

  it('filters by query on title or team', async () => {
    expect(ids((await call(chatsvcWith(), { query: 'falcon' })).data)).toEqual([GROUP]);
    expect(ids((await call(chatsvcWith(), { query: 'platform team' })).data)).toEqual([CHANNEL]);
  });

  it('says why each conversation matched a query', async () => {
    const reasons = (data: Record<string, unknown> | undefined) =>
      (data?.conversations as { conversation: { id: string; match_reason?: string } }[]).map((c) => [
        c.conversation.id,
        c.conversation.match_reason,
      ]);
    expect(reasons((await call(chatsvcWith(), { query: 'falcon' })).data)).toEqual([[GROUP, 'title']]);
    expect(reasons((await call(chatsvcWith(), { query: 'platform team' })).data)).toEqual([[CHANNEL, 'team']]);
    expect(reasons((await call(chatsvcWith(), { query: 'example friend' })).data)).toEqual([[DM, 'participant']]);
    expect(reasons((await call(chatsvcWith(), { query: MEETING })).data)).toEqual([[MEETING, 'id']]);
  });

  it('matches a conversation by its listing\'s last-message preview when no name does', async () => {
    const convs = CONVS.map((c) =>
      c.id === MEETING ? rawConv(MEETING, { topic: 'Weekly sync', at: T0 + 3 * MIN, horizon: T0 + 2 * MIN, text: 'Quarterly numbers are in' }) : c,
    );
    const { data } = await call(chatsvcWith({}, convs), { query: 'quarterly' });
    const conversations = data?.conversations as { conversation: { id: string; match_reason?: string } }[];

    expect(conversations.map((c) => [c.conversation.id, c.conversation.match_reason])).toEqual([[MEETING, 'message']]);
  });

  it('caps the conversations returned and lists the rest, confirmed, by id, title and unread_count', async () => {
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc, { max_conversations: 2 });

    expect(ids(data)).toEqual([DM, CHANNEL]);
    expect(data?.more_unread_conversations).toEqual([
      { id: GROUP, title: 'Project Falcon', unread_count: 2 },
      { id: MEETING, title: 'Weekly sync', unread_count: 1 },
    ]);
    // Every candidate is confirmed, so the rest are counted too.
    expect(data?.total_unread).toBe(1 + 3 + 2 + 1);
    expect(chatsvc.getMessages).toHaveBeenCalledTimes(4);
  });

  it('names an untitled group chat in more_unread_conversations by its participants', async () => {
    const UNTITLED = '19:untitled0001@thread.v2';
    const convs = [
      rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 9 * MIN, horizon: T0 }),
      rawConv(UNTITLED, { at: T0 + 4 * MIN, horizon: T0, sender: 'King, Greg' }),
    ];
    const msgs: Record<string, RawMessage[]> = {
      [GROUP]: [rawMsg(T0 + 9 * MIN, { conv: GROUP })],
      [UNTITLED]: [
        { ...rawMsg(T0 + MIN, { conv: UNTITLED, sender: 'Alpha, Ann' }), from: 'https://x/contacts/8:orgid:ann' },
        rawMsg(T0 + 4 * MIN, { conv: UNTITLED, sender: 'King, Greg' }),
      ],
    };
    const { data } = await call(fakeChatsvc(convs, { getMessages: perConversation(msgs) }), { max_conversations: 1 });
    expect(data?.more_unread_conversations).toEqual([
      { id: UNTITLED, title: 'Group chat: Alpha, Ann; King, Greg', participants: ['Alpha, Ann', 'King, Greg'], unread_count: 2 },
    ]);
  });

  it('never lists an unconfirmed candidate in more_unread_conversations', async () => {
    const recs = Array.from({ length: 3 }, (_, i) =>
      rawConv(`19:meeting_rec${i}@thread.v2`, { at: T0 + (20 + i) * MIN, horizon: T0, type: 'RichText/Media_CallRecording' }),
    );
    const recMsgs = Object.fromEntries(
      recs.map((c) => [c.id, [rawMsg(T0 + 15 * MIN, { conv: c.id, type: 'RichText/Media_CallRecording' })]]),
    );
    const { data } = await call(
      fakeChatsvc([...CONVS, ...recs], { getMessages: perConversation({ ...MESSAGES, ...recMsgs }) }),
      { max_conversations: 3 },
    );

    expect(ids(data)).toEqual([DM, CHANNEL, GROUP]);
    expect(data?.more_unread_conversations).toEqual([{ id: MEETING, title: 'Weekly sync', unread_count: 1 }]);
  });

  it('defaults to 20 conversations', async () => {
    const many = Array.from({ length: 25 }, (_, i) =>
      rawConv(`19:many${i}@thread.v2`, { topic: `Chat ${i}`, at: T0 + i * MIN, horizon: T0 - MIN }),
    );
    const manyMsgs = Object.fromEntries(many.map((c, i) => [c.id, [rawMsg(T0 + i * MIN, { conv: c.id })]]));
    const { data } = await call(fakeChatsvc(many, { getMessages: perConversation(manyMsgs) }));

    expect(data?.more_unread_conversations).toHaveLength(5);
    expect((data?.more_unread_conversations as { id: string }[])[0].id).toBe('19:many4@thread.v2');
  });

  it('keeps the newest per_conversation_limit unread messages and says more exist', async () => {
    const many = Array.from({ length: 8 }, (_, i) => rawMsg(T0 + (10 + i) * MIN, { text: `unread ${i}` }));
    const convs = [rawConv(GROUP, { topic: 'Busy', at: T0 + 17 * MIN, horizon: T0 })];
    const { data } = await call(fakeChatsvc(convs, { getMessages: perConversation({ [GROUP]: many }) }), {
      per_conversation_limit: 3,
    });

    const busy = entry(data, GROUP);
    expect(busy.messages?.map((m) => m.text)).toEqual(['unread 5', 'unread 6', 'unread 7']);
    expect(busy.truncated).toMatchObject({ returned: 3 });
    expect(busy.truncated?.hint).toMatch(/teams_list_messages/);
    expect(busy.unread_count).toBe(3);
  });

  it('defaults to 50 messages per conversation and adds no truncated field inside the limit', async () => {
    const sixty = Array.from({ length: 60 }, (_, i) => rawMsg(T0 + (10 + i) * MIN));
    const convs = [rawConv(GROUP, { at: T0 + 69 * MIN, horizon: T0 }), rawConv(DM, { at: T0 + 6 * MIN, horizon: T0 + 5 * MIN })];
    const { data } = await call(
      fakeChatsvc(convs, { getMessages: perConversation({ [GROUP]: sixty, [DM]: MESSAGES[DM] }) }),
    );

    expect(entry(data, GROUP).messages).toHaveLength(50);
    expect(entry(data, GROUP).truncated).toBeDefined();
    expect(entry(data, DM)).not.toHaveProperty('truncated');
  });

  it('rejects out-of-range limits', async () => {
    expect((await call(chatsvcWith(), { per_conversation_limit: 201 })).result.isError).toBe(true);
    expect((await call(chatsvcWith(), { max_conversations: 51 })).result.isError).toBe(true);
    expect((await call(chatsvcWith(), { max_conversations: 0 })).result.isError).toBe(true);
  });

  it('reads at most 4 conversations at once', async () => {
    const many = Array.from({ length: 10 }, (_, i) =>
      rawConv(`19:par${i}@thread.v2`, { at: T0 + i * MIN, horizon: T0 - MIN }),
    );
    let inFlight = 0;
    let peak = 0;
    const getMessages = vi.fn(async () => {
      inFlight++;
      peak = Math.max(peak, inFlight);
      await new Promise((r) => setTimeout(r, 5));
      inFlight--;
      return [];
    });
    await call(fakeChatsvc(many, { getMessages }));

    expect(getMessages).toHaveBeenCalledTimes(10);
    expect(peak).toBe(4);
  });

  it('drops a conversation whose unread messages turn out to be gone', async () => {
    const convs = [rawConv(GROUP, { at: T0 + 4 * MIN, horizon: T0 }), rawConv(DM, { at: T0 + 6 * MIN, horizon: T0 + 5 * MIN })];
    const deleted = [rawMsg(T0 + 4 * MIN, { deleted: true })];
    const { data } = await call(
      fakeChatsvc(convs, { getMessages: perConversation({ [GROUP]: deleted, [DM]: MESSAGES[DM] }) }),
    );

    expect(ids(data)).toEqual([DM]);
  });

  it('says total_unread is exact when nothing was capped or truncated', async () => {
    const { data } = await call(chatsvcWith());

    expect(data?.total_unread_exact).toBe(true);
  });

  it('says total_unread is exact when conversations were capped but every one was confirmed', async () => {
    const { data } = await call(chatsvcWith(), { max_conversations: 2 });

    expect(data?.total_unread_exact).toBe(true);
  });

  it('says total_unread is not exact when a per-conversation limit truncated', async () => {
    const { data } = await call(chatsvcWith(), { per_conversation_limit: 1 });

    expect(data?.total_unread_exact).toBe(false);
  });

  it('records a conversation whose read failed and carries on with the rest', async () => {
    const pager = perConversation();
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => {
      if (id === GROUP) throw new McpToolError('Teams chat service is rate-limiting requests (429).', { status: 429 });
      return pager(id, opts);
    });
    const { result, data } = await call(chatsvcWith({ getMessages }));

    expect(result.isError).toBeFalsy();
    expect(ids(data)).toEqual([DM, CHANNEL, MEETING]);
    const errors = data?.errors as { conversation: { id: string; title: string }; error: string }[];
    expect(errors).toHaveLength(1);
    expect(errors[0].conversation).toMatchObject({ id: GROUP, title: 'Project Falcon' });
    expect(errors[0].error).toMatch(/rate-limiting/);
    expect(data?.total_unread).toBe(1 + 3 + 1);
    expect(data?.total_unread_exact).toBe(false);
  });

  it('reads a recording-only meeting chat (lastimreceivedtime equal to the recording, as live) and leaves it out', async () => {
    const REC = '19:meeting_rec0001@thread.v2';
    const convs = [
      rawConv(REC, {
        topic: 'Recorded sync',
        at: T0 + 9 * MIN,
        horizon: T0 + 2 * MIN,
        lastIm: T0 + 9 * MIN,
        type: 'RichText/Media_CallRecording',
      }),
    ];
    const getMessages = perConversation({
      [REC]: [
        rawMsg(T0 + 2 * MIN, { conv: REC }),
        rawMsg(T0 + 5 * MIN, { conv: REC, type: 'Event/Call' }),
        rawMsg(T0 + 9 * MIN, { conv: REC, type: 'RichText/Media_CallRecording' }),
      ],
    });
    const { data } = await call(fakeChatsvc(convs, { getMessages }));

    expect(data).toMatchObject({ total_unread: 0, conversations: [], skipped_unreachable: 0 });
    expect(getMessages).toHaveBeenCalledTimes(1);
  });

  it('leaves out a chat whose stale lastMessage is real but whose news past the horizon is only system events (live)', async () => {
    const STALE = '19:stale0001@thread.v2';
    const convs = [
      rawConv(STALE, { topic: 'Stale', at: T0 + 1 * MIN, horizon: T0 + 2 * MIN, lastIm: T0 + 6 * MIN, type: 'RichText/Html' }),
    ];
    const getMessages = perConversation({
      [STALE]: [
        rawMsg(T0 + 1 * MIN, { conv: STALE, type: 'RichText/Html' }),
        rawMsg(T0 + 4 * MIN, { conv: STALE, type: 'Event/Call' }),
        rawMsg(T0 + 6 * MIN, { conv: STALE, type: 'ThreadActivity/DeleteMember' }),
      ],
    });
    const { data } = await call(fakeChatsvc(convs, { getMessages }));

    expect(data).toMatchObject({ total_unread: 0, conversations: [] });
    expect(getMessages).toHaveBeenCalledTimes(1);
  });

  it('reads a never-opened meeting chat (no read horizon) and returns its real messages from others', async () => {
    const NEW = '19:meeting_new0001@thread.v2';
    const never = rawConv(NEW, { topic: 'Never opened', at: T0 + 9 * MIN, type: 'Event/Call' });
    delete never.properties!.consumptionhorizon;
    const getMessages = perConversation({
      [NEW]: [
        rawMsg(T0 + 1 * MIN, { conv: NEW, type: 'ThreadActivity/AddMember' }),
        rawMsg(T0 + 3 * MIN, { conv: NEW, text: 'agenda attached' }),
        rawMsg(T0 + 4 * MIN, { conv: NEW, text: 'mine', fromMe: true }),
        rawMsg(T0 + 9 * MIN, { conv: NEW, type: 'Event/Call' }),
      ],
    });
    const { data } = await call(fakeChatsvc([never], { getMessages }));

    expect(ids(data)).toEqual([NEW]);
    expect(entry(data, NEW).messages?.map((m) => m.text)).toEqual(['agenda attached']);
    expect(entry(data, NEW).conversation).toMatchObject({ unread: true, unread_confirmed: true });
  });

  it('skips a channel that answers 404 or 403 silently, counting it in skipped_unreachable (live)', async () => {
    const gone = new McpToolError('Teams chat service returned HTTP 404.', { status: 404, kind: 'http' });
    const denied = new McpToolError('Teams chat service returned HTTP 403.', { status: 403, kind: 'http' });
    const pager = perConversation();
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => {
      if (id === CHANNEL) throw gone;
      if (id === MEETING) throw denied;
      return pager(id, opts);
    });
    const { result, data } = await call(chatsvcWith({ getMessages }));

    expect(result.isError).toBeFalsy();
    expect(ids(data)).toEqual([DM, GROUP]);
    expect(data?.skipped_unreachable).toBe(2);
    expect(data).not.toHaveProperty('errors');
    expect(data?.total_unread_exact).toBe(true);
  });

  it('returns an empty result, not an error, when every candidate is unreachable', async () => {
    const gone = new McpToolError('Teams chat service returned HTTP 404.', { status: 404, kind: 'http' });
    const { result, data } = await call(chatsvcWith({ getMessages: vi.fn().mockRejectedValue(gone) }));

    expect(result.isError).toBeFalsy();
    expect(data).toMatchObject({ total_unread: 0, conversations: [], skipped_unreachable: 4 });
  });

  it('reads a meeting chat whose only news is a call recording or transcript, and leaves it out', async () => {
    const convs = [
      rawConv(MEETING, {
        topic: 'Recorded sync',
        at: T0 + 9 * MIN,
        horizon: T0 + 2 * MIN,
        lastIm: T0 + 2 * MIN,
        type: 'RichText/Media_CallRecording',
      }),
      rawConv(GROUP, {
        topic: 'Transcribed',
        at: T0 + 8 * MIN,
        horizon: T0 + 2 * MIN,
        lastIm: T0 + 1 * MIN,
        type: 'RichText/Media_CallTranscript',
      }),
    ];
    const getMessages = perConversation({
      [MEETING]: [rawMsg(T0 + 9 * MIN, { conv: MEETING, type: 'RichText/Media_CallRecording' })],
      [GROUP]: [rawMsg(T0 + 8 * MIN, { type: 'RichText/Media_CallTranscript' })],
    });
    const { data } = await call(fakeChatsvc(convs, { getMessages }));

    expect(data).toMatchObject({ total_unread: 0, conversations: [] });
  });

  it('reads a chat whose last message is a recording but whose last real message is unread', async () => {
    const convs = [
      rawConv(GROUP, {
        topic: 'Project Falcon',
        at: T0 + 9 * MIN,
        horizon: T0 + 2 * MIN,
        lastIm: T0 + 4 * MIN,
        type: 'RichText/Media_CallRecording',
      }),
    ];
    const { data } = await call(fakeChatsvc(convs, { getMessages: perConversation() }));

    expect(ids(data)).toEqual([GROUP]);
    expect(entry(data, GROUP).unread_count).toBe(2);
  });

  describe('since_days', () => {
    const OLD = '19:old0001@thread.v2';
    const DORMANT = '19:dormant0001@thread.tacv2';
    /** Active an hour ago, but the only real unread message is 20 days old; a call event is the recent news. */
    const old = rawConv(OLD, { topic: 'Old news', at: NOW - 20 * DAY, horizon: NOW - 30 * DAY, lastIm: NOW - 1 * MIN });
    /** Nothing at all for over a year. */
    const dormant = rawConv(DORMANT, { topic: 'Dormant', at: NOW - 400 * DAY, horizon: NOW - 500 * DAY, spaceId: TEAM });
    const msgs = {
      ...MESSAGES,
      [OLD]: [
        rawMsg(NOW - 20 * DAY, { conv: OLD, text: 'old unread' }),
        rawMsg(NOW - 1 * MIN, { conv: OLD, type: 'Event/Call' }),
      ],
      [DORMANT]: [rawMsg(NOW - 400 * DAY, { conv: DORMANT, rootId: NOW - 400 * DAY, text: 'dormant post' })],
    };
    const convs = [...CONVS, old, dormant];

    it('defaults to 14 days: keeps in-window unread, skips and counts older unread, with a hint', async () => {
      const { data } = await call(fakeChatsvc(convs, { getMessages: perConversation(msgs) }));

      expect(ids(data)).toEqual([DM, CHANNEL, GROUP, MEETING]);
      expect(data?.older_unread_skipped).toBe(2);
      expect(String(data?.older_unread_hint)).toMatch(/14 days/);
      expect(String(data?.older_unread_hint)).toMatch(/since_days/);
      expect(data?.total_unread).toBe(7);
    });

    it('never reads a candidate whose newest activity is older than the window', async () => {
      const getMessages = perConversation(msgs);
      await call(fakeChatsvc(convs, { getMessages }));

      const fetched = getMessages.mock.calls.map((c) => c[0]);
      expect(fetched).not.toContain(DORMANT);
      expect(fetched).toContain(OLD);
    });

    it('a larger since_days brings older unread back', async () => {
      const month = await call(fakeChatsvc(convs, { getMessages: perConversation(msgs) }), { since_days: 30 });
      expect(ids(month.data)).toContain(OLD);
      expect(ids(month.data)).not.toContain(DORMANT);
      expect(month.data?.older_unread_skipped).toBe(1);

      const all = await call(fakeChatsvc(convs, { getMessages: perConversation(msgs) }), { since_days: 3650 });
      expect(ids(all.data)).toEqual(expect.arrayContaining([OLD, DORMANT]));
      expect(all.data?.older_unread_skipped).toBe(0);
      expect(all.data).not.toHaveProperty('older_unread_hint');
    });

    it('a smaller since_days narrows the window', async () => {
      const recent = rawConv(GROUP, { topic: 'Project Falcon', at: NOW - 2 * DAY, horizon: NOW - 3 * DAY });
      const getMessages = perConversation({ [GROUP]: [rawMsg(NOW - 2 * DAY, { text: 'two days ago' })] });
      const { data } = await call(fakeChatsvc([recent], { getMessages }), { since_days: 1 });

      expect(ids(data)).toEqual([]);
      expect(data?.older_unread_skipped).toBe(1);
      expect(getMessages).not.toHaveBeenCalled();
    });

    it('with older_unread, lists what the window left out: titled, newest first, counted only where read', async () => {
      const getMessages = perConversation(msgs);
      const { data } = await call(fakeChatsvc(convs, { getMessages }), { older_unread: true });

      expect(data?.older_unread_skipped).toBe(2);
      expect(data?.older_unread_conversations).toEqual([
        { id: OLD, kind: 'group', title: 'Old news', lastActivity: new Date(NOW - 20 * DAY).toISOString(), unread_count: 1 },
        {
          id: DORMANT,
          kind: 'channel',
          title: 'Dormant',
          team: 'Platform Team',
          lastActivity: new Date(NOW - 400 * DAY).toISOString(),
          unread_count_known: false,
        },
      ]);
      // Listing the dormant channel never reads past its marker.
      expect(getMessages.mock.calls.map((c) => c[0])).not.toContain(DORMANT);
    });

    it('titles an untitled group chat the window left out', async () => {
      const UNTITLED = '19:untitled0001@thread.v2';
      const c = rawConv(UNTITLED, { at: NOW - 20 * DAY, horizon: NOW - 30 * DAY, sender: 'Example Person' });
      const getMessages = perConversation({ [UNTITLED]: [rawMsg(NOW - 20 * DAY, { conv: UNTITLED, sender: 'Example Person' })] });
      const { data } = await call(fakeChatsvc([c], { getMessages }), { older_unread: true });

      expect((data?.older_unread_conversations as { title: string }[])[0].title).toBe('Group chat: Example Person');
    });

    it('omits older_unread_conversations without older_unread, and when nothing was left out', async () => {
      const without = await call(fakeChatsvc(convs, { getMessages: perConversation(msgs) }));
      expect(without.data).not.toHaveProperty('older_unread_conversations');

      const none = await call(fakeChatsvc(convs, { getMessages: perConversation(msgs) }), { older_unread: true, since_days: 3650 });
      expect(none.data).not.toHaveProperty('older_unread_conversations');
    });

    it('rejects since_days outside 1..3650', async () => {
      for (const since_days of [0, 3651, 1.5]) {
        const { result } = await call(chatsvcWith(), { since_days });
        expect(result.isError).toBe(true);
      }
    });
  });

  describe('hidden channels', () => {
    const HIDDEN = '19:hidden0001@thread.tacv2';
    const convs = [
      ...CONVS,
      rawConv(HIDDEN, { topic: 'Archive', at: T0 + 7 * MIN, horizon: T0, spaceId: TEAM, interest: 'NotInterested' }),
    ];
    const msgs = { ...MESSAGES, [HIDDEN]: [rawMsg(T0 + 7 * MIN, { conv: HIDDEN, rootId: T0 + 7 * MIN, text: 'hidden post' })] };

    it('leaves channels hidden in Teams out by default, without fetching them', async () => {
      const getMessages = perConversation(msgs);
      const { data } = await call(fakeChatsvc(convs, { getMessages }));

      expect(ids(data)).not.toContain(HIDDEN);
      expect(getMessages.mock.calls.map((c) => c[0])).not.toContain(HIDDEN);
    });

    it('includes them with include_hidden', async () => {
      const { data } = await call(fakeChatsvc(convs, { getMessages: perConversation(msgs) }), { include_hidden: true });

      expect(ids(data)).toContain(HIDDEN);
      expect(entry(data, HIDDEN).conversation).toMatchObject({ hidden: true });
    });
  });

  describe('muted chats', () => {
    const MUTED = '19:muted0001@thread.v2';
    const convs = [...CONVS, rawConv(MUTED, { topic: 'Noisy', at: T0 + 7 * MIN, horizon: T0, alerts: 'false' })];
    const msgs = { ...MESSAGES, [MUTED]: [rawMsg(T0 + 7 * MIN, { conv: MUTED, text: 'muted unread' })] };

    it('leaves chats muted in Teams out by default, without reading them, counting them in muted_skipped', async () => {
      const getMessages = perConversation(msgs);
      const { data } = await call(fakeChatsvc(convs, { getMessages }));

      expect(ids(data)).not.toContain(MUTED);
      expect(data?.muted_skipped).toBe(1);
      expect(getMessages.mock.calls.map((c) => c[0])).not.toContain(MUTED);
    });

    it('includes them with include_muted, marked muted', async () => {
      const { data } = await call(fakeChatsvc(convs, { getMessages: perConversation(msgs) }), { include_muted: true });

      expect(ids(data)).toContain(MUTED);
      expect(entry(data, MUTED).conversation).toMatchObject({ muted: true });
      expect(data?.muted_skipped).toBe(0);
    });

    it('never counts a muted chat that is not an unread candidate', async () => {
      const read = [...CONVS, rawConv(MUTED, { topic: 'Noisy', at: T0 + 7 * MIN, alerts: 'false' })];
      const { data } = await call(fakeChatsvc(read, { getMessages: perConversation(msgs) }));

      expect(data?.muted_skipped).toBe(0);
    });

    it('counts only muted candidates the kinds and query keep', async () => {
      const { data } = await call(fakeChatsvc(convs, { getMessages: perConversation(msgs) }), { query: 'Falcon' });

      expect(data?.muted_skipped).toBe(0);
    });
  });

  it('returns an empty result when nothing is unread', async () => {
    const { data } = await call(fakeChatsvc([rawConv(READ_GROUP, { topic: 'All caught up' })]));

    expect(data).toMatchObject({ total_unread: 0, conversations: [] });
  });

  it('returns its result inside the untrusted-data envelope', async () => {
    const { data } = await call(chatsvcWith());

    expect(data?.untrusted_content).toBe(true);
  });

  it('turns a token-capture failure into the open-a-Teams-tab hint', async () => {
    const chatsvc = fakeChatsvc([], {
      listConversations: vi.fn().mockRejectedValue(new Error('could not capture a Teams token from the browser (timed out).')),
    });
    const { result } = await call(chatsvc);

    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/click into any chat/i);
  });

  it('passes an already-actionable API error through unchanged when every read failed', async () => {
    const chatsvc = chatsvcWith({
      getMessages: vi.fn().mockRejectedValue(new McpToolError('Teams chat service is rate-limiting requests (429).', { status: 429 })),
    });
    const { result } = await call(chatsvc);

    expect(errorText(result)).toMatch(/rate-limiting/);
  });

  it('describes itself as cross-conversation, not marking anything read unless asked', async () => {
    const harness = await createTestHarness((server) => registerUnreadTools(server, { chatsvc: fakeChatsvc() }));
    const tools = (await harness.client.listTools()).tools;
    await harness.close();

    expect(tools.map((t) => t.name)).toEqual(['teams_list_unread_messages']);
    const tool = tools[0];
    expect(tool.description).toMatch(/every unread message/i);
    expect(tool.description).toMatch(/by default this does NOT mark anything read/i);
    expect(tool.description).toMatch(/badges are unaffected/i);
    expect(tool.description).toMatch(/mark_read/);
    expect(tool.description).toMatch(/teams_mark_read/);
    expect(tool.description).toMatch(/refused with format "summary"/);
    expect(tool.description).toMatch(/skipped: "truncated"/);
    expect(tool.description).not.toMatch(/marks the older unread messages it left out read too/);
    // mark_read makes it a write — a non-destructive, idempotent one.
    expect(tool.annotations).toMatchObject({ readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true });
    expect(tool.description).toMatch(/call recording/i);
    expect(tool.description).toMatch(/hidden/i);
    expect(tool.description).toMatch(/skipped_unreachable/);
    expect(tool.description).not.toMatch(/uncertain/);
    expect(Object.keys((tool.inputSchema as { properties?: object }).properties ?? {}).sort()).toEqual([
      'confirmToken',
      'cursors',
      'format',
      'from',
      'include_hidden',
      'include_muted',
      'kinds',
      'limit',
      'mark_read',
      'max_conversations',
      'older_unread',
      'per_conversation_limit',
      'query',
      'since_days',
      'tz',
    ]);
    expect(tool.description).toMatch(/time_local/);
    expect(tool.description).toMatch(/since_days/);
    expect(tool.description).toMatch(/older_unread_skipped/);
    expect(tool.description).toMatch(/older_unread_conversations/);
    expect(tool.description).toMatch(/muted_skipped/);
    expect(tool.description).toMatch(/include_muted/);
  });
});

describe('teams_list_unread_messages cursors', () => {
  const at = (min: number) => String(T0 + min * MIN);
  type CursorEntry = Entry & { next_cursor?: string };

  it('returns only the unread messages newer than a conversation\'s cursor', async () => {
    const { data } = await call(chatsvcWith(), { cursors: { [GROUP]: at(3) } });

    expect(entry(data, GROUP).messages?.map((m) => m.text)).toEqual(['group unread 2']);
    expect(entry(data, GROUP).unread_count).toBe(1);
    expect((entry(data, GROUP) as CursorEntry).next_cursor).toBe(at(4));
    // Conversations without a cursor are read down to their horizon as before.
    expect(entry(data, DM).messages?.map((m) => m.text)).toEqual(['dm unread']);
  });

  it('stops each read at the later of its read horizon and its cursor', async () => {
    const chatsvc = chatsvcWith();
    await call(chatsvc, { cursors: { [GROUP]: at(3), [DM]: at(1) } });

    const optsFor = (id: string) =>
      vi.mocked(chatsvc.getMessages).mock.calls.find((c) => c[0] === id)![1] as GetMessagesOptions;
    expect(optsFor(GROUP).stopAt?.(rawMsg(T0 + 3 * MIN))).toBe(true);
    expect(optsFor(GROUP).stopAt?.(rawMsg(T0 + 4 * MIN))).toBe(false);
    // A cursor below the horizon: the horizon still ends the read.
    expect(optsFor(DM).stopAt?.(rawMsg(T0 + 5 * MIN, { conv: DM }))).toBe(true);
  });

  it('leaves out a conversation with nothing new past its cursor, and keeps its cursor in next_cursors', async () => {
    const { data } = await call(chatsvcWith(), { cursors: { [GROUP]: at(4) } });

    expect(ids(data)).toEqual([DM, CHANNEL, MEETING]);
    expect(data?.total_unread).toBe(1 + 3 + 1);
    expect((data?.next_cursors as Record<string, string>)[GROUP]).toBe(at(4));
  });

  it('a cursor below the read horizon changes nothing', async () => {
    const { data } = await call(chatsvcWith(), { cursors: { [GROUP]: at(1) } });

    expect(entry(data, GROUP).messages?.map((m) => m.text)).toEqual(['group unread 1', 'group unread 2']);
  });

  it('gives every returned conversation next_cursor, its newest returned message id, even without cursors', async () => {
    const { data } = await call(chatsvcWith());

    expect((entry(data, GROUP) as CursorEntry).next_cursor).toBe(at(4));
    expect((entry(data, CHANNEL) as CursorEntry).next_cursor).toBe(at(5));
    expect((entry(data, DM) as CursorEntry).next_cursor).toBe(at(6));
  });

  it('returns next_cursors: the passed cursors updated by every returned conversation, ready to pass back', async () => {
    const UNKNOWN = '19:notalisted@thread.v2';
    const { data } = await call(chatsvcWith(), { cursors: { [GROUP]: at(3), [UNKNOWN]: '123' } });

    expect(data?.next_cursors).toEqual({
      [GROUP]: at(4),
      [UNKNOWN]: '123',
      [DM]: at(6),
      [CHANNEL]: at(5),
      [MEETING]: at(3),
    });

    // Passed straight back, nothing is new.
    const again = await call(chatsvcWith(), { cursors: data?.next_cursors });
    expect(ids(again.data)).toEqual([]);
    expect(again.data?.total_unread).toBe(0);
    expect(again.data?.next_cursors).toEqual(data?.next_cursors);
  });

  it('never advances next_cursors for a conversation only listed in more_unread_conversations', async () => {
    const { data } = await call(chatsvcWith(), { max_conversations: 1 });

    expect(ids(data)).toEqual([DM]);
    expect(Object.keys(data?.next_cursors as object)).toEqual([DM]);
  });

  it('with per_conversation_limit keeps the newest past the cursor and moves the cursor to the newest', async () => {
    const many = Array.from({ length: 8 }, (_, i) => rawMsg(T0 + (10 + i) * MIN, { text: `unread ${i}` }));
    const convs = [rawConv(GROUP, { topic: 'Busy', at: T0 + 17 * MIN, horizon: T0 })];
    const { data } = await call(fakeChatsvc(convs, { getMessages: perConversation({ [GROUP]: many }) }), {
      per_conversation_limit: 2,
      cursors: { [GROUP]: at(12) },
    });

    const busy = entry(data, GROUP) as CursorEntry;
    expect(busy.messages?.map((m) => m.text)).toEqual(['unread 6', 'unread 7']);
    expect(busy.truncated).toMatchObject({ returned: 2 });
    expect(busy.next_cursor).toBe(at(17));
  });

  it('is no truncation when the cursor, not the limit, ended the read', async () => {
    const many = Array.from({ length: 8 }, (_, i) => rawMsg(T0 + (10 + i) * MIN, { text: `unread ${i}` }));
    const convs = [rawConv(GROUP, { topic: 'Busy', at: T0 + 17 * MIN, horizon: T0 })];
    const { data } = await call(fakeChatsvc(convs, { getMessages: perConversation({ [GROUP]: many }) }), {
      per_conversation_limit: 2,
      cursors: { [GROUP]: at(15) },
    });

    expect(entry(data, GROUP)).not.toHaveProperty('truncated');
    expect(data?.total_unread_exact).toBe(true);
  });

  it('with since_days, new unread past the cursor but older than the window is still left out and counted', async () => {
    const OLD = '19:old0001@thread.v2';
    const old = rawConv(OLD, { topic: 'Old news', at: NOW - 20 * DAY, horizon: NOW - 30 * DAY, lastIm: NOW - 1 * MIN });
    const getMessages = perConversation({
      [OLD]: [rawMsg(NOW - 25 * DAY, { conv: OLD }), rawMsg(NOW - 20 * DAY, { conv: OLD, text: 'old unread' })],
    });
    const { data } = await call(fakeChatsvc([old], { getMessages }), { cursors: { [OLD]: String(NOW - 25 * DAY) } });

    expect(ids(data)).toEqual([]);
    expect(data?.older_unread_skipped).toBe(1);
  });

  it('rejects a cursor that is not a message id', async () => {
    for (const cursors of [{ [GROUP]: 'yesterday' }, { [GROUP]: '' }, { [GROUP]: 5 }, ['1']]) {
      const { result } = await call(chatsvcWith(), { cursors });
      expect(result.isError).toBe(true);
    }
  });

  it('stamps the result with fetched_at, the time of this live read', async () => {
    const { data } = await call(chatsvcWith());

    expect(data?.fetched_at).toBe(new Date(NOW).toISOString());
  });

  it('documents cursors, next_cursor and that nothing is cached', async () => {
    const harness = await createTestHarness((server) => registerUnreadTools(server, { chatsvc: fakeChatsvc() }));
    const tools = (await harness.client.listTools()).tools;
    await harness.close();

    const tool = tools[0];
    expect(tool.description).toMatch(/cursors/);
    expect(tool.description).toMatch(/next_cursor/);
    expect(tool.description).toMatch(/fetched_at/);
    expect(tool.description).toMatch(/no cache|not cached|never cached/i);
  });
});

describe('teams_list_unread_messages — mark_read', () => {
  const ENV_KEYS = ['MCP_CONFIRM_MODE', 'MCP_CONFIRM_TTL_SECONDS', 'MCP_CONFIRM_SECRET'] as const;
  let savedEnv: Record<string, string | undefined>;
  beforeEach(() => {
    savedEnv = Object.fromEntries(ENV_KEYS.map((k) => [k, process.env[k]]));
    for (const k of ENV_KEYS) delete process.env[k];
  });
  afterEach(() => {
    for (const k of ENV_KEYS) {
      if (savedEnv[k] === undefined) delete process.env[k];
      else process.env[k] = savedEnv[k];
    }
  });

  /** Phase 1, then (when phase 1 asked for one) phase 2 with its token. */
  async function callConfirmed(chatsvc: ChatServiceClient, args: Record<string, unknown>) {
    const harness = await createTestHarness((server) => registerUnreadTools(server, { chatsvc }));
    try {
      const first = await harness.callTool('teams_list_unread_messages', args);
      const one = first.isError ? undefined : (parseToolResult(first) as Record<string, unknown>);
      if (typeof one?.confirmToken !== 'string') return { preview: undefined, data: one };
      const second = await harness.callTool('teams_list_unread_messages', { ...args, confirmToken: one.confirmToken });
      return { preview: one, data: second.isError ? undefined : (parseToolResult(second) as Record<string, unknown>) };
    } finally {
      await harness.close();
    }
  }

  type Marked = Entry & { marked_read?: boolean; mark_read_warning?: string };
  const marked = (data: Record<string, unknown> | undefined, id: string) =>
    (data?.conversations as Marked[]).find((e) => e.conversation.id === id)!;

  it('never writes a read marker by default', async () => {
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const { data } = await call(chatsvcWith({ setConsumptionHorizon }));
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(entry(data, GROUP)).not.toHaveProperty('marked_read');
  });

  it('previews every move first, then marks each returned conversation up to its newest unread message', async () => {
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const { preview, data } = await callConfirmed(chatsvcWith({ setConsumptionHorizon }), { mark_read: true });

    expect(preview).not.toHaveProperty('conversations');
    const shown = JSON.stringify((preview?.preview as Record<string, unknown>).conversations);
    for (const id of [GROUP, DM, CHANNEL, MEETING]) expect(shown).toContain(id);

    expect(setConsumptionHorizon.mock.calls.sort()).toEqual(
      [
        [GROUP, String(T0 + 4 * MIN)],
        [DM, String(T0 + 6 * MIN)],
        [CHANNEL, String(T0 + 5 * MIN)],
        [MEETING, String(T0 + 3 * MIN)],
      ].sort(),
    );
    for (const id of [GROUP, DM, CHANNEL, MEETING]) expect(marked(data, id).marked_read).toBe(true);
  });

  it('marks only the conversations returned with their messages, not more_unread_conversations', async () => {
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const { data } = await callConfirmed(chatsvcWith({ setConsumptionHorizon }), { mark_read: true, max_conversations: 1 });
    expect(ids(data)).toEqual([DM]);
    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(DM, String(T0 + 6 * MIN));
  });

  it('asks nothing and writes nothing when nothing is unread', async () => {
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const { preview, data } = await callConfirmed(chatsvcWith({ setConsumptionHorizon }), {
      mark_read: true,
      query: 'All caught up',
    });
    expect(preview).toBeUndefined();
    expect(data?.conversations).toEqual([]);
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
  });

  it('keeps marking the others when one move fails, reporting it per conversation', async () => {
    const setConsumptionHorizon = vi.fn(async (id: string) => {
      if (id === DM) throw new McpToolError('Teams chat service returned HTTP 403.', { status: 403 });
    });
    const { data } = await callConfirmed(chatsvcWith({ setConsumptionHorizon }), { mark_read: true });
    expect(marked(data, DM).marked_read).toBe(false);
    expect(marked(data, DM).mark_read_warning).toMatch(/403/);
    expect(marked(data, GROUP).marked_read).toBe(true);
    expect((data?.conversations as Entry[]).length).toBe(4);
  });

  it('with cursors, marks only conversations with something new past the cursor, and next_cursors still advance', async () => {
    const setConsumptionHorizon = vi.fn(async (_id: string, _target: string) => undefined);
    const { data } = await callConfirmed(chatsvcWith({ setConsumptionHorizon }), {
      mark_read: true,
      cursors: { [GROUP]: String(T0 + 4 * MIN) },
    });

    expect(setConsumptionHorizon.mock.calls.map((c) => c[0])).not.toContain(GROUP);
    expect(marked(data, DM).marked_read).toBe(true);
    expect((data?.next_cursors as Record<string, string>)[GROUP]).toBe(String(T0 + 4 * MIN));
  });

  it('refuses mark_read with format summary, reading and marking nothing', async () => {
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const chatsvc = chatsvcWith({ setConsumptionHorizon });
    const { result } = await call(chatsvc, { mark_read: true, format: 'summary' });

    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/mark_read/);
    expect(errorText(result)).toMatch(/summary/);
    expect(chatsvc.getMessages).not.toHaveBeenCalled();
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
  });

  it('with format compact, marks each returned conversation up to its newest unread message', async () => {
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const { data } = await callConfirmed(chatsvcWith({ setConsumptionHorizon }), { mark_read: true, format: 'compact' });

    expect(entry(data, GROUP).messages!.map((m) => m.text)).toContain('group unread 2');
    expect(marked(data, GROUP).marked_read).toBe(true);
    expect(setConsumptionHorizon).toHaveBeenCalledWith(GROUP, String(T0 + 4 * MIN));
  });

  it('leaves the marker alone for a conversation per_conversation_limit cut, marking the rest', async () => {
    const setConsumptionHorizon = vi.fn(async (_id: string, _target: string) => undefined);
    // GROUP and CHANNEL hold more unread messages than one; DM and MEETING hold exactly one.
    const { preview, data } = await callConfirmed(chatsvcWith({ setConsumptionHorizon }), {
      mark_read: true,
      per_conversation_limit: 1,
    });

    const shown = JSON.stringify((preview?.preview as Record<string, unknown>).conversations);
    for (const id of [GROUP, CHANNEL]) expect(shown).not.toContain(id);
    expect(setConsumptionHorizon.mock.calls.map((c) => c[0]).sort()).toEqual([DM, MEETING].sort());
    for (const id of [GROUP, CHANNEL]) {
      expect(entry(data, id).truncated).toBeDefined();
      expect(marked(data, id).marked_read).toEqual({
        skipped: 'truncated',
        hint: expect.stringMatching(/per_conversation_limit/),
      });
    }
    for (const id of [DM, MEETING]) expect(marked(data, id).marked_read).toBe(true);
  });

  it('asks nothing and writes nothing when every returned conversation was truncated', async () => {
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const { preview, data } = await callConfirmed(chatsvcWith({ setConsumptionHorizon }), {
      mark_read: true,
      per_conversation_limit: 1,
      kinds: ['channel'],
    });
    expect(preview).toBeUndefined();
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(marked(data, CHANNEL).marked_read).toMatchObject({ skipped: 'truncated' });
  });

  it('with cursors and nothing new past them, asks nothing and writes nothing', async () => {
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const { preview, data } = await callConfirmed(chatsvcWith({ setConsumptionHorizon }), {
      mark_read: true,
      kinds: ['chat'],
      query: 'Falcon',
      cursors: { [GROUP]: String(T0 + 4 * MIN) },
    });
    expect(preview).toBeUndefined();
    expect(data?.conversations).toEqual([]);
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
  });

  it('never lets an outside_window thread root drive the mark target', async () => {
    const getMessage = vi.fn(async (_c: string, id: string) =>
      id === String(T0 - 10 * MIN) ? rawMsg(T0 - 10 * MIN, { conv: CHANNEL, text: 'old root' }) : undefined,
    );
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const { data } = await callConfirmed(chatsvcWith({ getMessage, setConsumptionHorizon }), {
      mark_read: true,
      kinds: ['channel'],
    });

    expect(entry(data, CHANNEL).threads![0].root).toMatchObject({ outside_window: true });
    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(CHANNEL, String(T0 + 5 * MIN));
  });
});
