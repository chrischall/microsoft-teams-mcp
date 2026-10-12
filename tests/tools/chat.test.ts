import { afterEach, beforeEach, describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerChatTools } from '../../src/tools/chat.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import { DAY, fakeChatsvc, ME, messagesFrom, pinClock, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
/** "Now" for the since_days window: an hour after every fixture's activity. */
const NOW = T0 + 60 * MIN;

beforeEach(() => pinClock(NOW));
afterEach(() => vi.useRealTimers());

const GROUP = '19:group0001@thread.v2';
const MEETING = '19:meeting_abc0001@thread.v2';
const DM = '19:aaaa_bbbb@unq.gbl.spaces';
const CHANNEL = '19:chan0001@thread.tacv2';

/** Unread group (newest), read meeting, unread 1:1 (oldest), and a channel that must not appear. */
const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 30 * MIN, horizon: T0 }),
  rawConv(MEETING, { topic: 'Weekly sync', at: T0 + 20 * MIN }),
  rawConv(DM, { at: T0 + 10 * MIN, horizon: T0, sender: 'Pat Example' }),
  rawConv(CHANNEL, { topic: 'General', at: T0 + 40 * MIN, horizon: T0 }),
];

/** What a read past each horizon finds: GROUP and DM hold real unread messages. */
const MESSAGES: Record<string, RawMessage[]> = {
  [GROUP]: [rawMsg(T0), rawMsg(T0 + 30 * MIN, { text: 'group unread' })],
  [DM]: [rawMsg(T0, { conv: DM }), rawMsg(T0 + 5 * MIN, { conv: DM }), rawMsg(T0 + 10 * MIN, { conv: DM })],
};

/** A `getMessages` that serves each conversation's own messages; an `Error`-valued entry is thrown instead. */
function perConversation(byId: Record<string, RawMessage[] | Error> = MESSAGES) {
  return vi.fn(async (id: string, opts: GetMessagesOptions) => {
    const v = byId[id] ?? [];
    if (v instanceof Error) throw v;
    return messagesFrom(v)(id, opts);
  });
}

/** `untitled`'s reads go to `untitled`, the fixture's own to {@link MESSAGES}. */
function perConversationWith(untitled: (id: string, opts: GetMessagesOptions) => Promise<RawMessage[]>) {
  const own = perConversation();
  return vi.fn(async (id: string, opts: GetMessagesOptions) => (id in MESSAGES ? own(id, opts) : untitled(id, opts)));
}

function chatsvcWith(convs = CONVS, byId: Record<string, RawMessage[] | Error> = MESSAGES): ChatServiceClient {
  return fakeChatsvc(convs, { getMessages: perConversation(byId) });
}

async function call(chatsvc: ChatServiceClient, args: Record<string, unknown> = {}) {
  const harness = await createTestHarness((server) => registerChatTools(server, { chatsvc }));
  try {
    const result = await harness.callTool('teams_list_chats', args);
    return { result, data: result.isError ? undefined : (parseToolResult(result) as Record<string, unknown>) };
  } finally {
    await harness.close();
  }
}

type Row = { id: string; kind: string; title: string; unread: boolean; unread_confirmed: boolean; unread_count?: number };

describe('teams_list_chats', () => {
  it('lists every chat (group, meeting, 1:1) from the API, newest activity first, without channels', async () => {
    const { data } = await call(fakeChatsvc(CONVS));
    const chats = data?.chats as Row[];

    expect(chats.map((c) => c.id)).toEqual([GROUP, MEETING, DM]);
    expect(chats.map((c) => c.kind)).toEqual(['group', 'meeting', 'oneOnOne']);
    expect(chats.find((c) => c.id === DM)?.title).toBe('Pat Example');
    expect(chats.map((c) => c.unread)).toEqual([true, false, true]);
  });

  it('marks every row\'s unread flag as an unconfirmed hint without unread_only, and reads no messages', async () => {
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc);

    expect((data?.chats as Row[]).every((c) => c.unread_confirmed === false)).toBe(true);
    expect((data?.chats as Row[]).some((c) => 'unread_count' in c)).toBe(false);
    expect(data).not.toHaveProperty('skipped_unreachable');
    expect(chatsvc.getMessages).not.toHaveBeenCalled();
  });

  it('keeps only chats a read confirms unread with unread_only, with their unread_count', async () => {
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc, { unread_only: true });
    const chats = data?.chats as Row[];

    expect(chats.map((c) => c.id)).toEqual([GROUP, DM]);
    expect(chats.map((c) => c.unread_count)).toEqual([1, 2]);
    expect(chats.every((c) => c.unread && c.unread_confirmed)).toBe(true);
    expect(data?.skipped_unreachable).toBe(0);
    // Only the candidates are read, each down to its own horizon, one small page at most.
    const calls = vi.mocked(chatsvc.getMessages).mock.calls;
    expect(calls.map((c) => c[0]).sort()).toEqual([DM, GROUP].sort());
    expect(calls.every((c) => (c[1] as GetMessagesOptions).max === 50)).toBe(true);
    expect((calls[0][1] as GetMessagesOptions).stopAt?.(rawMsg(T0))).toBe(true);
  });

  it('drops a recording-only meeting chat whose lastimreceivedtime equals the recording (live fact)', async () => {
    const REC = '19:meeting_rec0001@thread.v2';
    const recorded = rawConv(REC, {
      topic: 'Recorded sync',
      at: T0 + 50 * MIN,
      horizon: T0 + 45 * MIN,
      lastIm: T0 + 50 * MIN,
      type: 'RichText/Media_CallRecording',
    });
    const msgs = {
      ...MESSAGES,
      [REC]: [
        rawMsg(T0 + 45 * MIN, { conv: REC }),
        rawMsg(T0 + 48 * MIN, { conv: REC, type: 'Event/Call' }),
        rawMsg(T0 + 50 * MIN, { conv: REC, type: 'RichText/Media_CallRecording' }),
      ],
    };
    const chatsvc = chatsvcWith([...CONVS, recorded], msgs);
    const plain = await call(chatsvc);
    const { data } = await call(chatsvc, { unread_only: true });

    // The listing alone flags it — a hint only.
    expect((plain.data?.chats as Row[]).find((c) => c.id === REC)).toMatchObject({ unread: true, unread_confirmed: false });
    expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP, DM]);
  });

  it('drops a chat whose stale lastMessage is real but everything after the horizon is a system event (live fact)', async () => {
    const STALE = '19:stale0001@thread.v2';
    const stale = rawConv(STALE, {
      topic: 'Stale',
      at: T0 + 25 * MIN,
      horizon: T0 + 30 * MIN,
      lastIm: T0 + 35 * MIN,
      type: 'RichText/Html',
    });
    const msgs = {
      ...MESSAGES,
      [STALE]: [
        rawMsg(T0 + 25 * MIN, { conv: STALE, type: 'RichText/Html' }),
        rawMsg(T0 + 30 * MIN, { conv: STALE }),
        rawMsg(T0 + 33 * MIN, { conv: STALE, type: 'Event/Call' }),
        rawMsg(T0 + 35 * MIN, { conv: STALE, type: 'ThreadActivity/DeleteMember' }),
      ],
    };
    const { data } = await call(chatsvcWith([...CONVS, stale], msgs), { unread_only: true });

    expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP, DM]);
  });

  it('confirms a never-opened chat (no read horizon) holding a real message from someone else', async () => {
    const NEW = '19:meeting_new0001@thread.v2';
    const never = rawConv(NEW, { topic: 'Never opened', at: T0 + 60 * MIN, type: 'Event/Call' });
    delete never.properties!.consumptionhorizon;
    const msgs = {
      ...MESSAGES,
      [NEW]: [
        rawMsg(T0 + 55 * MIN, { conv: NEW, type: 'ThreadActivity/AddMember' }),
        rawMsg(T0 + 58 * MIN, { conv: NEW, text: 'welcome' }),
        rawMsg(T0 + 60 * MIN, { conv: NEW, type: 'Event/Call' }),
      ],
    };
    const { data } = await call(chatsvcWith([...CONVS, never], msgs), { unread_only: true });

    expect((data?.chats as Row[]).map((c) => c.id)).toEqual([NEW, GROUP, DM]);
    expect((data?.chats as Row[])[0]).toMatchObject({ unread_count: 1, unread_confirmed: true });
  });

  it('drops a never-opened chat holding only system events and the user\'s own messages', async () => {
    const NEW = '19:meeting_new0002@thread.v2';
    const never = rawConv(NEW, { topic: 'Never opened', at: T0 + 60 * MIN, type: 'Event/Call' });
    delete never.properties!.consumptionhorizon;
    const msgs = {
      ...MESSAGES,
      [NEW]: [rawMsg(T0 + 58 * MIN, { conv: NEW, fromMe: true }), rawMsg(T0 + 60 * MIN, { conv: NEW, type: 'Event/Call' })],
    };
    const { data } = await call(chatsvcWith([...CONVS, never], msgs), { unread_only: true });

    expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP, DM]);
  });

  it('skips a candidate that answers 404 or 403 silently, counting it in skipped_unreachable', async () => {
    const gone = new McpToolError('Teams chat service returned HTTP 404.', { status: 404, kind: 'http' });
    const denied = new McpToolError('Teams chat service returned HTTP 403.', { status: 403, kind: 'http' });
    const { result, data } = await call(chatsvcWith(CONVS, { ...MESSAGES, [GROUP]: gone, [DM]: denied }), {
      unread_only: true,
    });

    expect(result.isError).toBeFalsy();
    expect(data).toMatchObject({ chats: [], skipped_unreachable: 2 });
    expect(data).not.toHaveProperty('errors');
  });

  it('reports a candidate whose read failed otherwise in errors and keeps the rest', async () => {
    const limited = new McpToolError('Teams chat service is rate-limiting requests (429).', { status: 429 });
    const { data } = await call(chatsvcWith(CONVS, { ...MESSAGES, [GROUP]: limited }), { unread_only: true });

    expect((data?.chats as Row[]).map((c) => c.id)).toEqual([DM]);
    expect(data?.errors).toEqual([
      { conversation: { id: GROUP, title: 'Project Falcon' }, error: expect.stringMatching(/rate-limiting/) },
    ]);
  });

  it('fails the call when every candidate read failed', async () => {
    const limited = new McpToolError('Teams chat service is rate-limiting requests (429).', { status: 429 });
    const { result } = await call(chatsvcWith(CONVS, { [GROUP]: limited, [DM]: limited }), { unread_only: true });

    expect(result.isError).toBe(true);
    expect((result.content?.[0] as { text?: string })?.text).toMatch(/rate-limiting/);
  });

  it('says unread_count is a lower bound when more than one page past the horizon exists', async () => {
    const busy = Array.from({ length: 60 }, (_, i) => rawMsg(T0 + (i + 1) * 1000));
    const { data } = await call(chatsvcWith(CONVS, { ...MESSAGES, [GROUP]: busy }), { unread_only: true });
    const group = (data?.chats as (Row & { unread_count_exact?: boolean })[]).find((c) => c.id === GROUP);

    expect(group).toMatchObject({ unread_count: 50, unread_count_exact: false });
  });

  it('filters by a case-insensitive title substring with query', async () => {
    const { data } = await call(fakeChatsvc(CONVS), { query: 'falcon' });

    expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP]);
  });

  it('matches an exact conversation id with query', async () => {
    const { data } = await call(fakeChatsvc(CONVS), { query: MEETING });

    expect((data?.chats as Row[]).map((c) => c.id)).toEqual([MEETING]);
  });

  it('sorts a chat with no messages last', async () => {
    const empty = { id: '19:empty0001@thread.v2', threadProperties: { topic: 'Quiet' } };
    const { data } = await call(fakeChatsvc([empty, ...CONVS]));

    expect((data?.chats as Row[]).at(-1)?.id).toBe('19:empty0001@thread.v2');
  });

  it('returns the 50 most recent by default and honours limit, saying what was left out', async () => {
    const many = Array.from({ length: 60 }, (_, i) => rawConv(`19:g${i}@thread.v2`, { topic: `chat ${i}`, at: T0 + i * MIN }));

    const all = (await call(fakeChatsvc(many))).data!;
    const two = (await call(fakeChatsvc(many), { limit: 2 })).data!;

    expect(all.chats).toHaveLength(50);
    expect((all.chats as Row[])[0].title).toBe('chat 59');
    expect(all.truncated).toMatchObject({ total: 60, returned: 50 });
    expect((two.chats as Row[]).map((c) => c.title)).toEqual(['chat 59', 'chat 58']);
  });

  it('adds no truncated field when everything fits', async () => {
    const { data } = await call(fakeChatsvc(CONVS));

    expect(data).not.toHaveProperty('truncated');
  });

  describe('untitled group chats', () => {
    const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;
    const UNTITLED = Array.from({ length: 15 }, (_, i) => `19:untitled${String(i).padStart(4, '0')}@thread.v2`);
    const convs = UNTITLED.map((id, i) => rawConv(id, { at: T0 + i * MIN, sender: 'King, Greg' }));
    const other = (i: number) => `Member${String.fromCharCode(65 + i)}, Example`;
    const byId: Record<string, RawMessage[]> = Object.fromEntries(
      UNTITLED.map((id, i) => [
        id,
        [
          { ...rawMsg(T0 - 2 * MIN, { sender: other(i), conv: id }), from: FROM(`8:orgid:member${i}`) },
          rawMsg(T0 - MIN, { fromMe: true, conv: id }),
          rawMsg(T0 + i * MIN, { sender: 'King, Greg', conv: id }),
        ],
      ]),
    );

    it('titles each by its participants, so a shared last sender no longer makes them indistinguishable', async () => {
      const { data } = await call(chatsvcWith(convs, byId), { query: 'King' });
      const chats = data?.chats as (Row & { participants?: string[] })[];

      expect(chats).toHaveLength(15);
      expect(new Set(chats.map((c) => c.title)).size).toBe(15);
      expect(chats[0]).toMatchObject({
        id: UNTITLED[14],
        title: `Group chat: King, Greg; ${other(14)}`,
        participants: ['King, Greg', other(14)],
      });
    });

    it('matches query against participant names', async () => {
      const { data } = await call(chatsvcWith(convs, byId), { query: other(3) });
      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([UNTITLED[3]]);
    });

    it('reads only the chats it returns when there is no query', async () => {
      const chatsvc = chatsvcWith(convs, byId);
      const { data } = await call(chatsvc, { limit: 2 });

      expect((data?.chats as Row[]).map((c) => c.title)).toEqual([
        `Group chat: King, Greg; ${other(14)}`,
        `Group chat: King, Greg; ${other(13)}`,
      ]);
      expect(chatsvc.getMessages).toHaveBeenCalledTimes(2);
    });

    it('titles an untitled group chat named in unread_only errors by its participants', async () => {
      const id = UNTITLED[0];
      const unread = [rawConv(id, { at: T0 + 5 * MIN, horizon: T0, sender: 'King, Greg' })];
      const limited = new McpToolError('Teams chat service is rate-limiting requests (429).', { status: 429 });
      // The confirming read is refused; the later title read succeeds.
      let calls = 0;
      const getMessages = vi.fn(async (cid: string, opts: GetMessagesOptions) => {
        if (calls++ === 0) throw limited;
        return messagesFrom(byId[cid] ?? [])(cid, opts);
      });
      const { data } = await call(fakeChatsvc([...unread, ...CONVS], { getMessages: perConversationWith(getMessages) }), {
        unread_only: true,
      });

      expect(data?.errors).toEqual(
        expect.arrayContaining([
          { conversation: { id, title: `Group chat: King, Greg; ${other(0)}` }, error: expect.stringMatching(/rate-limiting/) },
        ]),
      );
    });

    it('keeps the last-sender title for a chat whose read failed, and still lists the rest', async () => {
      const { data } = await call(chatsvcWith(convs, { ...byId, [UNTITLED[14]]: new Error('boom') }), { limit: 2 });
      expect((data?.chats as Row[]).map((c) => c.title)).toEqual(['King, Greg', `Group chat: King, Greg; ${other(13)}`]);
    });
  });

  describe('1:1 chats', () => {
    const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;
    const PEER_OID = '00000000-0000-0000-0000-0000000000d4';
    const KING_DM = `19:${PEER_OID}_${ME.slice('8:orgid:'.length)}@unq.gbl.spaces`;
    const kingSays = (at: number): RawMessage => ({ ...rawMsg(at, { sender: 'King, Greg', conv: KING_DM }), from: FROM(`8:orgid:${PEER_OID}`) });
    // The user wrote last — the listing alone cannot say who this 1:1 is with.
    const convs = [...CONVS, rawConv(KING_DM, { at: T0 + 50 * MIN, fromMe: true, favorite: 'True' })];
    const byId = { ...MESSAGES, [KING_DM]: [rawMsg(T0 + 50 * MIN, { fromMe: true, conv: KING_DM }), kingSays(T0 + 45 * MIN)] };

    it('finds a 1:1 whose last message is the user\'s by the other person\'s name (live: a favorite 1:1 vanished)', async () => {
      const { data } = await call(chatsvcWith(convs, byId), { query: 'King' });
      expect(data?.chats).toEqual([
        expect.objectContaining({ id: KING_DM, kind: 'oneOnOne', title: 'King, Greg', participants: ['King, Greg'], favorite: true }),
      ]);
    });

    it('titles it by the other person in a plain listing', async () => {
      const { data } = await call(chatsvcWith(convs, byId));
      expect((data?.chats as Row[]).find((c) => c.id === KING_DM)?.title).toBe('King, Greg');
    });
  });

  describe('favorites and pinned chats', () => {
    const FAV = '19:fav0001@thread.v2';
    const PIN = '19:pin0001@thread.v2';
    const convs = [
      ...CONVS,
      rawConv(FAV, { topic: 'Favorite chat', at: T0 + MIN, favorite: 'true' }),
      rawConv(PIN, { topic: 'Pinned chat', at: T0 + 2 * MIN, pinned: 'True' }),
    ];

    it('lists favorites first, then pinned, then the rest by recency, flagging them', async () => {
      const { data } = await call(fakeChatsvc(convs));
      const chats = data?.chats as (Row & { favorite?: true; pinned?: true })[];
      expect(chats.map((c) => c.id)).toEqual([FAV, PIN, GROUP, MEETING, DM]);
      expect(chats[0]).toMatchObject({ favorite: true });
      expect(chats[0]).not.toHaveProperty('pinned');
      expect(chats[1]).toMatchObject({ pinned: true });
      expect(chats[2]).not.toHaveProperty('favorite');
    });

    it('keeps a favorite when limit trims the listing', async () => {
      const { data } = await call(fakeChatsvc(convs), { limit: 1 });
      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([FAV]);
    });

    it('keeps unread_only by recency, carrying the flags', async () => {
      const unreadFav = rawConv(FAV, { topic: 'Favorite chat', at: T0 + MIN, horizon: T0, favorite: 'true' });
      const { data } = await call(chatsvcWith([...CONVS, unreadFav], { ...MESSAGES, [FAV]: [rawMsg(T0 + MIN, { conv: FAV })] }), {
        unread_only: true,
      });
      const chats = data?.chats as (Row & { favorite?: true })[];
      expect(chats.map((c) => c.id)).toEqual([GROUP, DM, FAV]);
      expect(chats[2]).toMatchObject({ favorite: true });
    });
  });

  describe('since / until', () => {
    const iso = (ms: number) => new Date(ms).toISOString();
    const SILENT = '19:silent0001@thread.v2';
    const { lastMessage: _drop, ...silent } = rawConv(SILENT, { topic: 'No messages yet' });
    const convs = [...CONVS, silent];

    it('keeps chats whose last activity is at or after since', async () => {
      const { data } = await call(fakeChatsvc(convs), { since: iso(T0 + 20 * MIN) });
      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP, MEETING]);
    });

    it('keeps chats whose last activity is at or before until', async () => {
      const { data } = await call(fakeChatsvc(convs), { until: iso(T0 + 20 * MIN) });
      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([MEETING, DM]);
    });

    it('combines since and until into one window', async () => {
      const { data } = await call(fakeChatsvc(convs), { since: iso(T0 + 15 * MIN), until: iso(T0 + 25 * MIN) });
      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([MEETING]);
    });

    it('lists a chat with no last activity only when neither is given', async () => {
      const plain = await call(fakeChatsvc(convs));
      expect((plain.data?.chats as Row[]).map((c) => c.id)).toContain(SILENT);
      const { data } = await call(fakeChatsvc(convs), { since: iso(0) });
      expect((data?.chats as Row[]).map((c) => c.id)).not.toContain(SILENT);
    });

    it('filters before confirming unread_only, so chats outside the window are never read', async () => {
      const chatsvc = chatsvcWith();
      const { data } = await call(chatsvc, { unread_only: true, since: iso(T0 + 20 * MIN) });

      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP]);
      expect(vi.mocked(chatsvc.getMessages).mock.calls.map((c) => c[0])).toEqual([GROUP]);
    });

    it('rejects a since or until that is not an ISO 8601 date-time', async () => {
      for (const args of [{ since: 'yesterday' }, { until: 'not a date' }]) {
        const { result } = await call(fakeChatsvc(convs), args);
        expect(result.isError).toBe(true);
        expect((result.content?.[0] as { text?: string })?.text).toMatch(/ISO 8601/);
      }
    });

    it('rejects a since later than until', async () => {
      const { result } = await call(fakeChatsvc(convs), { since: iso(T0 + 30 * MIN), until: iso(T0) });
      expect(result.isError).toBe(true);
      expect((result.content?.[0] as { text?: string })?.text).toMatch(/since .*after until/i);
    });
  });

  describe('sort', () => {
    const FAV = '19:fav0001@thread.v2';
    const PIN = '19:pin0001@thread.v2';
    const LOWER = '19:lower0001@thread.v2';
    const convs = [
      ...CONVS,
      rawConv(FAV, { topic: 'Favorite chat', at: T0 + MIN, favorite: 'true' }),
      rawConv(PIN, { topic: 'Pinned chat', at: T0 + 2 * MIN, pinned: 'True' }),
      rawConv(LOWER, { topic: 'alpha team', at: T0 + 3 * MIN }),
    ];

    it('defaults to activity: favorites, pinned, then most recent', async () => {
      const dflt = await call(fakeChatsvc(convs));
      const explicit = await call(fakeChatsvc(convs), { sort: 'activity' });
      expect((dflt.data?.chats as Row[]).map((c) => c.id)).toEqual([FAV, PIN, GROUP, MEETING, DM, LOWER]);
      expect((explicit.data?.chats as Row[]).map((c) => c.id)).toEqual([FAV, PIN, GROUP, MEETING, DM, LOWER]);
    });

    it('activity_only orders by most recent activity alone, ignoring favorites and pinned', async () => {
      const { data } = await call(fakeChatsvc(convs), { sort: 'activity_only' });
      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP, MEETING, DM, LOWER, PIN, FAV]);
    });

    it('title orders alphabetically, case-insensitively, by the title returned (an untitled 1:1 by its person)', async () => {
      const { data } = await call(fakeChatsvc(convs), { sort: 'title' });
      const chats = data?.chats as Row[];
      expect(chats.map((c) => c.title)).toEqual([
        'alpha team',
        'Favorite chat',
        'Pat Example',
        'Pinned chat',
        'Project Falcon',
        'Weekly sync',
      ]);
    });

    it('limit keeps the head of the chosen order', async () => {
      const { data } = await call(fakeChatsvc(convs), { sort: 'title', limit: 2 });
      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([LOWER, FAV]);
      expect(data?.truncated).toMatchObject({ total: 6, returned: 2 });
    });

    it('applies to unread_only too', async () => {
      const { data } = await call(chatsvcWith(), { unread_only: true, sort: 'title' });
      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([DM, GROUP]);
    });

    it('rejects an unknown sort', async () => {
      const { result } = await call(fakeChatsvc(convs), { sort: 'newest' });
      expect(result.isError).toBe(true);
    });
  });

  it('returns its result inside the untrusted-data envelope', async () => {
    const { data } = await call(fakeChatsvc(CONVS));

    expect(data?.untrusted_content).toBe(true);
  });

  it('turns a token-capture failure into the open-a-Teams-tab hint', async () => {
    const chatsvc = fakeChatsvc([], {
      listConversations: vi
        .fn()
        .mockRejectedValue(new Error('could not capture a Teams token from the browser (timed out). Open a signed-in …')),
    });

    const { result } = await call(chatsvc);

    expect(result.isError).toBe(true);
    const text = (result.content?.[0] as { text?: string })?.text ?? '';
    expect(text).toMatch(/teams\.cloud\.microsoft/i);
    expect(text).toMatch(/click into any chat/i);
  });

  it('passes an already-actionable API error through unchanged', async () => {
    const chatsvc = fakeChatsvc([], {
      listConversations: vi.fn().mockRejectedValue(
        new McpToolError('Teams chat service is rate-limiting requests (429).', { status: 429, hint: 'Wait a minute before retrying.' }),
      ),
    });

    const { result } = await call(chatsvc);

    const text = (result.content?.[0] as { text?: string })?.text ?? '';
    expect(text).toMatch(/rate-limiting/);
    expect(text).toMatch(/Wait a minute/);
  });

  describe('since_days (with unread_only)', () => {
    const OLD = '19:old0001@thread.v2';
    const DORMANT = '19:dormant0001@thread.v2';
    const old = rawConv(OLD, { topic: 'Old news', at: NOW - 20 * DAY, horizon: NOW - 30 * DAY, lastIm: NOW - 1 * MIN });
    const dormant = rawConv(DORMANT, { topic: 'Dormant', at: NOW - 400 * DAY, horizon: NOW - 500 * DAY });
    const msgs = {
      ...MESSAGES,
      [OLD]: [rawMsg(NOW - 20 * DAY, { conv: OLD }), rawMsg(NOW - 1 * MIN, { conv: OLD, type: 'Event/Call' })],
      [DORMANT]: [rawMsg(NOW - 400 * DAY, { conv: DORMANT })],
    };
    const convs = [...CONVS, old, dormant];

    it('defaults to 14 days, counting what it left out and never reading a dormant chat', async () => {
      const chatsvc = chatsvcWith(convs, msgs);
      const { data } = await call(chatsvc, { unread_only: true });

      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP, DM]);
      expect(data?.older_unread_skipped).toBe(2);
      expect(String(data?.older_unread_hint)).toMatch(/since_days/);
      expect(vi.mocked(chatsvc.getMessages).mock.calls.map((c) => c[0])).not.toContain(DORMANT);
    });

    it('a larger since_days brings older unread chats back', async () => {
      const { data } = await call(chatsvcWith(convs, msgs), { unread_only: true, since_days: 3650 });

      expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP, DM, OLD, DORMANT]);
      expect(data?.older_unread_skipped).toBe(0);
      expect(data).not.toHaveProperty('older_unread_hint');
    });

    it('changes nothing without unread_only', async () => {
      const chatsvc = chatsvcWith(convs, msgs);
      const { data } = await call(chatsvc, { since_days: 1 });

      expect((data?.chats as Row[]).map((c) => c.id)).toContain(DORMANT);
      expect(data).not.toHaveProperty('older_unread_skipped');
      expect(chatsvc.getMessages).not.toHaveBeenCalled();
    });
  });

  it('says it covers every chat and that ids feed teams_read_conversation', async () => {
    const harness = await createTestHarness((server) => registerChatTools(server, { chatsvc: fakeChatsvc() }));
    const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_list_chats');

    expect(tool?.description).toMatch(/every chat/i);
    expect(tool?.description).toMatch(/teams_read_conversation/);
    expect(tool?.description).toMatch(/unread_confirmed: false/);
    expect(tool?.description).toMatch(/teams_get_unread/);
    expect(tool?.annotations?.readOnlyHint).toBe(true);
    expect(Object.keys((tool?.inputSchema as { properties?: object }).properties ?? {}).sort()).toEqual([
      'limit',
      'query',
      'since',
      'since_days',
      'sort',
      'unread_only',
      'until',
    ]);
    expect(tool?.description).toMatch(/since_days/);
    await harness.close();
  });

  it('registers only teams_list_chats (the open-chat DOM reader is gone)', async () => {
    const harness = await createTestHarness((server) => registerChatTools(server, { chatsvc: fakeChatsvc() }));

    expect((await harness.client.listTools()).tools.map((t) => t.name)).toEqual(['teams_list_chats']);
    await harness.close();
  });
});
