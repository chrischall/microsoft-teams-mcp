import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerUnreadTools } from '../../src/tools/unread.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import { fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;

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
    const result = await harness.callTool('teams_get_unread', args);
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

describe('teams_get_unread', () => {
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

  it('groups a channel\'s unread posts into threads, marking an already-read root as missing', async () => {
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
    expect(busy.truncated?.hint).toMatch(/teams_read_conversation/);
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

  it('describes itself as read-only, cross-conversation, and never marking anything read', async () => {
    const harness = await createTestHarness((server) => registerUnreadTools(server, { chatsvc: fakeChatsvc() }));
    const tools = (await harness.client.listTools()).tools;
    await harness.close();

    expect(tools.map((t) => t.name)).toEqual(['teams_get_unread']);
    const tool = tools[0];
    expect(tool.description).toMatch(/every unread message/i);
    expect(tool.description).toMatch(/does not mark/i);
    expect(tool.description).toMatch(/badges are unaffected/i);
    expect(tool.annotations?.readOnlyHint).toBe(true);
    expect(tool.annotations?.openWorldHint).toBe(true);
    expect(tool.description).toMatch(/call recording/i);
    expect(tool.description).toMatch(/hidden/i);
    expect(tool.description).toMatch(/skipped_unreachable/);
    expect(tool.description).not.toMatch(/uncertain/);
    expect(Object.keys((tool.inputSchema as { properties?: object }).properties ?? {}).sort()).toEqual([
      'include_hidden',
      'kinds',
      'max_conversations',
      'per_conversation_limit',
      'query',
    ]);
  });
});
