import { afterEach, beforeEach, describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerConversationTools, resolveListed } from '../../src/tools/conversation.js';
import type { ChatServiceClient, GetMessagesOptions } from '../../src/chatsvc.js';
import { fakeChatsvc, ME, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;

const GROUP = '19:group0001@thread.v2';
const OTHER_GROUP = '19:group0002@thread.v2';
const CHANNEL = '19:chan0001@thread.tacv2';
const TEAM = '19:team0001@thread.tacv2';

/** Read up to T0 + 2 min: the messages at +3 and +4 min are unread. */
const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * MIN, horizon: T0 + 2 * MIN }),
  rawConv(OTHER_GROUP, { topic: 'Falcon launch party', at: T0 }),
  rawConv(CHANNEL, { topic: 'Design', at: T0 + 4 * MIN, horizon: T0 + 2 * MIN, spaceId: TEAM }),
  rawConv(TEAM, { topic: 'Platform Team', at: T0 }),
];

const GROUP_MSGS = [
  rawMsg(T0 + 1 * MIN, { text: 'first' }),
  rawMsg(T0 + 2 * MIN, { text: 'second', fromMe: true }),
  rawMsg(T0 + 2.5 * MIN, { type: 'ThreadActivity/AddMember', text: '<addmember/>' }),
  rawMsg(T0 + 3 * MIN, { text: 'third' }),
  rawMsg(T0 + 3.5 * MIN, { text: 'mine, after the horizon', fromMe: true }),
  rawMsg(T0 + 4 * MIN, { text: 'fourth' }),
];

function chatsvcWith(messages = GROUP_MSGS, overrides: Record<string, unknown> = {}): ChatServiceClient {
  return fakeChatsvc(CONVS, { getMessages: messagesFrom(messages), ...overrides });
}

async function call(chatsvc: ChatServiceClient, args: Record<string, unknown>) {
  const harness = await createTestHarness((server) => registerConversationTools(server, { chatsvc }));
  try {
    const result = await harness.callTool('teams_read_conversation', args);
    return { result, data: result.isError ? undefined : (parseToolResult(result) as Record<string, unknown>) };
  } finally {
    await harness.close();
  }
}

function errorText(result: { content?: unknown[] }): string {
  return (result.content?.[0] as { text?: string })?.text ?? '';
}

type Msg = { id: string; text: string; fromMe: boolean; replyTo?: string };
type Thread = { root: Msg | { id: string; missing: true }; replies: Msg[] };

describe('teams_read_conversation', () => {
  it('reads a chat by id, oldest first, without system events', async () => {
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc, { conversation: GROUP });

    expect((data?.messages as Msg[]).map((m) => m.text)).toEqual([
      'first',
      'second',
      'third',
      'mine, after the horizon',
      'fourth',
    ]);
    expect(data?.conversation).toMatchObject({ id: GROUP, kind: 'group', title: 'Project Falcon', unread: true });
    expect(data?.unread_count).toBe(2);
    expect(data).not.toHaveProperty('threads');
    expect(data).not.toHaveProperty('truncated');
    expect(chatsvc.getMessages).toHaveBeenCalledWith(GROUP, expect.anything());
  });

  it('orders oldest first by arrival even when the service returns a page slightly out of order', async () => {
    // Live 2026-10-10: a quote-reply sent 6s after the message it quotes came
    // back AHEAD of it in the newest-first page, so a plain reverse() listed
    // the reply before its original.
    const outOfOrder = [GROUP_MSGS[5], GROUP_MSGS[3], GROUP_MSGS[4], GROUP_MSGS[2], GROUP_MSGS[1], GROUP_MSGS[0]];
    const chatsvc = chatsvcWith(GROUP_MSGS, { getMessages: vi.fn(async () => outOfOrder) });
    const { data } = await call(chatsvc, { conversation: GROUP });

    expect((data?.messages as Msg[]).map((m) => m.text)).toEqual([
      'first',
      'second',
      'third',
      'mine, after the horizon',
      'fourth',
    ]);
  });

  it('resolves a conversation by a unique name', async () => {
    const { data } = await call(chatsvcWith(), { conversation: 'project falcon' });

    expect(data?.conversation).toMatchObject({ id: GROUP, match_reason: 'title' });
  });

  it('never resolves a name by what was said in a chat (its last-message preview)', async () => {
    const convs = [...CONVS, rawConv('19:said0001@thread.v2', { topic: 'Lunch', at: T0, text: 'quarterly offsite agenda' })];
    const { result } = await call(fakeChatsvc(convs, { getMessages: messagesFrom(GROUP_MSGS) }), { conversation: 'offsite' });

    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/No chat or channel matches/);
  });

  it('returns an ambiguous name\'s candidates as data, not an error, reading nothing', async () => {
    const chatsvc = chatsvcWith();
    const { result, data } = await call(chatsvc, { conversation: 'falcon' });

    expect(result.isError).toBeFalsy();
    expect(data).toMatchObject({ ambiguous: true, query: 'falcon', hint: 'pass one id as conversation' });
    const candidates = data?.candidates as Record<string, unknown>[];
    expect(candidates.map((c) => c.id).sort()).toEqual([GROUP, OTHER_GROUP].sort());
    expect(candidates.find((c) => c.id === OTHER_GROUP)).toEqual({
      id: OTHER_GROUP,
      title: 'Falcon launch party',
      kind: 'group',
      lastActivity: new Date(T0).toISOString(),
      match_reason: 'title',
    });
    expect(data).not.toHaveProperty('messages');
    expect(chatsvc.getMessages).not.toHaveBeenCalled();
  });

  it('names a channel candidate\'s team', async () => {
    const convs = [...CONVS, rawConv('19:chan0002@thread.tacv2', { topic: 'Design reviews', at: T0, spaceId: TEAM })];
    const { data } = await call(fakeChatsvc(convs), { conversation: 'design' });

    const candidates = data?.candidates as Record<string, unknown>[];
    expect(candidates.find((c) => c.id === CHANNEL)).toMatchObject({ kind: 'channel', team: 'Platform Team' });
  });

  it('lists at most 10 candidates for a very ambiguous name, saying how many matched', async () => {
    const many = Array.from({ length: 14 }, (_, i) => rawConv(`19:dup${i}@thread.v2`, { topic: `Standup ${i}` }));
    const { data } = await call(fakeChatsvc(many), { conversation: 'standup' });

    const ids = (data?.candidates as { id: string }[]).map((c) => c.id);
    expect(ids).toHaveLength(10);
    expect(ids).toContain('19:dup9@thread.v2');
    expect(ids).not.toContain('19:dup10@thread.v2');
    expect(data?.match_count).toBe(14);
  });

  it('marks the candidates untrusted, since their titles are Teams text', async () => {
    const { result } = await call(chatsvcWith(), { conversation: 'falcon' });

    expect(errorText(result)).toMatch(/untrusted/i);
  });

  it('resolves an untitled group chat by a participant name, though every one has the same last sender', async () => {
    const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;
    const groups = ['19:untitled0001@thread.v2', '19:untitled0002@thread.v2', '19:untitled0003@thread.v2'];
    const convs = groups.map((id, i) => rawConv(id, { at: T0 + i * MIN, sender: 'King, Greg' }));
    const others = ['Alpha, Ann', 'Bravo, Bob', 'Charlie, Cy'];
    const byId = Object.fromEntries(
      groups.map((id, i) => [
        id,
        [
          { ...rawMsg(T0 - MIN, { sender: others[i], conv: id }), from: FROM(`8:orgid:other${i}`) },
          rawMsg(T0 + i * MIN, { sender: 'King, Greg', conv: id }),
        ],
      ]),
    );
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => messagesFrom(byId[id] ?? [])(id, opts));
    const chatsvc = fakeChatsvc(convs, { getMessages });

    const ambiguous = await call(chatsvc, { conversation: 'king' });
    expect(ambiguous.data?.ambiguous).toBe(true);
    expect((ambiguous.data?.candidates as { title: string }[]).map((c) => c.title)).toContain(
      'Group chat: King, Greg; Bravo, Bob',
    );

    const { data } = await call(chatsvc, { conversation: 'bravo' });
    expect(data?.conversation).toMatchObject({
      id: groups[1],
      title: 'Group chat: King, Greg; Bravo, Bob',
      participants: ['King, Greg', 'Bravo, Bob'],
    });
  });

  it('still throws on an ambiguous name for the write tools that resolve through resolveListed', async () => {
    const err = await resolveListed(chatsvcWith(), CONVS, ME, 'falcon', 'post to').catch((e: unknown) => e);

    expect(err).toBeInstanceOf(McpToolError);
    expect((err as McpToolError).message).toContain(GROUP);
    expect((err as McpToolError).message).toContain(OTHER_GROUP);
    expect((err as McpToolError).message).toContain('matched by title');
    expect((err as McpToolError).hint).toMatch(/id of the one to post to/);
  });

  it.each(['notes', 'Self', ' self ', '48:notes'])(
    'resolves the self-chat alias %j without reading any untitled group chat',
    async (conversation) => {
      const convs = [rawConv('19:untitled0001@thread.v2', { sender: 'King, Greg' }), rawConv('19:untitled0002@thread.v2')];
      const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => messagesFrom([])(id, opts));
      const chatsvc = fakeChatsvc(convs, { getMessages });

      const summary = await resolveListed(chatsvc, convs, ME, conversation, 'read');
      expect(summary).toMatchObject({ id: '48:notes', kind: 'self' });
      expect(getMessages).not.toHaveBeenCalled();
    },
  );

  it('errors when nothing matches, pointing at the listing tools', async () => {
    const { result } = await call(chatsvcWith(), { conversation: 'no such chat' });

    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/no chat or channel matches/i);
    expect(errorText(result)).toMatch(/teams_list_chats/);
  });

  it('returns only unread messages with unread_only, stopping at the read horizon', async () => {
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc, { conversation: GROUP, unread_only: true });

    expect((data?.messages as Msg[]).map((m) => m.text)).toEqual(['third', 'fourth']);
    expect(data?.unread_count).toBe(2);
    const opts = vi.mocked(chatsvc.getMessages).mock.calls[0][1] as GetMessagesOptions;
    expect(opts.stopAt?.(rawMsg(T0 + 2 * MIN))).toBe(true);
  });

  it('stops at since', async () => {
    const { data } = await call(chatsvcWith(), {
      conversation: GROUP,
      since: new Date(T0 + 3 * MIN).toISOString(),
    });

    expect((data?.messages as Msg[]).map((m) => m.text)).toEqual(['third', 'mine, after the horizon', 'fourth']);
  });

  it('rejects a since that is not a date-time', async () => {
    const { result } = await call(chatsvcWith(), { conversation: GROUP, since: 'yesterday' });

    expect(result.isError).toBe(true);
  });

  it('keeps the newest `limit` real messages (system events do not count) and says more exist', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, limit: 2 });

    expect((data?.messages as Msg[]).map((m) => m.text)).toEqual(['mine, after the horizon', 'fourth']);
    expect(data?.truncated).toMatchObject({ returned: 2 });
    expect((data?.truncated as { hint: string }).hint).toMatch(/limit/);
  });

  it('defaults to 50 messages', async () => {
    const many = Array.from({ length: 60 }, (_, i) => rawMsg(T0 + i * MIN));
    const { data } = await call(chatsvcWith(many), { conversation: GROUP });

    expect(data?.messages).toHaveLength(50);
    expect((data?.messages as Msg[]).at(-1)?.id).toBe(String(T0 + 59 * MIN));
    expect(data?.truncated).toBeDefined();
  });

  it('adds no truncated field when the history ends inside the limit', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, limit: 5 });

    expect(data?.messages).toHaveLength(5);
    expect(data).not.toHaveProperty('truncated');
  });

  it('groups a channel into threads, marking a root outside the window it cannot fetch as missing', async () => {
    const root1 = T0 - 10 * MIN; // outside the window
    const root2 = T0 + 1 * MIN;
    const msgs = [
      rawMsg(root1, { conv: CHANNEL, text: 'old root' }),
      rawMsg(T0 + 0.5 * MIN, { conv: CHANNEL, rootId: root1, text: 'late reply to old root' }),
      rawMsg(root2, { conv: CHANNEL, rootId: root2, text: 'new post' }),
      rawMsg(T0 + 3 * MIN, { conv: CHANNEL, rootId: root2, text: 'reply A' }),
      rawMsg(T0 + 4 * MIN, { conv: CHANNEL, rootId: root2, text: 'reply B' }),
    ];
    const { data } = await call(chatsvcWith(msgs), { conversation: 'design', since: new Date(T0).toISOString() });

    expect(data).not.toHaveProperty('messages');
    expect(data?.conversation).toMatchObject({ id: CHANNEL, kind: 'channel', team: 'Platform Team' });
    const threads = data?.threads as Thread[];
    expect(threads).toHaveLength(2);
    expect(threads[0].root).toEqual({ id: String(root1), missing: true });
    expect(threads[0].replies.map((r) => r.text)).toEqual(['late reply to old root']);
    expect((threads[1].root as Msg).text).toBe('new post');
    expect(threads[1].replies.map((r) => r.text)).toEqual(['reply A', 'reply B']);
    expect(data?.unread_count).toBe(2);
    // since stopped below the read horizon, so every unread message was seen.
    expect(data).not.toHaveProperty('unread_count_scope');
  });

  it('fills a root outside the window from the service, marked outside_window', async () => {
    const root = T0 - 10 * MIN;
    const msgs = [
      rawMsg(T0 + 3 * MIN, { conv: CHANNEL, rootId: root, text: 'late reply' }),
      rawMsg(T0 + 4 * MIN, { conv: CHANNEL, rootId: root, text: 'later reply' }),
    ];
    const stored = { ...rawMsg(root, { conv: CHANNEL, text: 'old root', sender: 'Example Poster' }), properties: { subject: 'Release plan' } };
    const getMessage = vi.fn(async (_c: string, id: string) => (id === String(root) ? stored : undefined));
    const { data } = await call(chatsvcWith(msgs, { getMessage }), { conversation: CHANNEL, since: new Date(T0).toISOString() });

    expect(getMessage).toHaveBeenCalledWith(CHANNEL, String(root));
    const threads = data?.threads as Thread[];
    expect(threads).toHaveLength(1);
    expect(threads[0].root).toMatchObject({
      id: String(root),
      text: 'old root',
      sender: 'Example Poster',
      subject: 'Release plan',
      time: new Date(root).toISOString(),
      outside_window: true,
    });
    expect(threads[0].root).not.toHaveProperty('missing');
    expect(threads[0].replies.map((r) => r.text)).toEqual(['late reply', 'later reply']);
    // The root is context, not part of the read: counts and cursor are the replies'.
    expect(data?.unread_count).toBe(2);
    expect(data?.next_cursor).toBe(String(T0 + 4 * MIN));
  });

  it('keeps a root it could not fetch as { id, missing: true } and still returns the read', async () => {
    const root = T0 - 10 * MIN;
    const msgs = [rawMsg(T0 + 3 * MIN, { conv: CHANNEL, rootId: root, text: 'late reply' })];
    const getMessage = vi.fn(async () => {
      throw new McpToolError('Teams chat service returned HTTP 403.', { status: 403 });
    });
    const { result, data } = await call(chatsvcWith(msgs, { getMessage }), { conversation: CHANNEL });

    expect(result.isError).toBeFalsy();
    expect((data?.threads as Thread[])[0].root).toEqual({ id: String(root), missing: true });
  });

  it('fetches no roots for a chat', async () => {
    const chatsvc = chatsvcWith();
    await call(chatsvc, { conversation: GROUP });

    expect(chatsvc.getMessage).not.toHaveBeenCalled();
  });

  it('fetches at most 25 roots in one call, across every conversation it reads', async () => {
    const CHANNEL2 = '19:chan0002@thread.tacv2';
    const repliesTo = (conv: string, base: number) =>
      Array.from({ length: 20 }, (_, i) => rawMsg(base + i * 1000, { conv, rootId: T0 - (i + 1) * MIN }));
    const byId: Record<string, ReturnType<typeof rawMsg>[]> = {
      [CHANNEL]: repliesTo(CHANNEL, T0 + MIN),
      [CHANNEL2]: repliesTo(CHANNEL2, T0 + 2 * MIN),
    };
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => messagesFrom(byId[id] ?? [])(id, opts));
    const getMessage = vi.fn(async (conv: string, id: string) => rawMsg(Number(id), { conv }));
    const convs = [...CONVS, rawConv(CHANNEL2, { topic: 'Ops', at: T0 + 4 * MIN, spaceId: TEAM })];
    const { data } = await call(fakeChatsvc(convs, { getMessages, getMessage }), { conversation: [CHANNEL, CHANNEL2] });

    expect(getMessage).toHaveBeenCalledTimes(25);
    const roots = (data?.conversations as { threads: Thread[] }[]).flatMap((c) => c.threads.map((t) => t.root));
    expect(roots).toHaveLength(40);
    expect(roots.filter((r) => 'outside_window' in r)).toHaveLength(25);
    expect(roots.filter((r) => 'missing' in r)).toHaveLength(15);
  });

  it('labels unread_count window-scoped when limit cut the read', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, limit: 2 });

    expect(data?.unread_count).toBe(1);
    expect(data?.unread_count_scope).toBe('window');
  });

  it('labels unread_count window-scoped when since cut the read above the read horizon', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, since: new Date(T0 + 3.5 * MIN).toISOString() });

    expect(data?.unread_count).toBe(1);
    expect(data?.unread_count_scope).toBe('window');
  });

  it('leaves unread_count unlabelled when the read was not cut', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP });

    expect(data).not.toHaveProperty('unread_count_scope');
  });

  it('says the result was truncated when the client stopped at its own page/message cap', async () => {
    const getMessages = vi.fn(async (_id: string, opts: GetMessagesOptions) => {
      opts.onCapped?.();
      return [rawMsg(T0 + 4 * MIN, { text: 'fourth' })];
    });
    const { data } = await call(chatsvcWith(GROUP_MSGS, { getMessages }), { conversation: GROUP });

    expect(data?.truncated).toMatchObject({ returned: 1 });
    expect(data?.unread_count_scope).toBe('window');
  });

  describe('settling the listing\'s unread hint', () => {
    const QUIET = '19:quiet0001@thread.v2';
    const quiet = rawConv(QUIET, { topic: 'Quiet', at: T0 + 5 * MIN, horizon: T0 + 2 * MIN, type: 'Event/Call' });
    const read = (messages: ReturnType<typeof rawMsg>[], args: Record<string, unknown> = {}) =>
      call(fakeChatsvc([quiet], { getMessages: messagesFrom(messages) }), { conversation: QUIET, ...args });

    it('is confirmed read when the fetch finds nothing unread', async () => {
      const { data } = await read([rawMsg(T0 + 1 * MIN), rawMsg(T0 + 5 * MIN, { type: 'Event/Call' })]);

      expect(data?.conversation).toMatchObject({ unread: false, unread_confirmed: true });
    });

    it('is confirmed unread when the fetch finds an unread message', async () => {
      const { data } = await read([rawMsg(T0 + 3 * MIN), rawMsg(T0 + 5 * MIN, { type: 'Event/Call' })]);

      expect(data?.conversation).toMatchObject({ unread: true, unread_confirmed: true });
    });

    it('stays an unconfirmed hint when a window cut the read before the horizon', async () => {
      const { data } = await read([rawMsg(T0 + 3 * MIN), rawMsg(T0 + 5 * MIN, { type: 'Event/Call' })], {
        since: new Date(T0 + 4 * MIN).toISOString(),
      });

      expect(data?.conversation).toMatchObject({ unread: true, unread_confirmed: false });
    });
  });

  describe('the 48:notes self-chat', () => {
    const NOTES = [rawMsg(T0 + 1 * MIN, { text: 'note to self', fromMe: true, conv: '48:notes' })];

    it.each(['48:notes', 'notes', 'Self'])('reads it as %j even when the listing does not carry it', async (conversation) => {
      const getMessages = messagesFrom(NOTES);
      const { data } = await call(fakeChatsvc(CONVS, { getMessages }), { conversation });

      expect(getMessages.mock.calls[0]?.[0]).toBe('48:notes');
      expect(data?.conversation).toMatchObject({ id: '48:notes', kind: 'self', title: 'Notes (self-chat)', unread: false });
      expect((data?.messages as Msg[]).map((m) => m.text)).toEqual(['note to self']);
    });

    it('reads it when the listing does carry it', async () => {
      const getMessages = messagesFrom(NOTES);
      const listed = [...CONVS, { id: '48:notes', properties: {}, threadProperties: {} }];
      const { data } = await call(fakeChatsvc(listed, { getMessages }), { conversation: 'notes' });

      expect(data?.conversation).toMatchObject({ id: '48:notes', kind: 'self' });
    });

    it('still refuses to resolve any other 48: id', async () => {
      const { result } = await call(fakeChatsvc(CONVS), { conversation: '48:calendar' });

      expect(result.isError).toBe(true);
    });
  });

  it('returns its result inside the untrusted-data envelope', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP });

    expect(data?.untrusted_content).toBe(true);
  });

  it('turns a token-capture failure into the open-a-Teams-tab hint', async () => {
    const chatsvc = fakeChatsvc([], {
      listConversations: vi.fn().mockRejectedValue(new Error('could not capture a Teams token from the browser (timed out).')),
    });
    const { result } = await call(chatsvc, { conversation: GROUP });

    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/click into any chat/i);
  });

  it('passes an already-actionable API error from the message read through unchanged', async () => {
    const chatsvc = chatsvcWith(GROUP_MSGS, {
      getMessages: vi.fn().mockRejectedValue(new McpToolError('Teams chat service is rate-limiting requests (429).', { status: 429 })),
    });
    const { result } = await call(chatsvc, { conversation: GROUP });

    expect(errorText(result)).toMatch(/rate-limiting/);
  });

  it('describes itself as any-conversation, not marking read unless asked, with attachments via teams_get_attachment', async () => {
    const harness = await createTestHarness((server) => registerConversationTools(server, { chatsvc: fakeChatsvc() }));
    const tools = (await harness.client.listTools()).tools;
    await harness.close();

    expect(tools.map((t) => t.name)).toEqual(['teams_read_conversation']);
    const tool = tools[0];
    expect(tool.description).toMatch(/any chat or channel/i);
    expect(tool.description).toMatch(/without opening/i);
    expect(tool.description).toMatch(/by default this does NOT mark anything read/i);
    expect(tool.description).toMatch(/mark_read/);
    expect(tool.description).toMatch(/teams_mark_read/);
    expect(tool.description).toMatch(/refused with format "summary"/);
    expect(tool.description).toMatch(/skipped/);
    expect(tool.description).toMatch(/teams_get_attachment/);
    // mark_read makes it a write — a non-destructive, idempotent one.
    expect(tool.annotations).toMatchObject({ readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: true });
    expect(Object.keys((tool.inputSchema as { properties?: object }).properties ?? {}).sort()).toEqual([
      'after_message_id',
      'confirmToken',
      'conversation',
      'format',
      'from',
      'limit',
      'mark_read',
      'since',
      'tz',
      'unread_only',
    ]);
    expect(tool.description).toMatch(/time_local/);
    expect(tool.description).toMatch(/TEAMS_TIME_ZONE/);
  });

  describe('time zones', () => {
    afterEach(() => vi.unstubAllEnvs());

    type Timed = { time: string; time_local?: string };

    it('keeps time in UTC and adds no time_local by default', async () => {
      vi.stubEnv('TEAMS_TIME_ZONE', '');
      const { data } = await call(chatsvcWith(), { conversation: GROUP });
      const [first] = data?.messages as Timed[];
      expect(first.time).toBe(new Date(T0 + MIN).toISOString());
      expect(first).not.toHaveProperty('time_local');
    });

    it('adds time_local in the tz asked for, next to the UTC time', async () => {
      const { data } = await call(chatsvcWith(), { conversation: GROUP, tz: 'America/New_York' });
      const [first] = data?.messages as Timed[];
      // T0 + 1 min is 2025-10-09T08:54:20Z; New York is on EDT (-04:00) then.
      expect(first.time).toBe('2025-10-09T08:54:20.000Z');
      expect(first.time_local).toBe('2025-10-09T04:54:20.000-04:00');
      expect(Object.keys(first).slice(0, 3)).toEqual(['id', 'time', 'time_local']);
    });

    it('adds time_local to channel thread posts too', async () => {
      const chatsvc = fakeChatsvc(CONVS, {
        getMessages: messagesFrom([rawMsg(T0 + 3 * MIN, { text: 'root', conv: CHANNEL })]),
      });
      const { data } = await call(chatsvc, { conversation: CHANNEL, tz: 'Asia/Kolkata' });
      const [thread] = data?.threads as { root: Timed }[];
      expect(thread.root.time_local).toBe('2025-10-09T14:26:20.000+05:30');
    });

    it('defaults tz to TEAMS_TIME_ZONE, and an explicit tz wins over it', async () => {
      vi.stubEnv('TEAMS_TIME_ZONE', 'Europe/London');
      const fromEnv = await call(chatsvcWith(), { conversation: GROUP });
      expect((fromEnv.data?.messages as Timed[])[0].time_local).toBe('2025-10-09T09:54:20.000+01:00');
      const explicit = await call(chatsvcWith(), { conversation: GROUP, tz: 'UTC' });
      expect((explicit.data?.messages as Timed[])[0].time_local).toBe('2025-10-09T08:54:20.000+00:00');
    });

    it('rejects a tz Intl does not know', async () => {
      const { result } = await call(chatsvcWith(), { conversation: GROUP, tz: 'Mars/Olympus_Mons' });
      expect(result.isError).toBe(true);
      expect(errorText(result)).toMatch(/IANA time zone/);
    });

    it('rejects a TEAMS_TIME_ZONE Intl does not know, naming the variable', async () => {
      vi.stubEnv('TEAMS_TIME_ZONE', 'Not/AZone');
      const { result } = await call(chatsvcWith(), { conversation: GROUP });
      expect(result.isError).toBe(true);
      expect(errorText(result)).toMatch(/TEAMS_TIME_ZONE/);
    });
  });
});

describe('teams_read_conversation after_message_id (cursor)', () => {
  const ids = (data: Record<string, unknown> | undefined) => (data?.messages as Msg[]).map((m) => m.id);
  const at = (min: number) => String(T0 + min * MIN);

  it('returns only messages newer than the cursor, excluding the cursor message itself', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, after_message_id: at(3) });

    expect(ids(data)).toEqual([at(3.5), at(4)]);
    expect(data?.next_cursor).toBe(at(4));
  });

  it('stops paging at the first message at or before the cursor', async () => {
    const chatsvc = chatsvcWith();
    await call(chatsvc, { conversation: GROUP, after_message_id: at(3) });

    const opts = vi.mocked(chatsvc.getMessages).mock.calls[0][1] as GetMessagesOptions;
    expect(opts.stopAt?.(rawMsg(T0 + 3 * MIN))).toBe(true);
    expect(opts.stopAt?.(rawMsg(T0 + 1 * MIN))).toBe(true);
    expect(opts.stopAt?.(rawMsg(T0 + 3.5 * MIN))).toBe(false);
  });

  it('returns nothing, and the passed cursor back, when nothing is new', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, after_message_id: at(4) });

    expect(data?.messages).toEqual([]);
    expect(data?.next_cursor).toBe(at(4));
    expect(data).not.toHaveProperty('truncated');
  });

  it('a cursor past only system events and my own sends returns no messages and keeps the cursor', async () => {
    const msgs = [...GROUP_MSGS, rawMsg(T0 + 5 * MIN, { type: 'Event/Call' })];
    const { data } = await call(chatsvcWith(msgs), { conversation: GROUP, after_message_id: at(4) });

    expect(data?.messages).toEqual([]);
    expect(data?.next_cursor).toBe(at(4));
  });

  it('returns next_cursor as the newest returned message id without a cursor', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP });

    expect(data?.next_cursor).toBe(at(4));
  });

  it('omits next_cursor when nothing was returned and no cursor was passed', async () => {
    const { data } = await call(chatsvcWith([]), { conversation: GROUP });

    expect(data).not.toHaveProperty('next_cursor');
  });

  it('feeds next_cursor back as after_message_id to read only what arrived since', async () => {
    const first = await call(chatsvcWith(GROUP_MSGS.slice(0, 4)), { conversation: GROUP });
    expect(first.data?.next_cursor).toBe(at(3));

    const second = await call(chatsvcWith(), { conversation: GROUP, after_message_id: first.data?.next_cursor });
    expect(ids(second.data)).toEqual([at(3.5), at(4)]);
  });

  it('with limit keeps the newest after the cursor, says it truncated, and moves the cursor to the newest', async () => {
    const many = Array.from({ length: 10 }, (_, i) => rawMsg(T0 + i * MIN));
    const { data } = await call(chatsvcWith(many), { conversation: GROUP, after_message_id: at(2), limit: 3 });

    expect(ids(data)).toEqual([at(7), at(8), at(9)]);
    expect(data?.truncated).toMatchObject({ returned: 3 });
    expect(data?.next_cursor).toBe(at(9));
  });

  it('is no truncation when the cursor, not the limit, ended the read', async () => {
    const many = Array.from({ length: 10 }, (_, i) => rawMsg(T0 + i * MIN));
    const { data } = await call(chatsvcWith(many), { conversation: GROUP, after_message_id: at(6), limit: 3 });

    expect(ids(data)).toEqual([at(7), at(8), at(9)]);
    expect(data).not.toHaveProperty('truncated');
  });

  it('with since, the later of the two boundaries wins', async () => {
    const sinceLater = await call(chatsvcWith(), {
      conversation: GROUP,
      after_message_id: at(1),
      since: new Date(T0 + 3.5 * MIN).toISOString(),
    });
    expect(ids(sinceLater.data)).toEqual([at(3.5), at(4)]);

    const cursorLater = await call(chatsvcWith(), {
      conversation: GROUP,
      after_message_id: at(3.5),
      since: new Date(T0 + 1 * MIN).toISOString(),
    });
    expect(ids(cursorLater.data)).toEqual([at(4)]);
  });

  it('with unread_only, a cursor below the read horizon changes nothing', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, unread_only: true, after_message_id: at(1) });

    expect((data?.messages as Msg[]).map((m) => m.text)).toEqual(['third', 'fourth']);
    expect(data?.unread_count).toBe(2);
    expect(data).not.toHaveProperty('unread_count_scope');
    expect(data?.next_cursor).toBe(at(4));
  });

  it('with unread_only, a cursor above the read horizon returns only the unread past it, window-scoped', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, unread_only: true, after_message_id: at(3) });

    expect((data?.messages as Msg[]).map((m) => m.text)).toEqual(['fourth']);
    expect(data?.unread_count).toBe(1);
    expect(data?.unread_count_scope).toBe('window');
  });

  it('with unread_only and nothing new past the cursor, returns nothing and the cursor', async () => {
    const { data } = await call(chatsvcWith(), { conversation: GROUP, unread_only: true, after_message_id: at(4) });

    expect(data?.messages).toEqual([]);
    expect(data?.next_cursor).toBe(at(4));
  });

  it('moves the cursor to the newest reply in a channel, wherever its thread sits', async () => {
    const root = T0 + 1 * MIN;
    const msgs = [
      rawMsg(root, { conv: CHANNEL, text: 'root' }),
      rawMsg(T0 + 5 * MIN, { conv: CHANNEL, rootId: root, text: 'new reply' }),
    ];
    const { data } = await call(chatsvcWith(msgs), { conversation: CHANNEL, after_message_id: at(2) });

    const threads = data?.threads as Thread[];
    expect(threads).toHaveLength(1);
    expect(threads[0].root).toEqual({ id: String(root), missing: true });
    expect(data?.next_cursor).toBe(at(5));
  });

  it('rejects an after_message_id that is not a message id', async () => {
    for (const after_message_id of ['yesterday', '', '12a', '-5']) {
      const { result } = await call(chatsvcWith(), { conversation: GROUP, after_message_id });
      expect(result.isError).toBe(true);
    }
  });

  it('stamps every result with fetched_at, the time of this live read', async () => {
    const before = Date.now();
    const { data } = await call(chatsvcWith(), { conversation: GROUP });

    const fetchedAt = Date.parse(String(data?.fetched_at));
    expect(fetchedAt).toBeGreaterThanOrEqual(before);
    expect(fetchedAt).toBeLessThanOrEqual(Date.now());
    expect(String(data?.fetched_at)).toMatch(/Z$/);
  });

  it('documents the cursor and that nothing is cached', async () => {
    const harness = await createTestHarness((server) => registerConversationTools(server, { chatsvc: chatsvcWith() }));
    try {
      const { tools } = await harness.client.listTools();
      const tool = tools.find((t) => t.name === 'teams_read_conversation')!;
      expect(tool.description).toMatch(/after_message_id/);
      expect(tool.description).toMatch(/next_cursor/);
      expect(tool.description).toMatch(/fetched_at/);
      expect(tool.description).toMatch(/no cache|not cached|never cached/i);
      expect(tool.inputSchema.properties).toHaveProperty('after_message_id');
    } finally {
      await harness.close();
    }
  });
});

describe('teams_read_conversation — mark_read', () => {
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
    const harness = await createTestHarness((server) => registerConversationTools(server, { chatsvc }));
    try {
      const first = await harness.callTool('teams_read_conversation', args);
      const one = first.isError ? undefined : (parseToolResult(first) as Record<string, unknown>);
      if (typeof one?.confirmToken !== 'string') return { first, preview: undefined, data: one };
      const second = await harness.callTool('teams_read_conversation', { ...args, confirmToken: one.confirmToken });
      return { first, preview: one, data: second.isError ? undefined : (parseToolResult(second) as Record<string, unknown>) };
    } finally {
      await harness.close();
    }
  }

  function withHorizonWrite(overrides: Record<string, unknown> = {}) {
    const setConsumptionHorizon = vi.fn(async () => undefined);
    return { chatsvc: chatsvcWith(GROUP_MSGS, { setConsumptionHorizon, ...overrides }), setConsumptionHorizon };
  }

  it('never writes the read marker by default, and asks for no confirmation', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { data } = await call(chatsvc, { conversation: GROUP });
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(data).not.toHaveProperty('marked_read');
    expect(data).not.toHaveProperty('confirmToken');
  });

  it('phase 1 previews the move and returns no messages; phase 2 reads and marks up to the newest returned', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { preview, data } = await callConfirmed(chatsvc, { conversation: GROUP, mark_read: true });

    expect(preview).not.toHaveProperty('messages');
    const shown = (preview?.preview as { conversations: { untrusted_content: boolean; data?: unknown } }).conversations;
    expect(JSON.stringify(shown)).toContain(String(T0 + 4 * MIN));
    expect(JSON.stringify(shown)).toContain('Project Falcon');

    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(GROUP, String(T0 + 4 * MIN));
    expect((data?.messages as Msg[]).map((m) => m.text)).toContain('fourth');
    expect(data?.marked_read).toBe(true);
  });

  it('leaves the marker alone when limit cut off unread messages older than those returned', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    // limit 2 returns +4 and +3.5 (mine); the unread "third" (+3) is past the horizon but was not returned.
    const { preview, data } = await callConfirmed(chatsvc, { conversation: GROUP, mark_read: true, limit: 2 });

    expect(preview).toBeUndefined();
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect((data?.messages as Msg[]).map((m) => m.text)).not.toContain('third');
    expect(data?.truncated).toBeDefined();
    expect(data?.marked_read).toEqual({ skipped: 'truncated', hint: expect.stringMatching(/limit/) });
  });

  it('with unread_only, a limit cut leaves the marker alone too', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { preview, data } = await callConfirmed(chatsvc, { conversation: GROUP, mark_read: true, unread_only: true, limit: 1 });

    expect(preview).toBeUndefined();
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(data?.marked_read).toMatchObject({ skipped: 'truncated' });
  });

  it('still marks when limit cut only messages already read', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    // limit 3 returns +4, +3.5 and +3 — every unread message; the cut falls at the horizon (+2).
    const { data } = await callConfirmed(chatsvc, { conversation: GROUP, mark_read: true, limit: 3 });

    expect(data?.truncated).toBeDefined();
    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(GROUP, String(T0 + 4 * MIN));
    expect(data?.marked_read).toBe(true);
  });

  it('leaves the marker alone when since cut off unread messages older than those returned', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { preview, data } = await callConfirmed(chatsvc, {
      conversation: GROUP,
      mark_read: true,
      since: new Date(T0 + 3.25 * MIN).toISOString(),
    });

    expect(preview).toBeUndefined();
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(data?.marked_read).toEqual({ skipped: 'since', hint: expect.stringMatching(/since/) });
  });

  it('still marks when since stops at or below the read horizon', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { data } = await callConfirmed(chatsvc, { conversation: GROUP, mark_read: true, since: new Date(T0 + 1.5 * MIN).toISOString() });

    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(GROUP, String(T0 + 4 * MIN));
    expect(data?.marked_read).toBe(true);
  });

  it('with an array, a truncated conversation is skipped while the others are marked under one confirmation', async () => {
    const msgs = {
      [GROUP]: GROUP_MSGS,
      [CHANNEL]: [rawMsg(T0 + 4 * MIN, { conv: CHANNEL, text: 'channel news' })],
    };
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => messagesFrom(msgs[id as keyof typeof msgs])(id, opts));
    const { preview, data } = await callConfirmed(chatsvcWith(GROUP_MSGS, { getMessages, setConsumptionHorizon }), {
      conversation: [GROUP, CHANNEL],
      mark_read: true,
      limit: 2,
    });

    const shown = JSON.stringify((preview?.preview as Record<string, unknown>).conversations);
    expect(shown).toContain(CHANNEL);
    expect(shown).not.toContain(GROUP);
    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(CHANNEL, String(T0 + 4 * MIN));
    const reads = data?.conversations as { conversation: { id: string }; marked_read?: unknown }[];
    expect(reads.map((r) => [r.conversation.id, r.marked_read])).toEqual([
      [GROUP, { skipped: 'truncated', hint: expect.any(String) }],
      [CHANNEL, true],
    ]);
  });

  it('marks nothing and asks nothing for a conversation with no messages', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite({ getMessages: messagesFrom([]) });
    const { preview, data } = await callConfirmed(chatsvc, { conversation: GROUP, mark_read: true });

    expect(preview).toBeUndefined();
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(data?.marked_read).toBe(false);
  });

  it('never lets an outside_window thread root drive the mark target', async () => {
    const root = T0 + 1 * MIN;
    const msgs = [
      rawMsg(T0 + 3 * MIN, { conv: CHANNEL, rootId: root, text: 'late reply' }),
      rawMsg(T0 + 4 * MIN, { conv: CHANNEL, rootId: root, text: 'later reply' }),
    ];
    // The root (+1) predates the window's replies; fetched by id, it is context, never the target.
    const stored = rawMsg(root, { conv: CHANNEL, text: 'old root' });
    const getMessage = vi.fn(async (_c: string, id: string) => (id === String(root) ? stored : undefined));
    const setConsumptionHorizon = vi.fn(async () => undefined);
    const { data } = await callConfirmed(chatsvcWith(msgs, { getMessage, setConsumptionHorizon }), {
      conversation: CHANNEL,
      mark_read: true,
    });

    expect(getMessage).toHaveBeenCalledWith(CHANNEL, String(root));
    expect(JSON.stringify(data?.threads)).toContain('outside_window');
    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(CHANNEL, String(T0 + 4 * MIN));
  });

  it('never moves the marker backwards: a read entirely before it is not marked', async () => {
    const setConsumptionHorizon = vi.fn(async () => undefined);
    // The only messages are at or before the horizon (+2 min) — every one is already read.
    const old = [rawMsg(T0 + 1 * MIN, { text: 'first' }), rawMsg(T0 + 2 * MIN, { text: 'second' })];
    const svc = chatsvcWith(old, { setConsumptionHorizon });
    const { preview, data } = await callConfirmed(svc, { conversation: GROUP, mark_read: true });
    expect(preview).toBeUndefined();
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(data?.marked_read).toBe(false);
  });

  it('reports a move Teams did not confirm as marked_read false with a warning', async () => {
    const { WriteOutcomeUnknownError } = await import('@chrischall/mcp-utils');
    const setConsumptionHorizon = vi.fn(async () => {
      throw new WriteOutcomeUnknownError('Teams', 'PUT', { cause: new Error('socket hang up') });
    });
    const { data } = await callConfirmed(chatsvcWith(GROUP_MSGS, { setConsumptionHorizon }), {
      conversation: GROUP,
      mark_read: true,
    });
    expect(data?.messages).toBeDefined();
    expect(data?.marked_read).toBe(false);
    expect(String(data?.mark_read_warning)).toMatch(/did not confirm/i);
  });

  it('with an array, one confirmation covers every conversation read, and each says marked_read', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { preview, data } = await callConfirmed(chatsvc, { conversation: [GROUP, CHANNEL], mark_read: true });

    expect(preview).not.toHaveProperty('conversations');
    const shown = JSON.stringify((preview?.preview as Record<string, unknown>).conversations);
    expect(shown).toContain(GROUP);
    expect(shown).toContain(CHANNEL);
    expect(setConsumptionHorizon.mock.calls.sort()).toEqual(
      [
        [GROUP, String(T0 + 4 * MIN)],
        [CHANNEL, String(T0 + 4 * MIN)],
      ].sort(),
    );
    const reads = data?.conversations as { conversation: { id: string }; marked_read?: boolean }[];
    expect(reads.map((r) => [r.conversation.id, r.marked_read])).toEqual([
      [GROUP, true],
      [CHANNEL, true],
    ]);
  });

  it('with an array, a conversation that could not be resolved is in errors and is not marked', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { data } = await callConfirmed(chatsvc, { conversation: [GROUP, 'no such chat'], mark_read: true });

    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(GROUP, String(T0 + 4 * MIN));
    expect((data?.errors as { conversation: string }[]).map((e) => e.conversation)).toEqual(['no such chat']);
  });

  it('with an array, an ambiguous name comes back as its candidates and nothing of it is marked', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { data } = await callConfirmed(chatsvc, { conversation: ['falcon', CHANNEL], mark_read: true });

    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(CHANNEL, String(T0 + 4 * MIN));
    const [ambiguous, channel] = data?.conversations as Record<string, unknown>[];
    expect(ambiguous).toMatchObject({ ambiguous: true, query: 'falcon' });
    expect(ambiguous).not.toHaveProperty('marked_read');
    expect(channel).toMatchObject({ conversation: { id: CHANNEL }, marked_read: true });
  });

  it('refuses mark_read with format summary, reading and marking nothing', async () => {
    for (const conversation of [GROUP, [GROUP, CHANNEL]]) {
      const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
      const { first } = await callConfirmed(chatsvc, { conversation, mark_read: true, format: 'summary' });

      expect(first.isError).toBe(true);
      expect(errorText(first)).toMatch(/mark_read/);
      expect(errorText(first)).toMatch(/summary/);
      expect(chatsvc.getMessages).not.toHaveBeenCalled();
      expect(setConsumptionHorizon).not.toHaveBeenCalled();
    }
  });

  it('with format compact, marks up to the newest message returned', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { data } = await callConfirmed(chatsvc, { conversation: GROUP, mark_read: true, format: 'compact' });

    expect((data?.messages as Msg[]).map((m) => m.text)).toContain('fourth');
    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(GROUP, String(T0 + 4 * MIN));
    expect(data?.marked_read).toBe(true);
  });

  it('with after_message_id and nothing new, marks nothing and asks nothing', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { preview, data } = await callConfirmed(chatsvc, {
      conversation: GROUP,
      mark_read: true,
      after_message_id: String(T0 + 4 * MIN),
    });

    expect(preview).toBeUndefined();
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(data?.marked_read).toBe(false);
    expect(data?.next_cursor).toBe(String(T0 + 4 * MIN));
  });

  it('an ambiguous name with mark_read returns the candidates and marks nothing', async () => {
    const { chatsvc, setConsumptionHorizon } = withHorizonWrite();
    const { preview, data } = await callConfirmed(chatsvc, { conversation: 'falcon', mark_read: true });

    expect(preview).toBeUndefined();
    expect(data?.ambiguous).toBe(true);
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
  });
});
