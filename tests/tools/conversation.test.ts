import { describe, it, expect, vi } from 'vitest';
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

  it('refuses an ambiguous name, naming the candidates and asking for the id', async () => {
    const chatsvc = chatsvcWith();
    const { result } = await call(chatsvc, { conversation: 'falcon' });

    expect(result.isError).toBe(true);
    const text = errorText(result);
    expect(text).toContain(GROUP);
    expect(text).toContain(OTHER_GROUP);
    expect(text).toContain('Falcon launch party');
    expect(text).toMatch(/\bid\b/);
    expect(text).toContain('matched by title');
    expect(chatsvc.getMessages).not.toHaveBeenCalled();
  });

  it('names at most 10 candidates for a very ambiguous name', async () => {
    const many = Array.from({ length: 14 }, (_, i) => rawConv(`19:dup${i}@thread.v2`, { topic: `Standup ${i}` }));
    const { result } = await call(fakeChatsvc(many), { conversation: 'standup' });

    const text = errorText(result);
    expect(text).toContain('19:dup9@thread.v2');
    expect(text).not.toContain('19:dup10@thread.v2');
    expect(text).toMatch(/14/);
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
    expect(ambiguous.result.isError).toBe(true);
    expect(errorText(ambiguous.result)).toContain('Group chat: King, Greg; Bravo, Bob');

    const { data } = await call(chatsvc, { conversation: 'bravo' });
    expect(data?.conversation).toMatchObject({
      id: groups[1],
      title: 'Group chat: King, Greg; Bravo, Bob',
      participants: ['King, Greg', 'Bravo, Bob'],
    });
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

  it('groups a channel into threads, marking a root outside the window as missing', async () => {
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

  it('describes itself as read-only, any-conversation, never marking read, with attachments via teams_get_attachment', async () => {
    const harness = await createTestHarness((server) => registerConversationTools(server, { chatsvc: fakeChatsvc() }));
    const tools = (await harness.client.listTools()).tools;
    await harness.close();

    expect(tools.map((t) => t.name)).toEqual(['teams_read_conversation']);
    const tool = tools[0];
    expect(tool.description).toMatch(/any chat or channel/i);
    expect(tool.description).toMatch(/without opening/i);
    expect(tool.description).toMatch(/does not mark/i);
    expect(tool.description).toMatch(/teams_get_attachment/);
    expect(tool.annotations?.readOnlyHint).toBe(true);
    expect(tool.annotations?.openWorldHint).toBe(true);
    expect(Object.keys((tool.inputSchema as { properties?: object }).properties ?? {}).sort()).toEqual([
      'conversation',
      'limit',
      'since',
      'unread_only',
    ]);
  });
});
