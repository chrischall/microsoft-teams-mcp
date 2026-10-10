import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerConversationTools } from '../../src/tools/conversation.js';
import type { ChatServiceClient, GetMessagesOptions } from '../../src/chatsvc.js';
import { fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

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

  it('resolves a conversation by a unique name', async () => {
    const { data } = await call(chatsvcWith(), { conversation: 'project falcon' });

    expect((data?.conversation as { id: string }).id).toBe(GROUP);
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

  describe('a conversation the listing could only call unread_uncertain', () => {
    const QUIET = '19:quiet0001@thread.v2';
    const base = rawConv(QUIET, { topic: 'Quiet', at: T0 + 5 * MIN, horizon: T0 + 2 * MIN });
    const quiet = { ...base, lastMessage: { ...base.lastMessage!, messagetype: 'Event/Call' } };
    const read = (messages: ReturnType<typeof rawMsg>[], args: Record<string, unknown> = {}) =>
      call(fakeChatsvc([quiet], { getMessages: messagesFrom(messages) }), { conversation: QUIET, ...args });

    it('is confirmed read when the fetch finds nothing unread', async () => {
      const { data } = await read([rawMsg(T0 + 1 * MIN), rawMsg(T0 + 5 * MIN, { type: 'Event/Call' })]);

      expect(data?.conversation).toMatchObject({ unread: false });
      expect(data?.conversation).not.toHaveProperty('unread_uncertain');
    });

    it('is confirmed unread when the fetch finds an unread message', async () => {
      const { data } = await read([rawMsg(T0 + 3 * MIN), rawMsg(T0 + 5 * MIN, { type: 'Event/Call' })]);

      expect(data?.conversation).toMatchObject({ unread: true });
      expect(data?.conversation).not.toHaveProperty('unread_uncertain');
    });

    it('stays uncertain when a window cut the read before the horizon', async () => {
      const { data } = await read([rawMsg(T0 + 3 * MIN), rawMsg(T0 + 5 * MIN, { type: 'Event/Call' })], {
        since: new Date(T0 + 4 * MIN).toISOString(),
      });

      expect(data?.conversation).toMatchObject({ unread: true, unread_uncertain: true });
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
