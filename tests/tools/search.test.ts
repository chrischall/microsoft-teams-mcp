import { describe, it, expect, vi, afterEach } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerSearchTools, matchesQuery } from '../../src/tools/search.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import { UNTRUSTED_CONTENT_NOTE, UNTRUSTED_DESCRIPTION_SUFFIX } from '../../src/tools/untrusted.js';
import { DAY, fakeChatsvc, pinClock, rawConv, rawMsg } from './fake-chatsvc.js';

const NOW = 1_760_000_000_000;
const MIN = 60_000;

const GROUP = '19:group0001@thread.v2';
const OTHER_GROUP = '19:group0002@thread.v2';
const CHANNEL = '19:chan0001@thread.tacv2';
const TEAM = '19:team0001@thread.tacv2';
const OLD = '19:group0003@thread.v2';

const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: NOW - 1 * MIN }),
  rawConv(OTHER_GROUP, { topic: 'Launch party', at: NOW - 2 * DAY }),
  rawConv(CHANNEL, { topic: 'Design', at: NOW - 3 * DAY, spaceId: TEAM }),
  rawConv(TEAM, { topic: 'Platform Team', at: NOW - 400 * DAY }),
  rawConv(OLD, { topic: 'Ancient history', at: NOW - 90 * DAY }),
];

/** Every conversation's messages, keyed by id. */
const MESSAGES: Record<string, RawMessage[]> = {
  [GROUP]: [
    rawMsg(NOW - 10 * MIN, { text: '<p>The <b>deploy</b> is done</p>', sender: 'Ada Example', conv: GROUP }),
    rawMsg(NOW - 5 * MIN, { text: 'lunch?', sender: 'Bob Example', conv: GROUP }),
    rawMsg(NOW - 4 * MIN, { text: 'Deploy again tomorrow', fromMe: true, conv: GROUP }),
    rawMsg(NOW - 3 * MIN, { text: 'deploy deleted', deleted: true, conv: GROUP }),
    rawMsg(NOW - 2 * MIN, { type: 'ThreadActivity/AddMember', text: '<addmember>deploy</addmember>', conv: GROUP }),
  ],
  [OTHER_GROUP]: [rawMsg(NOW - 2 * DAY, { text: 'party deploy plan', sender: 'Cy Example', conv: OTHER_GROUP })],
  [CHANNEL]: [
    rawMsg(NOW - 40 * DAY, { text: 'deploy, but too old', conv: CHANNEL }),
    rawMsg(NOW - 3 * DAY, { text: 'channel DEPLOY reply', sender: 'Ada Example', conv: CHANNEL, rootId: NOW - 4 * DAY }),
  ],
  [TEAM]: [],
  [OLD]: [rawMsg(NOW - 90 * DAY, { text: 'deploy from long ago', conv: OLD })],
};

/** A `getMessages` that pages each conversation's own messages, newest first, like the real client. */
function perConversation(byId: Record<string, RawMessage[]> = MESSAGES) {
  return vi.fn(async (id: string, opts: GetMessagesOptions) => {
    const newestFirst = [...(byId[id] ?? [])].sort((a, b) => Number(b.id) - Number(a.id));
    const out: RawMessage[] = [];
    for (const m of newestFirst) {
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

function chatsvcWith(overrides: Record<string, unknown> = {}): ChatServiceClient {
  return fakeChatsvc(CONVS, { getMessages: perConversation(), ...overrides });
}

async function call(chatsvc: ChatServiceClient, args: Record<string, unknown>) {
  const harness = await createTestHarness((server) => registerSearchTools(server, { chatsvc }));
  try {
    const result = await harness.callTool('teams_search_messages', args);
    return { result, data: result.isError ? undefined : (parseToolResult(result) as Record<string, unknown>) };
  } finally {
    await harness.close();
  }
}

function errorText(result: { content?: unknown[] }): string {
  return (result.content?.[0] as { text?: string })?.text ?? '';
}

type Hit = {
  conversation: { id: string; title: string; kind: string; team?: string };
  message: { id: string; time: string; sender: string; fromMe: boolean; text: string; replyTo?: string };
};

afterEach(() => vi.useRealTimers());

describe('matchesQuery', () => {
  it('matches a case-insensitive substring', () => {
    expect(matchesQuery('The Deploy is done', 'deploy')).toBe(true);
    expect(matchesQuery('nothing here', 'deploy')).toBe(false);
  });

  it('requires every whitespace-separated word, in any order', () => {
    expect(matchesQuery('party deploy plan', 'plan deploy')).toBe(true);
    expect(matchesQuery('party deploy plan', 'deploy budget')).toBe(false);
  });

  it('matches a quoted phrase as one term', () => {
    expect(matchesQuery('the deploy plan', '"deploy plan"')).toBe(true);
    expect(matchesQuery('plan the deploy', '"deploy plan"')).toBe(false);
  });

  it('matches nothing for a query of only whitespace or quotes', () => {
    expect(matchesQuery('anything', '  ""  ')).toBe(false);
  });
});

describe('teams_search_messages', () => {
  it('finds matching user messages across conversations, newest first, from the last 30 days', async () => {
    pinClock(NOW);
    const { data } = await call(chatsvcWith(), { query: 'deploy' });
    const hits = data?.hits as Hit[];

    expect(hits.map((h) => h.message.text)).toEqual([
      'Deploy again tomorrow',
      'The deploy is done',
      'party deploy plan',
      'channel DEPLOY reply',
    ]);
    expect(hits[0]).toMatchObject({
      conversation: { id: GROUP, title: 'Project Falcon', kind: 'group' },
      message: { fromMe: true, time: new Date(NOW - 4 * MIN).toISOString() },
    });
    expect(hits[3]).toMatchObject({
      conversation: { id: CHANNEL, kind: 'channel', team: 'Platform Team' },
      message: { sender: 'Ada Example', replyTo: String(NOW - 4 * DAY) },
    });
    expect(data?.total_hits).toBe(4);
    expect(data).not.toHaveProperty('truncated');
  });

  it('never returns deleted messages or system events, even when their content matches', async () => {
    pinClock(NOW);
    const { data } = await call(chatsvcWith(), { query: 'deploy', conversation: GROUP });
    expect((data?.hits as Hit[]).map((h) => h.message.text)).toEqual(['Deploy again tomorrow', 'The deploy is done']);
  });

  it('returns no attachments or reactions noise — just the message fields a hit needs', async () => {
    pinClock(NOW);
    const { data } = await call(chatsvcWith(), { query: 'lunch' });
    const [hit] = data?.hits as Hit[];
    expect(Object.keys(hit.message).sort()).toEqual(['edited', 'fromMe', 'id', 'sender', 'text', 'time']);
  });

  it('does not read a conversation whose newest activity is older than the window', async () => {
    pinClock(NOW);
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc, { query: 'deploy' });
    const read = vi.mocked(chatsvc.getMessages).mock.calls.map(([id]) => id);
    expect(read).not.toContain(OLD);
    expect(read).not.toContain(TEAM);
    expect(data?.scan).toMatchObject({
      method: 'client_scan',
      since: new Date(NOW - 30 * DAY).toISOString(),
      conversations_scanned: 3,
      conversations_outside_window: 2,
      conversations_not_scanned: 0,
    });
  });

  it('stops each read at the since cutoff, newest first', async () => {
    pinClock(NOW);
    const chatsvc = chatsvcWith();
    await call(chatsvc, { query: 'deploy' });
    const channelRead = vi.mocked(chatsvc.getMessages).mock.calls.find(([id]) => id === CHANNEL);
    const stopAt = channelRead?.[1].stopAt;
    expect(stopAt?.(rawMsg(NOW - 31 * DAY))).toBe(true);
    expect(stopAt?.(rawMsg(NOW - 29 * DAY))).toBe(false);
  });

  it('honours an explicit since, widening the window to older conversations', async () => {
    pinClock(NOW);
    const { data } = await call(chatsvcWith(), {
      query: 'deploy',
      since: new Date(NOW - 100 * DAY).toISOString(),
    });
    const texts = (data?.hits as Hit[]).map((h) => h.message.text);
    expect(texts).toContain('deploy, but too old');
    expect(texts).toContain('deploy from long ago');
  });

  it('honours an explicit since, narrowing the window', async () => {
    pinClock(NOW);
    const { data } = await call(chatsvcWith(), { query: 'deploy', since: new Date(NOW - 1 * DAY).toISOString() });
    expect((data?.hits as Hit[]).map((h) => h.message.text)).toEqual(['Deploy again tomorrow', 'The deploy is done']);
  });

  it('filters by a case-insensitive substring of the sender display name', async () => {
    pinClock(NOW);
    const { data } = await call(chatsvcWith(), { query: 'deploy', from: 'ada' });
    expect((data?.hits as Hit[]).map((h) => h.message.text)).toEqual(['The deploy is done', 'channel DEPLOY reply']);
  });

  it('narrows to one conversation by name, reading only that one', async () => {
    pinClock(NOW);
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc, { query: 'deploy', conversation: 'launch' });
    expect((data?.hits as Hit[]).map((h) => h.conversation.id)).toEqual([OTHER_GROUP]);
    expect(vi.mocked(chatsvc.getMessages).mock.calls.map(([id]) => id)).toEqual([OTHER_GROUP]);
    expect(data?.scan).toMatchObject({ conversations_scanned: 1 });
  });

  it('narrows to a named conversation even when its newest activity predates the window', async () => {
    pinClock(NOW);
    const { data } = await call(chatsvcWith(), { query: 'deploy', conversation: OLD, since: new Date(NOW - 100 * DAY).toISOString() });
    expect((data?.hits as Hit[]).map((h) => h.message.text)).toEqual(['deploy from long ago']);
  });

  it('fails clearly when the conversation name matches nothing', async () => {
    pinClock(NOW);
    const { result } = await call(chatsvcWith(), { query: 'deploy', conversation: 'no such chat' });
    expect(result.isError).toBe(true);
    expect(errorText(result)).toContain('No chat or channel matches');
  });

  it('keeps the newest `limit` hits and says when there were more', async () => {
    pinClock(NOW);
    const { data } = await call(chatsvcWith(), { query: 'deploy', limit: 2 });
    expect((data?.hits as Hit[]).map((h) => h.message.text)).toEqual(['Deploy again tomorrow', 'The deploy is done']);
    expect(data?.total_hits).toBe(4);
    expect(data?.truncated).toMatchObject({ returned: 2, total: 4 });
  });

  it('defaults limit to 25', async () => {
    pinClock(NOW);
    const many = Array.from({ length: 30 }, (_, i) => rawMsg(NOW - (i + 1) * MIN, { text: `deploy ${i}`, conv: GROUP }));
    const chatsvc = fakeChatsvc([CONVS[0]], { getMessages: perConversation({ [GROUP]: many }) });
    const { data } = await call(chatsvc, { query: 'deploy' });
    expect(data?.hits as Hit[]).toHaveLength(25);
    expect(data?.total_hits).toBe(30);
  });

  it('reports a conversation whose read hit the per-conversation cap as partially scanned', async () => {
    pinClock(NOW);
    const chatsvc = chatsvcWith({
      getMessages: vi.fn(async (_id: string, opts: GetMessagesOptions) => {
        opts.onCapped?.();
        return [];
      }),
    });
    const { data } = await call(chatsvc, { query: 'deploy', conversation: GROUP });
    expect(data?.scan).toMatchObject({ partially_scanned: [{ id: GROUP, title: 'Project Falcon' }] });
  });

  it('skips a conversation that answers 404/403 and counts it', async () => {
    pinClock(NOW);
    const ok = perConversation();
    const chatsvc = chatsvcWith({
      getMessages: vi.fn(async (id: string, opts: GetMessagesOptions) => {
        if (id === CHANNEL) throw new McpToolError('gone', { status: 404, kind: 'http' });
        return ok(id, opts);
      }),
    });
    const { data } = await call(chatsvc, { query: 'deploy' });
    expect((data?.hits as Hit[]).map((h) => h.conversation.id)).not.toContain(CHANNEL);
    expect(data?.scan).toMatchObject({ skipped_unreachable: 1 });
    expect(data).not.toHaveProperty('errors');
  });

  it('lists a conversation whose read failed in errors and returns the rest', async () => {
    pinClock(NOW);
    const ok = perConversation();
    const chatsvc = chatsvcWith({
      getMessages: vi.fn(async (id: string, opts: GetMessagesOptions) => {
        if (id === OTHER_GROUP) throw new McpToolError('boom', { status: 500, kind: 'http' });
        return ok(id, opts);
      }),
    });
    const { data } = await call(chatsvc, { query: 'deploy' });
    expect((data?.hits as Hit[]).length).toBe(3);
    expect(data?.errors).toEqual([expect.objectContaining({ conversation: expect.objectContaining({ id: OTHER_GROUP }) })]);
  });

  it('fails the call when every read failed', async () => {
    pinClock(NOW);
    const chatsvc = chatsvcWith({
      getMessages: vi.fn(async () => {
        throw new McpToolError('rate limited', { status: 429, kind: 'http' });
      }),
    });
    const { result } = await call(chatsvc, { query: 'deploy' });
    expect(result.isError).toBe(true);
  });

  it('scans the most recently active conversations first and reports what the cap left out', async () => {
    pinClock(NOW);
    const convs = Array.from({ length: 120 }, (_, i) =>
      rawConv(`19:bulk${String(i).padStart(4, '0')}@thread.v2`, { topic: `Bulk ${i}`, at: NOW - (i + 1) * MIN }),
    );
    const chatsvc = fakeChatsvc(convs, { getMessages: perConversation({}) });
    const { data } = await call(chatsvc, { query: 'deploy' });
    const read = vi.mocked(chatsvc.getMessages).mock.calls.map(([id]) => id);
    expect(read).toHaveLength(100);
    expect(read).toContain('19:bulk0000@thread.v2');
    expect(read).not.toContain('19:bulk0119@thread.v2');
    expect(data?.scan).toMatchObject({ conversations_scanned: 100, conversations_not_scanned: 20 });
    expect((data?.scan as { hint?: string }).hint).toMatch(/conversation/);
  });

  it('scans more conversations when max_conversations asks for them', async () => {
    pinClock(NOW);
    const convs = Array.from({ length: 120 }, (_, i) =>
      rawConv(`19:bulk${String(i).padStart(4, '0')}@thread.v2`, { topic: `Bulk ${i}`, at: NOW - (i + 1) * MIN }),
    );
    const chatsvc = fakeChatsvc(convs, { getMessages: perConversation({}) });
    const { data } = await call(chatsvc, { query: 'deploy', max_conversations: 120 });
    expect(vi.mocked(chatsvc.getMessages).mock.calls).toHaveLength(120);
    expect(data?.scan).toMatchObject({ conversations_scanned: 120, conversations_not_scanned: 0 });
    expect(data?.scan).not.toHaveProperty('hint');
  });

  it('names max_conversations in the hint when the cap left conversations out', async () => {
    pinClock(NOW);
    const convs = Array.from({ length: 3 }, (_, i) => rawConv(`19:bulk${i}@thread.v2`, { topic: `Bulk ${i}`, at: NOW - (i + 1) * MIN }));
    const chatsvc = fakeChatsvc(convs, { getMessages: perConversation({}) });
    const { data } = await call(chatsvc, { query: 'deploy', max_conversations: 2 });
    expect((data?.scan as { hint: string }).hint).toContain('max_conversations');
  });

  it('leaves out channels hidden in Teams unless one is named', async () => {
    pinClock(NOW);
    const hidden = rawConv('19:hidden0001@thread.tacv2', { topic: 'Muted', at: NOW - MIN, interest: 'NotInterested', spaceId: TEAM });
    const chatsvc = fakeChatsvc([...CONVS, hidden], { getMessages: perConversation() });
    await call(chatsvc, { query: 'deploy' });
    expect(vi.mocked(chatsvc.getMessages).mock.calls.map(([id]) => id)).not.toContain(hidden.id);
  });

  it('wraps the result in the untrusted-content envelope and warns in the description', async () => {
    pinClock(NOW);
    const { data } = await call(chatsvcWith(), { query: 'deploy' });
    expect(data?.untrusted_content).toBe(true);
    expect(data?.note).toBe(UNTRUSTED_CONTENT_NOTE);
    const harness = await createTestHarness((server) => registerSearchTools(server, { chatsvc: chatsvcWith() }));
    try {
      const tool = (await harness.listTools()).find((t) => t.name === 'teams_search_messages');
      expect(tool?.description).toContain(UNTRUSTED_DESCRIPTION_SUFFIX);
    } finally {
      await harness.close();
    }
  });

  it('rejects a since that is not a date-time', async () => {
    const { result } = await call(chatsvcWith(), { query: 'deploy', since: 'last tuesday' });
    expect(result.isError).toBe(true);
  });
});
