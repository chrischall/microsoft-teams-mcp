import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerChannelTools } from '../../src/tools/channels.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import { fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;

/*
 * Two teams. A team's own thread (the id its channels name as `spaceId`) is in
 * the listing, which is what names the team. The orphan channel's team thread
 * is not, so its team is unknown and it sorts after the named teams.
 */
const ALPHA = '19:alpha@thread.tacv2';
const ALPHA_DESIGN = '19:alpha-design@thread.tacv2';
const BRAVO = '19:bravo@thread.tacv2';
const BRAVO_OPS = '19:bravo-ops@thread.tacv2';
const ORPHAN = '19:orphan@thread.tacv2';
const CHAT = '19:chat0001@thread.v2';

const CONVS = [
  rawConv(BRAVO_OPS, { topic: 'Ops', spaceId: BRAVO, at: T0 + 5 * MIN, horizon: T0 }),
  rawConv(ALPHA_DESIGN, { topic: 'Design', spaceId: ALPHA, at: T0 + 4 * MIN }),
  rawConv(BRAVO, { topic: 'Bravo Team', spaceId: BRAVO, at: T0 + 3 * MIN }),
  rawConv(ALPHA, { topic: 'Alpha Team', spaceId: ALPHA, at: T0 + 2 * MIN, horizon: T0 }),
  rawConv(ORPHAN, { topic: 'Lonely', spaceId: '19:gone@thread.tacv2', at: T0 + 1 * MIN }),
  rawConv(CHAT, { topic: 'Not a channel', at: T0 + 9 * MIN, horizon: T0 }),
];

/** A real unread post past the horizon in each candidate channel. */
const post = (conv: string, at: number) => rawMsg(at, { conv, rootId: at });
const MESSAGES: Record<string, RawMessage[]> = {
  [BRAVO_OPS]: [post(BRAVO_OPS, T0), post(BRAVO_OPS, T0 + 5 * MIN)],
  [ALPHA]: [post(ALPHA, T0), post(ALPHA, T0 + 1 * MIN), post(ALPHA, T0 + 2 * MIN)],
};

/** A `getMessages` that serves each conversation's own messages; an `Error`-valued entry is thrown instead. */
function withMessages(convs = CONVS, byId: Record<string, RawMessage[] | Error> = MESSAGES): ChatServiceClient {
  const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => {
    const v = byId[id] ?? [];
    if (v instanceof Error) throw v;
    return messagesFrom(v)(id, opts);
  });
  return fakeChatsvc(convs, { getMessages });
}

async function call(chatsvc: ChatServiceClient, args: Record<string, unknown> = {}) {
  const harness = await createTestHarness((server) => registerChannelTools(server, { chatsvc }));
  try {
    const result = await harness.callTool('teams_list_teams_and_channels', args);
    return { result, data: result.isError ? undefined : (parseToolResult(result) as Record<string, unknown>) };
  } finally {
    await harness.close();
  }
}

type Row = { id: string; kind: string; title: string; team?: string; unread: boolean; unread_confirmed: boolean; unread_count?: number };

describe('teams_list_teams_and_channels', () => {
  it('lists every channel from the API sorted by team then title, a team\'s own thread as General, with no chats', async () => {
    const { data } = await call(fakeChatsvc(CONVS));
    const channels = data?.channels as Row[];

    expect(channels.map((c) => [c.team, c.title])).toEqual([
      ['Alpha Team', 'Design'],
      ['Alpha Team', 'General'],
      ['Bravo Team', 'General'],
      ['Bravo Team', 'Ops'],
      [undefined, 'Lonely'],
    ]);
    expect(channels.every((c) => c.kind === 'channel')).toBe(true);
  });

  it('marks each channel with an unconfirmed unread hint, reading no messages', async () => {
    const chatsvc = withMessages();
    const { data } = await call(chatsvc);
    const unread = (data?.channels as Row[]).filter((c) => c.unread).map((c) => c.id);

    expect(unread.sort()).toEqual([ALPHA, BRAVO_OPS].sort());
    expect((data?.channels as Row[]).every((c) => c.unread_confirmed === false)).toBe(true);
    expect(chatsvc.getMessages).not.toHaveBeenCalled();
  });

  it('keeps only channels a read confirms unread with unread_only, with their unread_count', async () => {
    const { data } = await call(withMessages(), { unread_only: true });
    const channels = data?.channels as Row[];

    expect(channels.map((c) => c.id)).toEqual([ALPHA, BRAVO_OPS]);
    expect(channels.map((c) => c.unread_count)).toEqual([2, 1]);
    expect(channels.every((c) => c.unread_confirmed)).toBe(true);
    expect(data?.skipped_unreachable).toBe(0);
  });

  it('drops a candidate channel whose posts past the horizon are only system events', async () => {
    const msgs = { ...MESSAGES, [ALPHA]: [post(ALPHA, T0), rawMsg(T0 + 2 * MIN, { conv: ALPHA, type: 'ThreadActivity/AddMember' })] };
    const { data } = await call(withMessages(CONVS, msgs), { unread_only: true });

    expect((data?.channels as Row[]).map((c) => c.id)).toEqual([BRAVO_OPS]);
  });

  it('skips a channel that answers 404 (gone, or no access) silently, counting it in skipped_unreachable (live fact)', async () => {
    const gone = new McpToolError('Teams chat service returned HTTP 404.', { status: 404, kind: 'http' });
    const { result, data } = await call(withMessages(CONVS, { ...MESSAGES, [ALPHA]: gone }), { unread_only: true });

    expect(result.isError).toBeFalsy();
    expect((data?.channels as Row[]).map((c) => c.id)).toEqual([BRAVO_OPS]);
    expect(data?.skipped_unreachable).toBe(1);
    expect(data).not.toHaveProperty('errors');
  });

  it('filters by channel title or team name with query', async () => {
    const byTeam = (await call(fakeChatsvc(CONVS), { query: 'bravo' })).data!;
    const byTitle = (await call(fakeChatsvc(CONVS), { query: 'DESIGN' })).data!;

    expect((byTeam.channels as Row[]).map((c) => c.id)).toEqual([BRAVO, BRAVO_OPS]);
    expect((byTitle.channels as Row[]).map((c) => c.id)).toEqual([ALPHA_DESIGN]);
  });

  it('honours limit and says what was left out', async () => {
    const { data } = await call(fakeChatsvc(CONVS), { limit: 2 });

    expect((data?.channels as Row[]).map((c) => c.title)).toEqual(['Design', 'General']);
    expect(data?.truncated).toMatchObject({ total: 5, returned: 2 });
  });

  describe('hidden channels', () => {
    // Unread, but hidden in Teams; and a visible unread channel with no interest field.
    const HIDDEN = '19:alpha-hidden@thread.tacv2';
    const WITH_HIDDEN = [
      ...CONVS,
      rawConv(HIDDEN, { topic: 'Archive', spaceId: ALPHA, at: T0 + 6 * MIN, horizon: T0, interest: 'NotInterested' }),
    ];
    const HIDDEN_MSGS = { ...MESSAGES, [HIDDEN]: [post(HIDDEN, T0 + 6 * MIN)] };

    it('lists a hidden channel marked hidden, after every visible one', async () => {
      const { data } = await call(fakeChatsvc(WITH_HIDDEN));
      const channels = data?.channels as (Row & { hidden?: true })[];

      expect(channels.at(-1)).toMatchObject({ id: HIDDEN, hidden: true });
      expect(channels.filter((c) => c.hidden).map((c) => c.id)).toEqual([HIDDEN]);
    });

    it('leaves a hidden channel out of unread_only by default', async () => {
      const chatsvc = withMessages(WITH_HIDDEN, HIDDEN_MSGS);
      const { data } = await call(chatsvc, { unread_only: true });

      expect((data?.channels as Row[]).map((c) => c.id)).toEqual([ALPHA, BRAVO_OPS]);
      expect(vi.mocked(chatsvc.getMessages).mock.calls.map((c) => c[0])).not.toContain(HIDDEN);
    });

    it('includes it with include_hidden', async () => {
      const { data } = await call(withMessages(WITH_HIDDEN, HIDDEN_MSGS), { unread_only: true, include_hidden: true });

      expect((data?.channels as Row[]).map((c) => c.id)).toEqual([ALPHA, BRAVO_OPS, HIDDEN]);
    });
  });

  it('returns its result inside the untrusted-data envelope', async () => {
    const { data } = await call(fakeChatsvc(CONVS));

    expect(data?.untrusted_content).toBe(true);
  });

  it('turns a token-capture failure into the open-a-Teams-tab hint', async () => {
    const chatsvc = fakeChatsvc([], {
      listConversations: vi.fn().mockRejectedValue(new Error('could not capture a Teams token from the browser (x).')),
    });

    const { result } = await call(chatsvc);

    expect(result.isError).toBe(true);
    expect((result.content?.[0] as { text?: string })?.text ?? '').toMatch(/click into any chat/i);
  });

  it('describes channel coverage, takes only optional filters, and is read-only', async () => {
    const harness = await createTestHarness((server) => registerChannelTools(server, { chatsvc: fakeChatsvc() }));
    const tools = (await harness.client.listTools()).tools;
    const tool = tools.find((t) => t.name === 'teams_list_teams_and_channels');
    const schema = tool?.inputSchema as { properties?: object; required?: string[] };

    expect(tools.map((t) => t.name)).toEqual(['teams_list_teams_and_channels']);
    expect(tool?.description).toMatch(/every channel/i);
    expect(tool?.description).toMatch(/teams_read_conversation/);
    expect(tool?.annotations?.readOnlyHint).toBe(true);
    expect(Object.keys(schema.properties ?? {}).sort()).toEqual(['include_hidden', 'limit', 'query', 'unread_only']);
    expect(tool?.description).toMatch(/hidden/i);
    expect(schema.required ?? []).toEqual([]);
    await harness.close();
  });
});
