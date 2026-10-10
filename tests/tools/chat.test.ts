import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerChatTools } from '../../src/tools/chat.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import { fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;

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
      'unread_only',
    ]);
    await harness.close();
  });

  it('registers only teams_list_chats (the open-chat DOM reader is gone)', async () => {
    const harness = await createTestHarness((server) => registerChatTools(server, { chatsvc: fakeChatsvc() }));

    expect((await harness.client.listTools()).tools.map((t) => t.name)).toEqual(['teams_list_chats']);
    await harness.close();
  });
});
