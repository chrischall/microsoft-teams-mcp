import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerChatTools } from '../../src/tools/chat.js';
import type { ChatServiceClient } from '../../src/chatsvc.js';
import { fakeChatsvc, rawConv } from './fake-chatsvc.js';

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

async function call(chatsvc: ChatServiceClient, args: Record<string, unknown> = {}) {
  const harness = await createTestHarness((server) => registerChatTools(server, { chatsvc }));
  try {
    const result = await harness.callTool('teams_list_chats', args);
    return { result, data: result.isError ? undefined : (parseToolResult(result) as Record<string, unknown>) };
  } finally {
    await harness.close();
  }
}

type Row = { id: string; kind: string; title: string; unread: boolean };

describe('teams_list_chats', () => {
  it('lists every chat (group, meeting, 1:1) from the API, newest activity first, without channels', async () => {
    const { data } = await call(fakeChatsvc(CONVS));
    const chats = data?.chats as Row[];

    expect(chats.map((c) => c.id)).toEqual([GROUP, MEETING, DM]);
    expect(chats.map((c) => c.kind)).toEqual(['group', 'meeting', 'oneOnOne']);
    expect(chats.find((c) => c.id === DM)?.title).toBe('Pat Example');
    expect(chats.map((c) => c.unread)).toEqual([true, false, true]);
  });

  it('keeps only unread chats with unread_only', async () => {
    const { data } = await call(fakeChatsvc(CONVS), { unread_only: true });

    expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP, DM]);
  });

  it('lists the 48:notes self-chat when Teams lists it, as kind self', async () => {
    const { data } = await call(fakeChatsvc([...CONVS, { id: '48:notes', properties: {}, threadProperties: {} }]));
    const self = (data?.chats as Row[]).find((c) => c.id === '48:notes');

    expect(self).toMatchObject({ kind: 'self', title: 'Notes (self-chat)', unread: false });
  });

  it('does not call a meeting chat unread for a call recording after the last read message', async () => {
    const recorded = rawConv('19:meeting_rec0001@thread.v2', {
      topic: 'Recorded sync',
      at: T0 + 50 * MIN,
      horizon: T0 + 45 * MIN,
      lastIm: T0 + 44 * MIN,
      type: 'RichText/Media_CallRecording',
    });
    const { data } = await call(fakeChatsvc([...CONVS, recorded]), { unread_only: true });

    expect((data?.chats as Row[]).map((c) => c.id)).toEqual([GROUP, DM]);
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
