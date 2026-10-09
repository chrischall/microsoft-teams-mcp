import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { registerChatTools } from '../../src/tools/chat.js';
import type { TeamsClient } from '../../src/client.js';

function fakeClient(overrides: Partial<TeamsClient> = {}): TeamsClient {
  return {
    listChats: vi.fn().mockResolvedValue([]),
    getOpenChatMessages: vi.fn().mockResolvedValue([]),
    getSelectedConversations: vi.fn().mockResolvedValue([]),
    ...overrides,
  } as unknown as TeamsClient;
}

describe('teams_list_chats', () => {
  it('says the Chat view may not be open when nothing is rendered', async () => {
    const client = fakeClient({ listChats: vi.fn().mockResolvedValue([]) });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_list_chats')) as Record<string, unknown>;

    expect(data.rows).toEqual([]);
    expect(String(data.empty_hint)).toMatch(/Chat view/);
    await harness.close();
  });

  it('adds no empty_hint when rows were read', async () => {
    const client = fakeClient({ listChats: vi.fn().mockResolvedValue([{ title: 'Standup' }]) });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_list_chats')) as Record<string, unknown>;

    expect(data).not.toHaveProperty('empty_hint');
    await harness.close();
  });

  it('returns the chats from the client', async () => {
    const client = fakeClient({
      listChats: vi.fn().mockResolvedValue([
        { title: 'Standup', preview: 'hi', time: '9:00 AM', conversationKey: 'k', itemType: 'chat' },
      ]),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const result = await harness.callTool('teams_list_chats');
    const data = (parseToolResult(result) as { rows: unknown }).rows;

    expect(data).toEqual([
      { title: 'Standup', preview: 'hi', time: '9:00 AM', conversationKey: 'k', itemType: 'chat' },
    ]);
    await harness.close();
  });

  it('surfaces a no-tab failure with an actionable hint', async () => {
    const client = fakeClient({
      listChats: vi.fn().mockRejectedValue(new Error('could not reach a signed-in Teams tab on any of […]')),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const result = await harness.callTool('teams_list_chats');

    expect(result.isError).toBe(true);
    const text = (result.content?.[0] as { text?: string })?.text ?? '';
    expect(text).toMatch(/signed-in Teams tab/i);
    expect(text).toMatch(/teams.cloud.microsoft/i);
    await harness.close();
  });
});

describe('teams_get_open_chat_messages', () => {
  it('returns messages from the client', async () => {
    const client = fakeClient({
      getOpenChatMessages: vi.fn().mockResolvedValue([
        { sender: 'Alice', time: '2026-09-21T10:00:00.000Z', messageId: 'timestamp-1', text: 'hi' },
      ]),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const result = await harness.callTool('teams_get_open_chat_messages');
    const data = (parseToolResult(result) as { messages: unknown }).messages;

    expect(data).toEqual([
      { sender: 'Alice', time: '2026-09-21T10:00:00.000Z', messageId: 'timestamp-1', text: 'hi' },
    ]);
    await harness.close();
  });

  it('says which conversation the messages came from', async () => {
    const client = fakeClient({
      getOpenChatMessages: vi.fn().mockResolvedValue([{ sender: 'Bob', text: 'hi' }]),
      getSelectedConversations: vi.fn().mockResolvedValue([{ title: 'Bob Smith', conversationKey: '19:abc@thread.v2', itemType: 'chat' }]),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_get_open_chat_messages')) as Record<string, unknown>;

    expect(data.conversation).toEqual({ title: 'Bob Smith', conversationKey: '19:abc@thread.v2', itemType: 'chat' });
    expect(String(data.conversation_check)).toMatch(/Bob Smith/);
    expect(String(data.conversation_check)).toMatch(/confirm/i);
    await harness.close();
  });

  it('warns when it cannot identify the open conversation', async () => {
    const client = fakeClient({
      getOpenChatMessages: vi.fn().mockResolvedValue([{ sender: 'Bob', text: 'hi' }]),
      getSelectedConversations: vi.fn().mockResolvedValue([]),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_get_open_chat_messages')) as Record<string, unknown>;

    expect(data.conversation).toBeNull();
    expect(String(data.conversation_check)).toMatch(/could not identify/i);
    expect(data.messages).toEqual([{ sender: 'Bob', text: 'hi' }]);
    await harness.close();
  });

  it('does not report a selected channel row as the open chat', async () => {
    const client = fakeClient({
      getOpenChatMessages: vi.fn().mockResolvedValue([{ sender: 'Bob', text: 'hi' }]),
      getSelectedConversations: vi
        .fn()
        .mockResolvedValue([{ title: 'General', conversationKey: '19:def@thread.tacv2', itemType: 'channel' }]),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_get_open_chat_messages')) as Record<string, unknown>;

    expect(data.conversation).toBeNull();
    expect(String(data.conversation_check)).toMatch(/could not identify/i);
    expect(String(data.conversation_check)).not.toMatch(/General/);
    await harness.close();
  });

  it('picks the chat row when a channel row is also marked selected', async () => {
    const client = fakeClient({
      getSelectedConversations: vi.fn().mockResolvedValue([
        { title: 'General', conversationKey: '19:def@thread.tacv2', itemType: 'channel' },
        { title: 'Bob Smith', conversationKey: '19:abc@thread.v2', itemType: 'chat' },
      ]),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_get_open_chat_messages')) as Record<string, unknown>;

    expect(data.conversation).toEqual({ title: 'Bob Smith', conversationKey: '19:abc@thread.v2', itemType: 'chat' });
    expect(String(data.conversation_check)).toMatch(/Bob Smith/);
    await harness.close();
  });

  it('cannot identify the chat when several chat rows are marked selected', async () => {
    const client = fakeClient({
      getSelectedConversations: vi.fn().mockResolvedValue([
        { title: 'Alice', conversationKey: '19:a@thread.v2', itemType: 'chat' },
        { title: 'Bob Smith', conversationKey: '19:abc@thread.v2', itemType: 'chat' },
      ]),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_get_open_chat_messages')) as Record<string, unknown>;

    expect(data.conversation).toBeNull();
    expect(String(data.conversation_check)).toMatch(/could not identify/i);
    await harness.close();
  });

  it('cannot identify the chat when the selected row carries no itemType', async () => {
    const client = fakeClient({
      getSelectedConversations: vi.fn().mockResolvedValue([{ title: 'Bob Smith', conversationKey: '19:abc@thread.v2' }]),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_get_open_chat_messages')) as Record<string, unknown>;

    expect(data.conversation).toBeNull();
    expect(String(data.conversation_check)).toMatch(/could not identify/i);
    await harness.close();
  });

  it('still returns the messages when reading the conversation identity fails', async () => {
    const client = fakeClient({
      getOpenChatMessages: vi.fn().mockResolvedValue([{ sender: 'Bob', text: 'hi' }]),
      getSelectedConversations: vi.fn().mockRejectedValue(new Error('read_dom_list name not in declared set: openConversation')),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const result = await harness.callTool('teams_get_open_chat_messages');
    const data = parseToolResult(result) as Record<string, unknown>;

    expect(result.isError).toBeFalsy();
    expect(data.conversation).toBeNull();
    expect(data.messages).toEqual([{ sender: 'Bob', text: 'hi' }]);
    await harness.close();
  });

  it('tells the model to confirm the conversation and warns about multiple Teams tabs', async () => {
    const harness = await createTestHarness((server) => registerChatTools(server, fakeClient()));
    const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_get_open_chat_messages');

    expect(tool?.description).toMatch(/confirm/i);
    expect(tool?.description).toMatch(/more than one .*tab/i);
    await harness.close();
  });

  it('surfaces a not-yet-approved bridge failure with an actionable hint', async () => {
    const client = fakeClient({
      getOpenChatMessages: vi
        .fn()
        .mockRejectedValue(new Error('read_dom_list name not in declared set: chatMessages')),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const result = await harness.callTool('teams_get_open_chat_messages');

    expect(result.isError).toBe(true);
    const text = (result.content?.[0] as { text?: string })?.text ?? '';
    expect(text).toMatch(/ContextMint Bridge/i);
    await harness.close();
  });

  it('takes only optional trimming arguments', async () => {
    const client = fakeClient();
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_get_open_chat_messages');
    const schema = tool?.inputSchema as { properties?: Record<string, unknown>; required?: string[] };

    expect(Object.keys(schema.properties ?? {}).sort()).toEqual(['limit', 'since']);
    expect(schema.required ?? []).toEqual([]);
    await harness.close();
  });
});

const messages = (n: number) =>
  Array.from({ length: n }, (_, i) => ({
    sender: 'Alice',
    time: new Date(Date.UTC(2026, 8, 1, 0, i)).toISOString(),
    messageId: `m${i}`,
    text: `msg ${i}`,
  }));

describe('teams_get_open_chat_messages trimming', () => {
  it('returns the newest 50 messages by default and says how many were left out', async () => {
    const client = fakeClient({ getOpenChatMessages: vi.fn().mockResolvedValue(messages(60)) });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_get_open_chat_messages')) as Record<string, unknown>;
    const rows = data.messages as { messageId: string }[];

    expect(rows).toHaveLength(50);
    expect(rows[0].messageId).toBe('m10');
    expect(rows[49].messageId).toBe('m59');
    expect(data.truncated).toMatchObject({ total: 60, returned: 50 });
    await harness.close();
  });

  it('honours an explicit limit', async () => {
    const client = fakeClient({ getOpenChatMessages: vi.fn().mockResolvedValue(messages(10)) });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(
      await harness.callTool('teams_get_open_chat_messages', { limit: 3 }),
    ) as Record<string, unknown>;

    expect((data.messages as { messageId: string }[]).map((m) => m.messageId)).toEqual(['m7', 'm8', 'm9']);
    await harness.close();
  });

  it('adds no truncated field when everything fits', async () => {
    const client = fakeClient({ getOpenChatMessages: vi.fn().mockResolvedValue(messages(3)) });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_get_open_chat_messages')) as Record<string, unknown>;

    expect(data.messages).toHaveLength(3);
    expect(data).not.toHaveProperty('truncated');
    await harness.close();
  });

  it('drops messages older than since, keeping ones whose time cannot be read', async () => {
    const rows = [...messages(5), { sender: 'Bob', text: 'no time' }];
    const client = fakeClient({ getOpenChatMessages: vi.fn().mockResolvedValue(rows) });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(
      await harness.callTool('teams_get_open_chat_messages', { since: '2026-09-01T00:03:00Z' }),
    ) as Record<string, unknown>;

    expect((data.messages as { text: string }[]).map((m) => m.text)).toEqual(['msg 3', 'msg 4', 'no time']);
    await harness.close();
  });

  it('rejects a since that is not a date-time', async () => {
    const harness = await createTestHarness((server) => registerChatTools(server, fakeClient()));

    const result = await harness.callTool('teams_get_open_chat_messages', { since: 'yesterday' });

    expect(result.isError).toBe(true);
    await harness.close();
  });

  it('rejects a limit above the 200 rows the tab can return', async () => {
    const harness = await createTestHarness((server) => registerChatTools(server, fakeClient()));

    const result = await harness.callTool('teams_get_open_chat_messages', { limit: 201 });

    expect(result.isError).toBe(true);
    await harness.close();
  });
});

describe('teams_list_chats trimming', () => {
  it('returns the first (most recent) 50 sidebar rows by default, and honours limit', async () => {
    const rows = Array.from({ length: 60 }, (_, i) => ({ title: `chat ${i}` }));
    const client = fakeClient({ listChats: vi.fn().mockResolvedValue(rows) });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const all = parseToolResult(await harness.callTool('teams_list_chats')) as Record<string, unknown>;
    const two = parseToolResult(await harness.callTool('teams_list_chats', { limit: 2 })) as Record<string, unknown>;

    expect((all.rows as unknown[]).length).toBe(50);
    expect((all.rows as { title: string }[])[0].title).toBe('chat 0');
    expect(all.truncated).toMatchObject({ total: 60, returned: 50 });
    expect(two.rows).toEqual([{ title: 'chat 0' }, { title: 'chat 1' }]);
    await harness.close();
  });
});

describe('tool annotations', () => {
  it('both chat tools are read-only', async () => {
    const client = fakeClient();
    const harness = await createTestHarness((server) => registerChatTools(server, client));
    const tools = (await harness.client.listTools()).tools;

    for (const name of ['teams_list_chats', 'teams_get_open_chat_messages']) {
      expect(tools.find((t) => t.name === name)?.annotations?.readOnlyHint, name).toBe(true);
    }
    await harness.close();
  });
});
