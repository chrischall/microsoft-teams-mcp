import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { registerChatTools } from '../../src/tools/chat.js';
import type { TeamsClient } from '../../src/client.js';

function fakeClient(overrides: Partial<TeamsClient> = {}): TeamsClient {
  return {
    listChats: vi.fn().mockResolvedValue([]),
    getOpenChatMessages: vi.fn().mockResolvedValue([]),
    getOpenConversation: vi.fn().mockResolvedValue(null),
    ...overrides,
  } as unknown as TeamsClient;
}

describe('teams_list_chats', () => {
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
      getOpenConversation: vi
        .fn()
        .mockResolvedValue({ title: 'Bob Smith', conversationKey: '19:abc@thread.v2', itemType: 'chat' }),
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
      getOpenConversation: vi.fn().mockResolvedValue(null),
    });
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const data = parseToolResult(await harness.callTool('teams_get_open_chat_messages')) as Record<string, unknown>;

    expect(data.conversation).toBeNull();
    expect(String(data.conversation_check)).toMatch(/could not identify/i);
    expect(data.messages).toEqual([{ sender: 'Bob', text: 'hi' }]);
    await harness.close();
  });

  it('still returns the messages when reading the conversation identity fails', async () => {
    const client = fakeClient({
      getOpenChatMessages: vi.fn().mockResolvedValue([{ sender: 'Bob', text: 'hi' }]),
      getOpenConversation: vi.fn().mockRejectedValue(new Error('read_dom_list name not in declared set: openConversation')),
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
    const tool = (await harness.listTools()).find((t) => t.name === 'teams_get_open_chat_messages');

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
    expect(text).toMatch(/Transporter/i);
    await harness.close();
  });

  it('takes no arguments', async () => {
    const client = fakeClient();
    const harness = await createTestHarness((server) => registerChatTools(server, client));

    const tools = await harness.listTools();
    const tool = tools.find((t) => t.name === 'teams_get_open_chat_messages');

    expect(tool).toBeDefined();
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
