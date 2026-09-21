import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { registerChatTools } from '../../src/tools/chat.js';
import type { TeamsClient } from '../../src/client.js';

function fakeClient(overrides: Partial<TeamsClient> = {}): TeamsClient {
  return {
    listChats: vi.fn().mockResolvedValue([]),
    getOpenChatMessages: vi.fn().mockResolvedValue([]),
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
    const data = parseToolResult(result);

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
    const data = parseToolResult(result);

    expect(data).toEqual([
      { sender: 'Alice', time: '2026-09-21T10:00:00.000Z', messageId: 'timestamp-1', text: 'hi' },
    ]);
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
