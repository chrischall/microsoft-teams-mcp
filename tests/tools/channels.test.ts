import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { registerChannelTools } from '../../src/tools/channels.js';
import type { TeamsClient } from '../../src/client.js';

function fakeClient(overrides: Partial<TeamsClient> = {}): TeamsClient {
  return {
    listTeamsAndChannels: vi.fn().mockResolvedValue([]),
    getOpenChannelPosts: vi.fn().mockResolvedValue([]),
    ...overrides,
  } as unknown as TeamsClient;
}

describe('teams_list_teams_and_channels', () => {
  it('returns the teams and channels from the client', async () => {
    const client = fakeClient({
      listTeamsAndChannels: vi.fn().mockResolvedValue([
        { title: 'General', teamName: 'TruAudience', time: '9/18', conversationKey: 'k', itemType: 'channel' },
        { title: 'TruAudience', time: '', conversationKey: 'k2', itemType: 'team' },
      ]),
    });
    const harness = await createTestHarness((server) => registerChannelTools(server, client));

    const result = await harness.callTool('teams_list_teams_and_channels');
    const data = (parseToolResult(result) as { rows: unknown }).rows;

    expect(data).toEqual([
      { title: 'General', teamName: 'TruAudience', time: '9/18', conversationKey: 'k', itemType: 'channel' },
      { title: 'TruAudience', time: '', conversationKey: 'k2', itemType: 'team' },
    ]);
    await harness.close();
  });

  it('surfaces a no-tab failure with an actionable hint', async () => {
    const client = fakeClient({
      listTeamsAndChannels: vi
        .fn()
        .mockRejectedValue(new Error('could not reach a signed-in Teams tab on any of […]')),
    });
    const harness = await createTestHarness((server) => registerChannelTools(server, client));

    const result = await harness.callTool('teams_list_teams_and_channels');

    expect(result.isError).toBe(true);
    const text = (result.content?.[0] as { text?: string })?.text ?? '';
    expect(text).toMatch(/signed-in Teams tab/i);
    await harness.close();
  });
});

describe('teams_get_open_channel_posts', () => {
  it('returns posts from the client', async () => {
    const client = fakeClient({
      getOpenChannelPosts: vi.fn().mockResolvedValue([
        {
          sender: 'Alice',
          subject: 'Deploy freeze',
          time: 'Tuesday, August 4, 2026 7:50 AM',
          messageId: 'timestamp-1',
          text: 'hi',
        },
      ]),
    });
    const harness = await createTestHarness((server) => registerChannelTools(server, client));

    const result = await harness.callTool('teams_get_open_channel_posts');
    const data = (parseToolResult(result) as { rows: unknown }).rows;

    expect(data).toEqual([
      {
        sender: 'Alice',
        subject: 'Deploy freeze',
        time: 'Tuesday, August 4, 2026 7:50 AM',
        messageId: 'timestamp-1',
        text: 'hi',
      },
    ]);
    await harness.close();
  });

  it('surfaces a not-yet-approved bridge failure with an actionable hint', async () => {
    const client = fakeClient({
      getOpenChannelPosts: vi
        .fn()
        .mockRejectedValue(new Error('read_dom_list name not in declared set: channelPosts')),
    });
    const harness = await createTestHarness((server) => registerChannelTools(server, client));

    const result = await harness.callTool('teams_get_open_channel_posts');

    expect(result.isError).toBe(true);
    const text = (result.content?.[0] as { text?: string })?.text ?? '';
    expect(text).toMatch(/Transporter/i);
    await harness.close();
  });

  it('takes no arguments', async () => {
    const client = fakeClient();
    const harness = await createTestHarness((server) => registerChannelTools(server, client));

    const tools = await harness.listTools();
    expect(tools.find((t) => t.name === 'teams_get_open_channel_posts')).toBeDefined();
    await harness.close();
  });
});

describe('tool annotations', () => {
  it('both channel tools are read-only', async () => {
    const client = fakeClient();
    const harness = await createTestHarness((server) => registerChannelTools(server, client));
    const tools = (await harness.client.listTools()).tools;

    for (const name of ['teams_list_teams_and_channels', 'teams_get_open_channel_posts']) {
      expect(tools.find((t) => t.name === name)?.annotations?.readOnlyHint, name).toBe(true);
    }
    await harness.close();
  });
});
