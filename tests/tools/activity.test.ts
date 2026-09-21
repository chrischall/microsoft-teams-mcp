import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { registerActivityTools } from '../../src/tools/activity.js';
import type { TeamsClient } from '../../src/client.js';

function fakeClient(overrides: Partial<TeamsClient> = {}): TeamsClient {
  return {
    getActivity: vi.fn().mockResolvedValue([]),
    ...overrides,
  } as unknown as TeamsClient;
}

describe('teams_get_activity', () => {
  it('returns the activity feed from the client', async () => {
    const client = fakeClient({
      getActivity: vi.fn().mockResolvedValue([
        {
          title: 'Sanderson, Xi mentioned MS DevOps',
          preview: 'MS DevOps can someone please help with this harness deployment job?',
          time: '1:19 PM',
          location: 'NS_MS Product CE > MS DevOps',
        },
      ]),
    });
    const harness = await createTestHarness((server) => registerActivityTools(server, client));

    const result = await harness.callTool('teams_get_activity');
    const data = parseToolResult(result);

    expect(data).toEqual([
      {
        title: 'Sanderson, Xi mentioned MS DevOps',
        preview: 'MS DevOps can someone please help with this harness deployment job?',
        time: '1:19 PM',
        location: 'NS_MS Product CE > MS DevOps',
      },
    ]);
    await harness.close();
  });

  it('surfaces a no-tab failure with an actionable hint', async () => {
    const client = fakeClient({
      getActivity: vi.fn().mockRejectedValue(new Error('could not reach a signed-in Teams tab on teams.cloud.microsoft (…)')),
    });
    const harness = await createTestHarness((server) => registerActivityTools(server, client));

    const result = await harness.callTool('teams_get_activity');

    expect(result.isError).toBe(true);
    const text = (result.content?.[0] as { text?: string })?.text ?? '';
    expect(text).toMatch(/signed-in Teams tab/i);
    expect(text).toMatch(/teams.cloud.microsoft/i);
    await harness.close();
  });

  it('takes no arguments', async () => {
    const client = fakeClient();
    const harness = await createTestHarness((server) => registerActivityTools(server, client));

    const tools = await harness.listTools();
    const tool = tools.find((t) => t.name === 'teams_get_activity');

    expect(tool).toBeDefined();
    await harness.close();
  });

  it('is read-only', async () => {
    const client = fakeClient();
    const harness = await createTestHarness((server) => registerActivityTools(server, client));
    const tools = (await harness.client.listTools()).tools;

    expect(tools.find((t) => t.name === 'teams_get_activity')?.annotations?.readOnlyHint).toBe(true);
    await harness.close();
  });
});
