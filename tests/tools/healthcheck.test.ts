import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import type { FetchproxyTransport } from '@chrischall/mcp-utils/fetchproxy';
import { registerHealthcheckTool } from '../../src/tools/healthcheck.js';
import type { TeamsClient } from '../../src/client.js';

const BRIDGE = {
  role: 'host' as const,
  port: 37149,
  server_version: 'test',
  fetch_timeout_ms: 30_000,
  last_success_at: null,
  last_failure_at: null,
  last_failure_reason: null,
  consecutive_failures: 0,
  session_state: 'ready',
  pending_pair_code: null,
  extension_connected: true,
};

/** A transport whose runProbe just runs the probe, like the real one minus the bridge. */
function fakeTransport(): FetchproxyTransport {
  return {
    status: vi.fn().mockReturnValue({ lastExtensionMessageAt: null }),
    runProbe: vi.fn(async (fn: (p: string) => Promise<unknown>, p: string) => {
      try {
        await fn(p);
        return { ok: true, elapsed_ms: 1, bridge: BRIDGE };
      } catch (e) {
        return { ok: false, elapsed_ms: 1, bridge: BRIDGE, error: { kind: 'other', message: (e as Error).message } };
      }
    }),
  } as unknown as FetchproxyTransport;
}

function fakeClient(rows: unknown[]): TeamsClient {
  return {
    listChats: vi.fn().mockResolvedValue(rows),
    transport: fakeTransport(),
  } as unknown as TeamsClient;
}

describe('teams_healthcheck', () => {
  it('reports healthy when the chat list renders', async () => {
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(server, fakeClient([{ title: 'Standup' }])),
    );

    const data = parseToolResult(await harness.callTool('teams_healthcheck')) as Record<string, unknown>;

    expect(data.ok).toBe(true);
    await harness.close();
  });

  it('does not report healthy when no chat list is rendered (wrong Teams view)', async () => {
    const harness = await createTestHarness((server) => registerHealthcheckTool(server, fakeClient([])));

    const data = parseToolResult(await harness.callTool('teams_healthcheck')) as Record<string, unknown>;

    expect(data.ok).toBe(false);
    expect((data.error as { kind?: string } | undefined)?.kind).toBe('chat_view_not_open');
    expect(String(data.hint)).toMatch(/Chat view/i);
    await harness.close();
  });
});
