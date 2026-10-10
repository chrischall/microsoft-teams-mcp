import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import type { FetchproxyTransport } from '@chrischall/mcp-utils/fetchproxy';
import { registerHealthcheckTool } from '../../src/tools/healthcheck.js';
import type { TeamsClient } from '../../src/client.js';
import type { ChatServiceClient } from '../../src/chatsvc.js';

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

function deps(listConversations: () => Promise<unknown[]>): { teams: TeamsClient; chatsvc: ChatServiceClient } {
  return {
    teams: { transport: fakeTransport() } as unknown as TeamsClient,
    chatsvc: { listConversations: vi.fn(listConversations) } as unknown as ChatServiceClient,
  };
}

describe('teams_healthcheck', () => {
  it('reports healthy when the chat service lists conversations', async () => {
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(server, deps(async () => [{ id: '19:a@thread.v2' }])),
    );

    const data = parseToolResult(await harness.callTool('teams_healthcheck')) as Record<string, unknown>;

    expect(data.ok).toBe(true);
    await harness.close();
  });

  it('reports a token-capture failure with the open-a-Teams-tab hint', async () => {
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(
        server,
        deps(async () => {
          throw new Error('could not capture a Teams token from the browser (timed out). Open …');
        }),
      ),
    );

    const data = parseToolResult(await harness.callTool('teams_healthcheck')) as Record<string, unknown>;

    expect(data.ok).toBe(false);
    expect((data.error as { kind?: string } | undefined)?.kind).toBe('token_capture');
    expect(String(data.hint)).toMatch(/click into any chat/i);
    await harness.close();
  });
});
