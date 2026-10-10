import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import type { FetchproxyTransport } from '@chrischall/mcp-utils/fetchproxy';
import { registerHealthcheckTool } from '../../src/tools/healthcheck.js';
import type { TeamsClient } from '../../src/client.js';
import { ChatServiceClient } from '../../src/chatsvc.js';

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
    chatsvc: {
      listConversations: vi.fn(listConversations),
      tokenSource: 'browser capture',
      apiBase: 'https://teams.cloud.microsoft/api/chatsvc/amer/v1',
      tokenExpiresAt: vi.fn().mockReturnValue(null),
    } as unknown as ChatServiceClient,
  };
}

/** A JWT carrying `exp` and an invented `oid`. The signature is irrelevant here. */
function jwt(secondsFromNow: number): string {
  const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'none' })}.${b64({
    exp: Math.floor(Date.now() / 1000) + secondsFromNow,
    oid: '00000000-0000-0000-0000-0000000000a1',
  })}.signature-part-that-must-never-be-echoed`;
}

/** A REAL client over a mocked fetch: one page of `n` conversations. */
function realChatsvc(token: string, n: number): ChatServiceClient {
  const conversations = Array.from({ length: n }, (_, i) => ({ id: `19:group${i}@thread.v2` }));
  return new ChatServiceClient({
    // Explicit env, cache off — never the developer's real ~/.microsoft-teams-mcp.
    env: { TEAMS_TOKEN_CACHE: 'false' },
    captureToken: async () => token,
    fetchImpl: (async () =>
      new Response(JSON.stringify({ conversations, _metadata: {} }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })) as unknown as typeof fetch,
  });
}

type ChatServiceBlock = {
  api_base?: string;
  token_source?: string;
  token_expires_at?: string | null;
  token_minutes_remaining?: number | null;
  conversations?: number | null;
};

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

  it('reports the token source and expiry and the conversation count — never the token', async () => {
    const token = jwt(90 * 60);
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(server, {
        teams: { transport: fakeTransport() } as unknown as TeamsClient,
        chatsvc: realChatsvc(token, 3),
      }),
    );

    const result = await harness.callTool('teams_healthcheck');
    const data = parseToolResult(result) as Record<string, unknown>;
    const svc = data.chat_service as ChatServiceBlock;

    expect(data.ok).toBe(true);
    expect(svc.api_base).toBe('https://teams.cloud.microsoft/api/chatsvc/amer/v1');
    expect(svc.token_source).toBe('browser capture');
    expect(Date.parse(String(svc.token_expires_at))).toBeGreaterThan(Date.now());
    expect(svc.token_minutes_remaining).toBeGreaterThanOrEqual(88);
    expect(svc.token_minutes_remaining).toBeLessThanOrEqual(90);
    expect(svc.conversations).toBe(3);

    const text = JSON.stringify(result);
    expect(text).not.toContain(token);
    expect(text).not.toContain('signature-part-that-must-never-be-echoed');
    await harness.close();
  });

  it('reports no token expiry (null, not minutes since 1970) when the capture failed', async () => {
    const chatsvc = new ChatServiceClient({
      env: { TEAMS_TOKEN_CACHE: 'false' },
      captureToken: async () => {
        throw new Error('could not capture a Teams token from the browser (timed out). Open …');
      },
      fetchImpl: (async () => new Response('{}')) as unknown as typeof fetch,
    });
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(server, { teams: { transport: fakeTransport() } as unknown as TeamsClient, chatsvc }),
    );

    const data = parseToolResult(await harness.callTool('teams_healthcheck')) as Record<string, unknown>;
    const svc = data.chat_service as ChatServiceBlock;

    expect(data.ok).toBe(false);
    expect(svc.token_expires_at).toBeNull();
    expect(svc.token_minutes_remaining).toBeNull();
    expect(svc.conversations).toBeNull();
    await harness.close();
  });

  it('names a directly-supplied token as the source', async () => {
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(server, {
        teams: { transport: fakeTransport() } as unknown as TeamsClient,
        chatsvc: new ChatServiceClient({
          env: { TEAMS_TOKEN_CACHE: 'false', TEAMS_ACCESS_TOKEN: jwt(3600) },
          captureToken: async () => {
            throw new Error('must not capture when a token is supplied');
          },
          fetchImpl: (async () =>
            new Response(JSON.stringify({ conversations: [], _metadata: {} }), { status: 200 })) as unknown as typeof fetch,
        }),
      }),
    );

    const data = parseToolResult(await harness.callTool('teams_healthcheck')) as Record<string, unknown>;

    expect((data.chat_service as ChatServiceBlock).token_source).toBe('TEAMS_ACCESS_TOKEN');
    await harness.close();
  });

  it('still reports the chat-service block, with no expiry and no count, when no token was captured', async () => {
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(
        server,
        deps(async () => {
          throw new Error('could not capture a Teams token from the browser (timed out).');
        }),
      ),
    );

    const data = parseToolResult(await harness.callTool('teams_healthcheck')) as Record<string, unknown>;
    const svc = data.chat_service as ChatServiceBlock;

    expect(data.ok).toBe(false);
    expect(svc.token_source).toBe('browser capture');
    expect(svc.token_expires_at).toBeNull();
    expect(svc.token_minutes_remaining).toBeNull();
    expect(svc.conversations).toBeNull();
    // The bridge section stays: the token is captured through it, and Activity reads through it.
    expect(data.bridge).toBeDefined();
    await harness.close();
  });

  it('describes itself as a read-only, open-world probe of both paths', async () => {
    const harness = await createTestHarness((server) => registerHealthcheckTool(server, deps(async () => [])));

    const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_healthcheck');

    expect(tool?.annotations?.readOnlyHint).toBe(true);
    expect(tool?.annotations?.openWorldHint).toBe(true);
    expect(tool?.description).toMatch(/chat.service/i);
    expect(tool?.description).toMatch(/never the token/i);
    await harness.close();
  });
});
