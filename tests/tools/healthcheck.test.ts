import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { registerHealthcheckTool } from '../../src/tools/healthcheck.js';
import { ChatServiceClient } from '../../src/chatsvc.js';

function deps(
  listConversations: () => Promise<unknown[]>,
  ensureToken: () => Promise<void> = async () => {},
): { chatsvc: ChatServiceClient } {
  return {
    chatsvc: {
      ensureToken: vi.fn(ensureToken),
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
      registerHealthcheckTool(server, { chatsvc }),
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
    const listConversations = vi.fn(async () => []);
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(
        server,
        deps(listConversations, async () => {
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
    expect((data.error as { kind?: string } | undefined)?.kind).toBe('token_capture');
    expect(String(data.hint)).toMatch(/click into any chat/i);
    expect((data.credential as { resolved?: boolean }).resolved).toBe(false);
    // No token, no probe: a listing without one would only 401.
    expect(listConversations).not.toHaveBeenCalled();
    await harness.close();
  });

  it('reports no_credential when the bridge is disabled and no token is supplied', async () => {
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(server, {
        chatsvc: new ChatServiceClient({ env: { TEAMS_TOKEN_CACHE: 'false', TEAMS_DISABLE_FETCHPROXY: '1' } }),
      }),
    );

    const data = parseToolResult(await harness.callTool('teams_healthcheck')) as Record<string, unknown>;

    expect(data.ok).toBe(false);
    expect((data.error as { kind?: string } | undefined)?.kind).toBe('no_credential');
    await harness.close();
  });

  it('has no bridge section — the bridge is only touched to capture the token', async () => {
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(server, deps(async () => [{ id: '19:a@thread.v2' }])),
    );

    const data = parseToolResult(await harness.callTool('teams_healthcheck')) as Record<string, unknown>;

    expect(data).not.toHaveProperty('bridge');
    expect(data.credential).toMatchObject({ source: 'browser capture', resolved: true });
    expect(data.probe).toMatchObject({ url: expect.stringContaining('teams.cloud.microsoft') });
    await harness.close();
  });

  it('notes the SharePoint token is captured on demand per host and short-lived, without needing one to pass', async () => {
    // A healthy run holds a chat-service token but no SharePoint token (none is
    // minted until a shared file is opened). The note must appear and health
    // must not depend on an SP token existing.
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(server, deps(async () => [{ id: '19:a@thread.v2' }])),
    );

    const data = parseToolResult(await harness.callTool('teams_healthcheck')) as Record<string, unknown>;

    expect(data.ok).toBe(true);
    const sharepoint = String(data.sharepoint);
    expect(sharepoint).toMatch(/on demand/i);
    expect(sharepoint).toMatch(/per (sharepoint )?host/i);
    expect(sharepoint).toMatch(/short-lived/i);
    await harness.close();
  });

  it('describes itself as a read-only, open-world probe of the chat service', async () => {
    const harness = await createTestHarness((server) => registerHealthcheckTool(server, deps(async () => [])));

    const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_healthcheck');

    expect(tool?.annotations?.readOnlyHint).toBe(true);
    expect(tool?.annotations?.openWorldHint).toBe(true);
    expect(tool?.description).toMatch(/chat.service/i);
    expect(tool?.description).toMatch(/never the token/i);
    await harness.close();
  });
});
