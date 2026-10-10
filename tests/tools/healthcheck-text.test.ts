// runCredentialHealthcheck is mocked here (and only here) to return plain text,
// which the real one never does today but nothing guarantees it won't.
import { describe, it, expect, vi } from 'vitest';
import { createTestHarness } from '@chrischall/mcp-utils/test';
import type { ChatServiceClient } from '../../src/chatsvc.js';

vi.mock('@chrischall/mcp-utils/healthcheck', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@chrischall/mcp-utils/healthcheck')>()),
  runCredentialHealthcheck: vi.fn(async () => ({ content: [{ type: 'text', text: 'healthcheck unavailable: not JSON' }] })),
}));

const { registerHealthcheckTool } = await import('../../src/tools/healthcheck.js');

describe('teams_healthcheck with a non-JSON healthcheck result', () => {
  it('keeps the healthcheck text and still reports the chat-service block', async () => {
    const chatsvc = {
      listConversations: vi.fn().mockResolvedValue([]),
      tokenSource: 'browser capture',
      apiBase: 'https://teams.cloud.microsoft/api/chatsvc/amer/v1',
      tokenExpiresAt: vi.fn().mockReturnValue(null),
    } as unknown as ChatServiceClient;
    const harness = await createTestHarness((server) =>
      registerHealthcheckTool(server, { chatsvc }),
    );
    try {
      const result = await harness.callTool('teams_healthcheck', {});
      const texts = (result.content as { type: string; text: string }[]).map((c) => c.text);

      expect(result.isError).toBeFalsy();
      expect(texts[0]).toBe('healthcheck unavailable: not JSON');
      expect(JSON.parse(texts[1])).toMatchObject({ chat_service: { token_source: 'browser capture' } });
    } finally {
      await harness.close();
    }
  });
});
