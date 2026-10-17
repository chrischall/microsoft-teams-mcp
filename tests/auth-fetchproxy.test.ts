import { describe, expect, it, vi, beforeEach } from 'vitest';

const captureRequestHeader = vi.fn();
const start = vi.fn(async () => {});
const close = vi.fn(async () => {});
const createFetchproxyTransport = vi.fn((_opts: Record<string, unknown>) => ({
  start,
  close,
  server: { captureRequestHeader },
}));

vi.mock('@chrischall/mcp-utils/fetchproxy', async (importOriginal) => {
  const real = await importOriginal<typeof import('@chrischall/mcp-utils/fetchproxy')>();
  return { ...real, createFetchproxyTransport };
});

const {
  stripBearer,
  captureTokenViaFetchproxy,
  CAPTURE_DECL,
  TRUST_DOMAINS,
} = await import('../src/auth-fetchproxy.js');

// Opaque filler: long enough to pass the length guard, never a real token.
const FAKE_TOKEN = 'x'.repeat(200);

beforeEach(() => {
  captureRequestHeader.mockReset();
  start.mockClear();
  close.mockClear();
  createFetchproxyTransport.mockClear();
});

describe('stripBearer', () => {
  it('drops a Bearer prefix in any case and trims', () => {
    expect(stripBearer('Bearer abc')).toBe('abc');
    expect(stripBearer('bearer  abc ')).toBe('abc');
    expect(stripBearer('  abc  ')).toBe('abc');
  });
});

describe('capture declaration', () => {
  it('captures the authorization header off chatsvc requests on teams.cloud.microsoft', () => {
    expect(CAPTURE_DECL).toEqual({
      host: 'teams.cloud.microsoft',
      path: '/api/chatsvc/*',
      headerName: 'authorization',
    });
    expect(TRUST_DOMAINS).toEqual(['teams.cloud.microsoft']);
  });
});

describe('captureTokenViaFetchproxy', () => {
  it('captures with the declared host/path/header and returns the bare token', async () => {
    captureRequestHeader.mockResolvedValue(`Bearer ${FAKE_TOKEN}`);

    await expect(captureTokenViaFetchproxy()).resolves.toBe(FAKE_TOKEN);

    expect(start).toHaveBeenCalledTimes(1);
    expect(captureRequestHeader).toHaveBeenCalledTimes(1);
    expect(captureRequestHeader).toHaveBeenCalledWith(
      expect.objectContaining({
        host: 'teams.cloud.microsoft',
        path: '/api/chatsvc/*',
        headerName: 'authorization',
        timeoutMs: 30_000,
      }),
    );
    const opts = createFetchproxyTransport.mock.calls[0][0] as {
      domains: string[];
      port: number;
      capabilities: string[];
      captureHeaders: unknown[];
    };
    expect(opts.domains).toEqual(['teams.cloud.microsoft']);
    expect(opts.port).toBe(37_149);
    // The header capture beside mcp-utils' always-on fetch verb — and no DOM
    // read: the Activity feed comes from the chat service, not the page.
    expect([...opts.capabilities].sort()).toEqual(['capture_request_header', 'fetch']);
    expect(opts.captureHeaders).toEqual([CAPTURE_DECL]);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('derives the bridge deadline from the default 30s capture window', async () => {
    captureRequestHeader.mockResolvedValue(FAKE_TOKEN);
    await captureTokenViaFetchproxy({});

    const opts = createFetchproxyTransport.mock.calls[0][0] as { fetchTimeoutMs: number; captureWindowMs: number };
    expect(opts.fetchTimeoutMs).toBe(45_000);
    expect(opts.captureWindowMs).toBe(30_000);
  });

  it('honours TEAMS_CAPTURE_TIMEOUT in the capture window AND the bridge deadline', async () => {
    captureRequestHeader.mockResolvedValue(FAKE_TOKEN);
    await captureTokenViaFetchproxy({ TEAMS_CAPTURE_TIMEOUT: '120' });

    expect(captureRequestHeader).toHaveBeenCalledWith(expect.objectContaining({ timeoutMs: 120_000 }));
    const opts = createFetchproxyTransport.mock.calls[0][0] as { fetchTimeoutMs: number; captureWindowMs: number };
    expect(opts.fetchTimeoutMs).toBe(135_000);
    expect(opts.captureWindowMs).toBe(120_000);
  });

  it('never sets the bridge deadline under the library 30s, however short the window', async () => {
    captureRequestHeader.mockResolvedValue(FAKE_TOKEN);
    await captureTokenViaFetchproxy({ TEAMS_CAPTURE_TIMEOUT: '1' });

    const opts = createFetchproxyTransport.mock.calls[0][0] as { fetchTimeoutMs: number; captureWindowMs: number };
    expect(opts.fetchTimeoutMs).toBe(30_000);
    expect(opts.captureWindowMs).toBe(1_000);
  });

  it('reads TEAMS_CAPTURE_TIMEOUT from process.env when the capture runs, not at import', async () => {
    captureRequestHeader.mockResolvedValue(FAKE_TOKEN);
    vi.stubEnv('TEAMS_CAPTURE_TIMEOUT', '60');
    try {
      await captureTokenViaFetchproxy();
    } finally {
      vi.unstubAllEnvs();
    }
    expect(captureRequestHeader).toHaveBeenCalledWith(expect.objectContaining({ timeoutMs: 60_000 }));
    expect((createFetchproxyTransport.mock.calls[0][0] as { fetchTimeoutMs: number }).fetchTimeoutMs).toBe(75_000);
  });

  it('rejects a header too short to be a token, with the click-into-a-chat hint', async () => {
    captureRequestHeader.mockResolvedValue('Bearer short');

    const err = await captureTokenViaFetchproxy().catch((e: Error) => e);
    expect((err as Error).message).toMatch(/too short/);
    expect((err as Error).message).toMatch(/teams\.cloud\.microsoft/);
    expect((err as Error).message).toMatch(/click into any chat/i);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('rejects a capture that returned no header', async () => {
    captureRequestHeader.mockResolvedValue(undefined);
    await expect(captureTokenViaFetchproxy()).rejects.toThrow(/no header/);
  });

  it('wraps a capture timeout with actionable guidance', async () => {
    captureRequestHeader.mockRejectedValue(new Error('timed out'));
    await expect(captureTokenViaFetchproxy()).rejects.toThrow(
      /could not capture a Teams token.*timed out.*signed-in teams\.cloud\.microsoft tab/s,
    );
  });

  it('surfaces the token even when closing the transport fails', async () => {
    captureRequestHeader.mockResolvedValue(FAKE_TOKEN);
    close.mockRejectedValueOnce(new Error('already closed'));
    await expect(captureTokenViaFetchproxy()).resolves.toBe(FAKE_TOKEN);
  });
});
