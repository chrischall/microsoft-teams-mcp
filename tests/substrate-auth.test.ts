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

const { captureSubstrateTokenViaFetchproxy, SUBSTRATE_CAPTURE_DECL, SUBSTRATE_TRUST_DOMAINS } = await import(
  '../src/substrate-auth.js'
);

// Opaque filler: long enough to pass the length guard, never a real token.
const FAKE_TOKEN = 'x'.repeat(200);

beforeEach(() => {
  captureRequestHeader.mockReset();
  start.mockClear();
  close.mockClear();
  createFetchproxyTransport.mockClear();
});

describe('substrate trust scope', () => {
  it('trusts exactly substrate.office.com, and captures off its search service', () => {
    expect(SUBSTRATE_TRUST_DOMAINS).toEqual(['substrate.office.com']);
    expect(SUBSTRATE_CAPTURE_DECL).toEqual({
      host: 'substrate.office.com',
      path: '/searchservice/*',
      headerName: 'authorization',
    });
  });
});

describe('captureSubstrateTokenViaFetchproxy', () => {
  it('captures the authorization header off a search-service request and returns the bare token', async () => {
    captureRequestHeader.mockResolvedValue(`Bearer ${FAKE_TOKEN}`);

    await expect(captureSubstrateTokenViaFetchproxy({})).resolves.toBe(FAKE_TOKEN);

    expect(start).toHaveBeenCalledTimes(1);
    expect(captureRequestHeader).toHaveBeenCalledWith(
      expect.objectContaining({
        host: 'substrate.office.com',
        path: '/searchservice/*',
        headerName: 'authorization',
        timeoutMs: 30_000,
      }),
    );
    const opts = createFetchproxyTransport.mock.calls[0][0] as {
      domains: string[];
      port: number;
      capabilities: string[];
      captureHeaders: unknown[];
      fetchTimeoutMs: number;
    };
    expect(opts.domains).toEqual(['substrate.office.com']);
    expect(opts.port).toBe(37_149);
    expect([...opts.capabilities].sort()).toEqual(['capture_request_header', 'fetch']);
    expect(opts.captureHeaders).toEqual([SUBSTRATE_CAPTURE_DECL]);
    expect(opts.fetchTimeoutMs).toBe(45_000);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('honours TEAMS_CAPTURE_TIMEOUT in the capture window and the bridge deadline', async () => {
    captureRequestHeader.mockResolvedValue(FAKE_TOKEN);
    await captureSubstrateTokenViaFetchproxy({ TEAMS_CAPTURE_TIMEOUT: '120' });
    expect(captureRequestHeader).toHaveBeenCalledWith(expect.objectContaining({ timeoutMs: 120_000 }));
    const opts = createFetchproxyTransport.mock.calls[0][0] as { fetchTimeoutMs: number; captureWindowMs: number };
    expect(opts.fetchTimeoutMs).toBe(135_000);
    expect(opts.captureWindowMs).toBe(120_000);
  });

  it('rejects a header too short to be a token', async () => {
    captureRequestHeader.mockResolvedValue('Bearer short');
    await expect(captureSubstrateTokenViaFetchproxy({})).rejects.toThrow(/too short/);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('rejects a capture that returned no header', async () => {
    captureRequestHeader.mockResolvedValue(undefined);
    await expect(captureSubstrateTokenViaFetchproxy({})).rejects.toThrow(/no header/);
  });

  it('wraps a capture failure with guidance to run a search in the Teams tab', async () => {
    captureRequestHeader.mockRejectedValue(new Error('timed out'));
    const err = (await captureSubstrateTokenViaFetchproxy({}).catch((e: Error) => e)) as Error;
    expect(err.message).toMatch(/could not capture a Substrate search token/);
    expect(err.message).toMatch(/timed out/);
    expect(err.message).toMatch(/search box/i);
  });

  it('surfaces the token even when closing the transport fails', async () => {
    captureRequestHeader.mockResolvedValue(FAKE_TOKEN);
    close.mockRejectedValueOnce(new Error('already closed'));
    await expect(captureSubstrateTokenViaFetchproxy({})).resolves.toBe(FAKE_TOKEN);
  });
});
