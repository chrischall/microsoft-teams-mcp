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
  isSharePointHost,
  captureHostFor,
  captureSharePointTokenViaFetchproxy,
  SHAREPOINT_TRUST_DOMAINS,
} = await import('../src/sharepoint-auth.js');

// Opaque filler: long enough to pass the length guard, never a real token.
const FAKE_TOKEN = 'x'.repeat(200);

// An opaque tenant host that follows the real `*-my.sharepoint.com` shape
// without naming a real tenant.
const HOST = 'contoso-my.sharepoint.com';

beforeEach(() => {
  captureRequestHeader.mockReset();
  start.mockClear();
  close.mockClear();
  createFetchproxyTransport.mockClear();
});

describe('isSharePointHost', () => {
  it('accepts concrete *.sharepoint.com tenant subdomains', () => {
    expect(isSharePointHost('transunion-my.sharepoint.com')).toBe(true);
    expect(isSharePointHost('transunion.sharepoint.com')).toBe(true);
    expect(isSharePointHost('transunionind-my.sharepoint.com')).toBe(true);
    // Case is normalised.
    expect(isSharePointHost('Contoso-MY.SharePoint.com')).toBe(true);
  });

  it('rejects non-SharePoint and spoofed hosts', () => {
    expect(isSharePointHost('evil.com')).toBe(false);
    expect(isSharePointHost('sharepoint.com.evil.com')).toBe(false);
    expect(isSharePointHost('sharepoint.com')).toBe(false);
    expect(isSharePointHost('graph.microsoft.com')).toBe(false);
    expect(isSharePointHost('')).toBe(false);
  });
});

describe('captureHostFor', () => {
  it('derives the -my OneDrive host for a tenant-root (team-site) host', () => {
    // The browser authenticates the -my host, not the root, and the -my token
    // reads root files too (verified live 2026-10-11).
    expect(captureHostFor('transunion.sharepoint.com')).toBe('transunion-my.sharepoint.com');
    expect(captureHostFor('Contoso.SharePoint.com')).toBe('contoso-my.sharepoint.com');
  });

  it('leaves a -my host (incl. a region OneDrive) unchanged', () => {
    expect(captureHostFor('transunion-my.sharepoint.com')).toBe('transunion-my.sharepoint.com');
    expect(captureHostFor('transunionind-my.sharepoint.com')).toBe('transunionind-my.sharepoint.com');
  });

  it('always returns a host that isSharePointHost accepts', () => {
    for (const h of ['transunion.sharepoint.com', 'transunion-my.sharepoint.com', 'contoso.sharepoint.com']) {
      expect(isSharePointHost(captureHostFor(h))).toBe(true);
    }
  });
});

describe('trust scope', () => {
  it('trusts the broad sharepoint.com domain (covers every tenant host)', () => {
    expect(SHAREPOINT_TRUST_DOMAINS).toEqual(['sharepoint.com']);
  });
});

describe('captureSharePointTokenViaFetchproxy', () => {
  it('captures the authorization header for the given host and returns the bare token', async () => {
    captureRequestHeader.mockResolvedValue(`Bearer ${FAKE_TOKEN}`);

    await expect(captureSharePointTokenViaFetchproxy(HOST)).resolves.toBe(FAKE_TOKEN);

    expect(start).toHaveBeenCalledTimes(1);
    expect(captureRequestHeader).toHaveBeenCalledTimes(1);
    expect(captureRequestHeader).toHaveBeenCalledWith(
      expect.objectContaining({
        host: HOST,
        path: '/*',
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
    // The trust scope is the broad family domain, but the capture decl is the
    // concrete host (HOSTNAME_RE rejects a wildcard capture host).
    expect(opts.domains).toEqual(['sharepoint.com']);
    expect(opts.port).toBe(37_149);
    expect([...opts.capabilities].sort()).toEqual(['capture_request_header', 'fetch']);
    expect(opts.captureHeaders).toEqual([{ host: HOST, path: '/*', headerName: 'authorization' }]);
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('refuses a non-SharePoint host before touching the bridge', async () => {
    await expect(captureSharePointTokenViaFetchproxy('evil.com')).rejects.toThrow(/sharepoint\.com/i);
    expect(createFetchproxyTransport).not.toHaveBeenCalled();
    expect(captureRequestHeader).not.toHaveBeenCalled();
  });

  it('derives the bridge deadline from the default 30s capture window', async () => {
    captureRequestHeader.mockResolvedValue(FAKE_TOKEN);
    await captureSharePointTokenViaFetchproxy(HOST, {});

    const opts = createFetchproxyTransport.mock.calls[0][0] as {
      fetchTimeoutMs: number;
      captureWindowMs: number;
    };
    expect(opts.fetchTimeoutMs).toBe(45_000);
    expect(opts.captureWindowMs).toBe(30_000);
  });

  it('honours TEAMS_CAPTURE_TIMEOUT in the capture window AND the bridge deadline', async () => {
    captureRequestHeader.mockResolvedValue(FAKE_TOKEN);
    await captureSharePointTokenViaFetchproxy(HOST, { TEAMS_CAPTURE_TIMEOUT: '120' });

    expect(captureRequestHeader).toHaveBeenCalledWith(expect.objectContaining({ timeoutMs: 120_000 }));
    const opts = createFetchproxyTransport.mock.calls[0][0] as {
      fetchTimeoutMs: number;
      captureWindowMs: number;
    };
    expect(opts.fetchTimeoutMs).toBe(135_000);
    expect(opts.captureWindowMs).toBe(120_000);
  });

  it('strips the Bearer prefix from the captured header', async () => {
    captureRequestHeader.mockResolvedValue(`  Bearer   ${FAKE_TOKEN}  `);
    await expect(captureSharePointTokenViaFetchproxy(HOST)).resolves.toBe(FAKE_TOKEN);
  });

  it('rejects a header too short to be a token, naming the host to open', async () => {
    captureRequestHeader.mockResolvedValue('Bearer short');

    const err = await captureSharePointTokenViaFetchproxy(HOST).catch((e: Error) => e);
    expect((err as Error).message).toMatch(/too short/);
    expect((err as Error).message).toMatch(new RegExp(HOST.replace(/\./g, '\\.')));
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('rejects a capture that returned no header', async () => {
    captureRequestHeader.mockResolvedValue(undefined);
    await expect(captureSharePointTokenViaFetchproxy(HOST)).rejects.toThrow(/no header/);
  });

  it('wraps a capture failure with guidance to interact with Teams / open a OneDrive file', async () => {
    captureRequestHeader.mockRejectedValue(new Error('timed out'));
    const err = await captureSharePointTokenViaFetchproxy(HOST).catch((e: Error) => e);
    expect((err as Error).message).toMatch(/timed out/);
    // The token is captured from the OneDrive host, not by opening the team-site
    // file, so the hint points at interacting with Teams / one's own OneDrive.
    expect((err as Error).message).toMatch(/open one of your own OneDrive files/i);
    expect((err as Error).message).toMatch(new RegExp(HOST.replace(/\./g, '\\.')));
  });

  it('surfaces the token even when closing the transport fails', async () => {
    captureRequestHeader.mockResolvedValue(FAKE_TOKEN);
    close.mockRejectedValueOnce(new Error('already closed'));
    await expect(captureSharePointTokenViaFetchproxy(HOST)).resolves.toBe(FAKE_TOKEN);
  });
});
