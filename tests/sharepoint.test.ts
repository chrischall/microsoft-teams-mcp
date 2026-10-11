import { describe, expect, it, vi } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import { SharePointClient } from '../src/sharepoint.js';

/**
 * A JWT carrying `exp`. The signature is irrelevant — `decodeJwtExp` only reads
 * the payload. Opaque filler, never a real token.
 */
function jwt(secondsFromNow = 3000): string {
  const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'none' })}.${b64({ exp: Math.floor(Date.now() / 1000) + secondsFromNow })}.x`;
}

/**
 * Test env. Every client gets an EXPLICIT env with the cache off — otherwise it
 * defaults on and resolves to the developer's real `~/.microsoft-teams-mcp`.
 */
function env(overrides: Record<string, string> = {}): NodeJS.ProcessEnv {
  return { TEAMS_TOKEN_CACHE: 'false', ...overrides };
}

type FetchMock = ReturnType<typeof vi.fn<(input: RequestInfo | URL, init?: RequestInit) => Promise<Response>>>;

function client(fetchImpl: FetchMock, capture: (host: string) => Promise<string> = async () => jwt()) {
  return new SharePointClient({
    env: env(),
    captureToken: capture,
    fetchImpl: fetchImpl as unknown as typeof fetch,
  });
}

function urlOf(mock: FetchMock, n: number): string {
  return String(mock.mock.calls[n]?.[0]);
}

function headersOf(mock: FetchMock, n: number): Headers {
  return new Headers(mock.mock.calls[n]?.[1]?.headers);
}

/** The unpadded base64url share id the client must build, derived independently. */
function shareOf(url: string): string {
  const b64 = Buffer.from(url, 'utf8')
    .toString('base64')
    .replace(/=+$/, '')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');
  return `u!${b64}`;
}

// Opaque fillers — never a real tenant host or a real file name.
const HOST = 'contoso-my.sharepoint.com';
const OBJECT_URL = `https://${HOST}/personal/user_contoso_com/Documents/report.pdf`;
const DOWNLOAD_URL = 'https://cdn.example.test/blob?sig=opaque-presigned-token';

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

function driveItem(downloadUrl = DOWNLOAD_URL, extra: Record<string, unknown> = {}): Response {
  return jsonResponse({ id: 'item-opaque', size: 3, '@content.downloadUrl': downloadUrl, ...extra });
}

function bytesResponse(bytes: number[], contentType = 'image/png', status = 200): Response {
  return new Response(new Uint8Array(bytes), { status, headers: { 'content-type': contentType } });
}

/** A fetch that answers the driveItem call then the download-url call in order. */
function twoStep(meta: Response, payload: Response): FetchMock {
  return vi.fn(async (input) => (String(input).includes('/_api/v2.1/shares/') ? meta : payload));
}

describe('SharePointClient.downloadFile', () => {
  it('builds the v2.1 shares driveItem URL with the unpadded base64url of the objectUrl', async () => {
    const fetchImpl = twoStep(driveItem(), bytesResponse([0x89, 0x50, 0x4e]));
    await client(fetchImpl).downloadFile(OBJECT_URL, 25 * 1024 * 1024);
    expect(urlOf(fetchImpl, 0)).toBe(
      `https://${HOST}/_api/v2.1/shares/${shareOf(OBJECT_URL)}/driveItem?$select=id,size,@content.downloadUrl`,
    );
  });

  it('sends the bearer on the driveItem call and NO authorization on the download-url call', async () => {
    const token = jwt();
    const fetchImpl = twoStep(driveItem(), bytesResponse([1, 2, 3]));
    await client(fetchImpl, async () => token).downloadFile(OBJECT_URL, 25 * 1024 * 1024);

    const authed = headersOf(fetchImpl, 0);
    expect(authed.get('authorization')).toBe(`Bearer ${token}`);
    expect(authed.get('behavioroverride')).toBe('redirectAs404');

    // The pre-signed CDN link must never carry the SharePoint bearer.
    expect(urlOf(fetchImpl, 1)).toBe(DOWNLOAD_URL);
    expect(headersOf(fetchImpl, 1).get('authorization')).toBeNull();
    expect(fetchImpl.mock.calls[1]?.[1]?.redirect).toBe('follow');
  });

  it('returns the mime type (without params), the bytes and their length', async () => {
    const fetchImpl = twoStep(driveItem(), bytesResponse([0xff, 0xd8, 0xff], 'image/png; charset=binary'));
    const out = await client(fetchImpl).downloadFile(OBJECT_URL, 25 * 1024 * 1024);
    expect(out.mimeType).toBe('image/png');
    expect(Buffer.isBuffer(out.data)).toBe(true);
    expect([...out.data]).toEqual([0xff, 0xd8, 0xff]);
    expect(out.size).toBe(3);
  });

  it.each([
    'https://evil.com/personal/x/Documents/report.pdf',
    'https://sharepoint.com/personal/x/Documents/report.pdf',
    'https://contoso-my.sharepoint.com.evil.com/x/report.pdf',
    'http://contoso-my.sharepoint.com/x/report.pdf',
    'not a url',
  ])('refuses %s before any fetch', async (bad) => {
    const fetchImpl: FetchMock = vi.fn(async () => bytesResponse([1]));
    await expect(client(fetchImpl).downloadFile(bad, 25 * 1024 * 1024)).rejects.toBeInstanceOf(McpToolError);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('surfaces a 403 on the driveItem call as a forbidden McpToolError, with no second fetch', async () => {
    const fetchImpl = twoStep(jsonResponse({ error: 'denied' }, 403), bytesResponse([1]));
    const p = client(fetchImpl).downloadFile(OBJECT_URL, 25 * 1024 * 1024);
    await expect(p).rejects.toBeInstanceOf(McpToolError);
    await expect(p).rejects.toMatchObject({ status: 403 });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('maps a 404 on the driveItem call to a not-found McpToolError', async () => {
    const fetchImpl = twoStep(jsonResponse({}, 404), bytesResponse([1]));
    await expect(client(fetchImpl).downloadFile(OBJECT_URL, 25 * 1024 * 1024)).rejects.toMatchObject({ status: 404 });
  });

  it('passes maxBytes to the bounded read of the download url', async () => {
    // A body whose Content-Length exceeds the cap is refused by fetchBounded
    // before it is read — proving the cap reached the second GET.
    const tooBig = new Response(new Uint8Array(32), {
      status: 200,
      headers: { 'content-type': 'application/octet-stream', 'content-length': '32' },
    });
    const fetchImpl = twoStep(driveItem(), tooBig);
    await expect(client(fetchImpl).downloadFile(OBJECT_URL, 8)).rejects.toBeInstanceOf(McpToolError);
  });

  it('refuses a download url that is not https, before fetching it', async () => {
    const fetchImpl = twoStep(driveItem('http://cdn.example.test/blob'), bytesResponse([1]));
    await expect(client(fetchImpl).downloadFile(OBJECT_URL, 25 * 1024 * 1024)).rejects.toBeInstanceOf(McpToolError);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('re-captures once and replays once on a 401 from the driveItem call', async () => {
    const capture = vi.fn(async () => jwt());
    let n = 0;
    const fetchImpl: FetchMock = vi.fn(async (input) => {
      if (String(input).includes('/_api/v2.1/shares/')) return n++ === 0 ? jsonResponse({}, 401) : driveItem();
      return bytesResponse([7]);
    });
    const out = await client(fetchImpl, capture).downloadFile(OBJECT_URL, 25 * 1024 * 1024);
    expect(out.size).toBe(1);
    expect(capture).toHaveBeenCalledTimes(2);
  });

  it('captures one token per host and reuses it across files on that host', async () => {
    const capture = vi.fn(async () => jwt());
    // Fresh responses per call — a `Response` body can only be read once.
    const fetchImpl: FetchMock = vi.fn(async (input) =>
      String(input).includes('/_api/v2.1/shares/') ? driveItem() : bytesResponse([1]),
    );
    const c = client(fetchImpl, capture);
    await c.downloadFile(OBJECT_URL, 25 * 1024 * 1024);
    await c.downloadFile(`https://${HOST}/personal/user_contoso_com/Documents/other.pdf`, 25 * 1024 * 1024);
    expect(capture).toHaveBeenCalledTimes(1);
  });
});

describe('SharePointClient configuration', () => {
  it('does not throw with the bridge disabled, and defers a no_credential error to downloadFile', async () => {
    const c = new SharePointClient({ env: env({ TEAMS_DISABLE_FETCHPROXY: '1' }) });
    const p = c.downloadFile(OBJECT_URL, 25 * 1024 * 1024);
    await expect(p).rejects.toBeInstanceOf(McpToolError);
    await expect(p).rejects.toMatchObject({ kind: 'no_credential' });
  });

  it('reports whether it is configured', () => {
    expect(new SharePointClient({ env: env() }).isConfigured()).toBe(true);
    expect(new SharePointClient({ env: env({ TEAMS_DISABLE_FETCHPROXY: '1' }) }).isConfigured()).toBe(false);
  });
});
