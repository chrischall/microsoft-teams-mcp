import { mkdtempSync, readFileSync, writeFileSync, rmSync, mkdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

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

const { BRIDGE_TRUST_DOMAINS, bridgeCaptureDecls, knownSharePointHosts, rememberSharePointHost, bridgeScopePath } =
  await import('../src/bridge-scope.js');
const { captureTokenViaFetchproxy, CAPTURE_DECL } = await import('../src/auth-fetchproxy.js');
const { captureSubstrateTokenViaFetchproxy, SUBSTRATE_CAPTURE_DECL } = await import('../src/substrate-auth.js');
const { captureSharePointTokenViaFetchproxy } = await import('../src/sharepoint-auth.js');

// Opaque filler: long enough to pass the length guard, never a real token.
const FAKE_TOKEN = 'x'.repeat(200);
const SP_A = 'contoso-my.sharepoint.com';
const SP_B = 'contosoind-my.sharepoint.com';

let dir: string;
let env: NodeJS.ProcessEnv;

beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'teams-bridge-scope-'));
  env = { TEAMS_SHAREPOINT_TOKEN_FILE: join(dir, 'sp.json') } as NodeJS.ProcessEnv;
  captureRequestHeader.mockReset();
  captureRequestHeader.mockResolvedValue(`Bearer ${FAKE_TOKEN}`);
  createFetchproxyTransport.mockClear();
});

afterEach(() => rmSync(dir, { recursive: true, force: true }));

type HelloOpts = { domains: string[]; captureHeaders: { host: string; path?: string; headerName: string }[]; capabilities: string[] };
const lastHello = (): HelloOpts => createFetchproxyTransport.mock.calls.at(-1)![0] as unknown as HelloOpts;

describe('the one bridge scope every capture declares', () => {
  it('trusts the three hosts the server ever captures from', () => {
    expect([...BRIDGE_TRUST_DOMAINS].sort()).toEqual(['sharepoint.com', 'substrate.office.com', 'teams.cloud.microsoft']);
  });

  it('declares the chat-service and search captures, then each known SharePoint host, in a stable order', () => {
    expect(bridgeCaptureDecls([SP_B, SP_A, SP_B])).toEqual([
      CAPTURE_DECL,
      SUBSTRATE_CAPTURE_DECL,
      { host: SP_A, path: '/*', headerName: 'authorization' },
      { host: SP_B, path: '/*', headerName: 'authorization' },
    ]);
  });

  it('drops anything that is not a concrete SharePoint host', () => {
    expect(bridgeCaptureDecls(['evil.com', 'sharepoint.com', SP_A])).toEqual([
      CAPTURE_DECL,
      SUBSTRATE_CAPTURE_DECL,
      { host: SP_A, path: '/*', headerName: 'authorization' },
    ]);
  });

  // The regression: the bridge re-pairs whenever the declared domain set
  // differs from the approved one, and replaces the approved capture list on
  // every grant — so captures that each declared only their own host
  // re-prompted on every switch between chat, search and files.
  it('is identical for the chat-service, search and SharePoint captures', async () => {
    rememberSharePointHost(SP_A, env);

    await captureTokenViaFetchproxy(env);
    const chat = lastHello();
    await captureSubstrateTokenViaFetchproxy(env);
    const search = lastHello();
    await captureSharePointTokenViaFetchproxy(SP_A, env);
    const files = lastHello();

    for (const hello of [chat, search, files]) {
      expect([...hello.domains].sort()).toEqual([...BRIDGE_TRUST_DOMAINS].sort());
      expect(hello.captureHeaders).toEqual(bridgeCaptureDecls([SP_A]));
      expect([...hello.capabilities].sort()).toEqual(['capture_request_header', 'fetch']);
    }
  });

  it('still captures only the header the caller asked for', async () => {
    await captureSubstrateTokenViaFetchproxy(env);
    expect(captureRequestHeader).toHaveBeenCalledTimes(1);
    expect(captureRequestHeader).toHaveBeenCalledWith(expect.objectContaining({ host: 'substrate.office.com' }));
  });
});

describe('known SharePoint hosts', () => {
  it('starts empty and lives beside the SharePoint token cache', () => {
    expect(knownSharePointHosts(env)).toEqual([]);
    expect(bridgeScopePath(env)).toBe(join(dir, 'bridge-scope.json'));
  });

  it('is remembered across runs once a capture of that host succeeds', async () => {
    await captureSharePointTokenViaFetchproxy(SP_B, env);
    expect(JSON.parse(readFileSync(bridgeScopePath(env), 'utf8')).state).toEqual({ sharePointHosts: [SP_B] });
    expect(knownSharePointHosts(env)).toEqual([SP_B]);

    // A later chat-service capture declares it too, so the scope does not shrink.
    await captureTokenViaFetchproxy(env);
    expect(lastHello().captureHeaders).toEqual(bridgeCaptureDecls([SP_B]));
  });

  it('declares a new host on the capture that needs it, but remembers it only once that capture succeeds', async () => {
    captureRequestHeader.mockRejectedValueOnce(new Error('timed out'));
    await expect(captureSharePointTokenViaFetchproxy(SP_A, env)).rejects.toThrow();
    expect(lastHello().captureHeaders).toEqual(bridgeCaptureDecls([SP_A]));
    expect(knownSharePointHosts(env)).toEqual([]);

    // Not approved yet, so a chat-service capture does not grow the scope with it.
    await captureTokenViaFetchproxy(env);
    expect(lastHello().captureHeaders).toEqual(bridgeCaptureDecls([]));
  });

  it('ignores a corrupt or hostile file rather than failing a capture', async () => {
    mkdirSync(dir, { recursive: true });
    writeFileSync(
      bridgeScopePath(env),
      JSON.stringify({ v: 1, state: { sharePointHosts: ['evil.com', 7, 'contoso-my.sharepoint.com'] } }),
    );
    expect(knownSharePointHosts(env)).toEqual([SP_A]);
    writeFileSync(bridgeScopePath(env), 'not json');
    expect(knownSharePointHosts(env)).toEqual([]);
    await expect(captureTokenViaFetchproxy(env)).resolves.toBe(FAKE_TOKEN);
  });

  it('is not written when TEAMS_TOKEN_CACHE is off', () => {
    const off = { ...env, TEAMS_TOKEN_CACHE: 'false' };
    rememberSharePointHost(SP_A, off);
    expect(() => readFileSync(bridgeScopePath(off))).toThrow();
  });
});
