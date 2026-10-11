import { describe, expect, it, vi, afterEach } from 'vitest';
import { mkdtempSync, existsSync, writeFileSync, mkdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  createSharePointTokenCache,
  sharePointTokenCachePath,
} from '../src/sharepoint-token-cache.js';

const dir = () => mkdtempSync(join(tmpdir(), 'teams-sp-cache-'));

// Opaque fillers — never a real tenant host or a real token.
const HOST_A = 'contoso-my.sharepoint.com';
const HOST_B = 'contoso.sharepoint.com';

afterEach(() => vi.restoreAllMocks());

describe('sharepoint token cache', () => {
  it('resolves ~/.microsoft-teams-mcp/sharepoint-tokens.json by default', () => {
    const p = sharePointTokenCachePath({ HOME: '/home/x' } as NodeJS.ProcessEnv);
    expect(p).toContain('.microsoft-teams-mcp');
    expect(p.endsWith('sharepoint-tokens.json')).toBe(true);
  });

  it('honours TEAMS_SHAREPOINT_TOKEN_FILE', () => {
    const p = join(dir(), 'sp.json');
    expect(
      sharePointTokenCachePath({ TEAMS_SHAREPOINT_TOKEN_FILE: p } as NodeJS.ProcessEnv),
    ).toBe(p);
  });

  it('is off when explicitly disabled', () => {
    expect(
      createSharePointTokenCache({ TEAMS_TOKEN_CACHE: 'false' } as NodeJS.ProcessEnv),
    ).toBeNull();
  });

  it('caches nothing when the bridge — the only SP token source — is disabled', () => {
    expect(
      createSharePointTokenCache({
        TEAMS_DISABLE_FETCHPROXY: '1',
        TEAMS_SHAREPOINT_TOKEN_FILE: join(dir(), 'sp.json'),
      } as NodeJS.ProcessEnv),
    ).toBeNull();
  });

  it('round-trips one host without clobbering a sibling', () => {
    const file = join(dir(), 'sharepoint-tokens.json');
    const cache = createSharePointTokenCache({
      TEAMS_SHAREPOINT_TOKEN_FILE: file,
    } as NodeJS.ProcessEnv);
    expect(cache).not.toBeNull();

    const a = cache!.forHost(HOST_A);
    const b = cache!.forHost(HOST_B);
    const tokA = { accessToken: 'aaaa', expiresAt: Date.now() + 1000 };
    const tokB = { accessToken: 'bbbb', expiresAt: Date.now() + 2000 };

    a.save(tokA);
    b.save(tokB);
    expect(existsSync(file)).toBe(true);
    expect(a.load()).toMatchObject({ accessToken: 'aaaa' });
    expect(b.load()).toMatchObject({ accessToken: 'bbbb' });

    // Clearing one host leaves the sibling intact.
    a.clear();
    expect(a.load()).toBeNull();
    expect(b.load()).toMatchObject({ accessToken: 'bbbb' });
  });

  it('drops a corrupt entry without nuking its sibling', () => {
    const file = join(dir(), 'sharepoint-tokens.json');
    // Written straight to disk: one malformed slice, one good one.
    writeFileSync(
      file,
      JSON.stringify({
        [HOST_A]: { accessToken: '', expiresAt: 1 }, // empty token → malformed
        [HOST_B]: { accessToken: 'good', expiresAt: Date.now() + 1000 },
      }),
    );
    const cache = createSharePointTokenCache({
      TEAMS_SHAREPOINT_TOKEN_FILE: file,
    } as NodeJS.ProcessEnv)!;
    expect(cache.forHost(HOST_A).load()).toBeNull();
    expect(cache.forHost(HOST_B).load()).toMatchObject({ accessToken: 'good' });
  });

  it('treats a non-object blob as empty rather than crashing', () => {
    const file = join(dir(), 'sharepoint-tokens.json');
    writeFileSync(file, JSON.stringify(['not', 'a', 'map']));
    const cache = createSharePointTokenCache({
      TEAMS_SHAREPOINT_TOKEN_FILE: file,
    } as NodeJS.ProcessEnv)!;
    expect(cache.forHost(HOST_A).load()).toBeNull();
  });

  it('reports a failed write without throwing into the caller', () => {
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const file = join(dir(), 'sharepoint-tokens.json');
    const cache = createSharePointTokenCache({
      TEAMS_SHAREPOINT_TOKEN_FILE: file,
    } as NodeJS.ProcessEnv)!;
    const host = cache.forHost(HOST_A);
    // A directory sitting where the token file should go makes the atomic
    // rename-over-target fail, so the underlying save throws.
    mkdirSync(file);
    expect(() =>
      host.save({ accessToken: 'x'.repeat(10), expiresAt: Date.now() + 1000 }),
    ).not.toThrow();
    expect(spy).toHaveBeenCalled();
    expect(spy.mock.calls[0][0]).toContain('microsoft-teams-mcp');
  });
});
