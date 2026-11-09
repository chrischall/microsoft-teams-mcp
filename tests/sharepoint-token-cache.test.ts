import { describe, expect, it, vi, afterEach } from 'vitest';
import { mkdtempSync, existsSync, writeFileSync, mkdirSync, readFileSync } from 'node:fs';
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

  // One SharePoint token is accepted on every host of the tenant (verified live
  // 2026-11-09: a token captured from one OneDrive host returned 200 on the
  // other OneDrive region and on the team-site root), so there is one entry.
  it('round-trips the one shared token', () => {
    const file = join(dir(), 'sharepoint-tokens.json');
    const cache = createSharePointTokenCache({ TEAMS_SHAREPOINT_TOKEN_FILE: file } as NodeJS.ProcessEnv);
    expect(cache).not.toBeNull();

    const shared = cache!.shared();
    expect(shared.load()).toBeNull();
    shared.save({ accessToken: 'aaaa', expiresAt: Date.now() + 1000 });
    expect(existsSync(file)).toBe(true);
    expect(shared.load()).toMatchObject({ accessToken: 'aaaa' });
    shared.clear();
    expect(shared.load()).toBeNull();
  });

  it('reuses the newest token an older per-host cache left behind, so an upgrade does not re-capture', () => {
    const file = join(dir(), 'sharepoint-tokens.json');
    writeFileSync(
      file,
      JSON.stringify({
        v: 1,
        state: {
          [HOST_A]: { accessToken: 'older', expiresAt: Date.now() + 1000 },
          [HOST_B]: { accessToken: 'newer', expiresAt: Date.now() + 5000 },
        },
      }),
    );
    const shared = createSharePointTokenCache({ TEAMS_SHAREPOINT_TOKEN_FILE: file } as NodeJS.ProcessEnv)!.shared();
    expect(shared.load()).toMatchObject({ accessToken: 'newer' });
  });

  it('replaces the per-host entries when it saves, and clears them all', () => {
    const file = join(dir(), 'sharepoint-tokens.json');
    writeFileSync(
      file,
      JSON.stringify({ v: 1, state: { [HOST_A]: { accessToken: 'legacy', expiresAt: Date.now() + 9000 } } }),
    );
    const shared = createSharePointTokenCache({ TEAMS_SHAREPOINT_TOKEN_FILE: file } as NodeJS.ProcessEnv)!.shared();
    shared.save({ accessToken: 'fresh', expiresAt: Date.now() + 1000 });
    expect(JSON.parse(readFileSync(file, 'utf8')).state).toEqual({
      '*': { accessToken: 'fresh', expiresAt: expect.any(Number) },
    });
    shared.clear();
    expect(shared.load()).toBeNull();
  });

  it('skips a corrupt entry and uses a good one', () => {
    const file = join(dir(), 'sharepoint-tokens.json');
    writeFileSync(
      file,
      JSON.stringify({
        v: 1,
        state: {
          [HOST_A]: { accessToken: '', expiresAt: 1 }, // empty token → malformed
          [HOST_B]: { accessToken: 'good', expiresAt: Date.now() + 1000 },
        },
      }),
    );
    const cache = createSharePointTokenCache({ TEAMS_SHAREPOINT_TOKEN_FILE: file } as NodeJS.ProcessEnv)!;
    expect(cache.shared().load()).toMatchObject({ accessToken: 'good' });
  });

  it('treats a non-object blob as empty rather than crashing', () => {
    const file = join(dir(), 'sharepoint-tokens.json');
    writeFileSync(file, JSON.stringify(['not', 'a', 'map']));
    const cache = createSharePointTokenCache({ TEAMS_SHAREPOINT_TOKEN_FILE: file } as NodeJS.ProcessEnv)!;
    expect(cache.shared().load()).toBeNull();
  });

  it('reports a failed write without throwing into the caller', () => {
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    const file = join(dir(), 'sharepoint-tokens.json');
    const cache = createSharePointTokenCache({
      TEAMS_SHAREPOINT_TOKEN_FILE: file,
    } as NodeJS.ProcessEnv)!;
    const host = cache.shared();
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
