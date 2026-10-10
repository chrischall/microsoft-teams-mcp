import { describe, expect, it, vi, afterEach } from 'vitest';
import { mkdtempSync, existsSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createTokenCache, tokenCachePath, reportCacheWriteFailure } from '../src/token-cache.js';

const dir = () => mkdtempSync(join(tmpdir(), 'teams-cache-'));

afterEach(() => vi.restoreAllMocks());

describe('token cache', () => {
  it('resolves ~/.microsoft-teams-mcp/token.json by default', () => {
    const p = tokenCachePath({ HOME: '/home/x' } as NodeJS.ProcessEnv);
    expect(p).toContain('.microsoft-teams-mcp');
    expect(p.endsWith('token.json')).toBe(true);
  });

  it('honours TEAMS_TOKEN_FILE', () => {
    const p = join(dir(), 'custom.json');
    expect(tokenCachePath({ TEAMS_TOKEN_FILE: p } as NodeJS.ProcessEnv)).toBe(p);
  });

  it('is off when explicitly disabled', () => {
    expect(createTokenCache({ TEAMS_TOKEN_CACHE: 'false' } as NodeJS.ProcessEnv)).toBeNull();
  });

  it('caches nothing when the token IS the environment variable', () => {
    expect(
      createTokenCache({
        TEAMS_ACCESS_TOKEN: 'tok',
        TEAMS_TOKEN_FILE: join(dir(), 't.json'),
      } as NodeJS.ProcessEnv),
    ).toBeNull();
  });

  it('caches nothing when the bridge is the only source and it is disabled', () => {
    expect(
      createTokenCache({
        TEAMS_DISABLE_FETCHPROXY: '1',
        TEAMS_TOKEN_FILE: join(dir(), 't.json'),
      } as NodeJS.ProcessEnv),
    ).toBeNull();
  });

  it('round-trips a captured token', () => {
    const file = join(dir(), 'token.json');
    const cache = createTokenCache({ TEAMS_TOKEN_FILE: file } as NodeJS.ProcessEnv);
    expect(cache).not.toBeNull();
    const tokens = { accessToken: 'a', expiresAt: Date.now() + 1000 };
    cache!.save(tokens);
    expect(existsSync(file)).toBe(true);
    expect(readFileSync(file, 'utf8')).toContain('accessToken');
    expect(cache!.load()).toMatchObject({ accessToken: 'a' });
    cache!.clear();
    expect(cache!.load()).toBeNull();
  });

  it('rejects a malformed record', () => {
    const file = join(dir(), 'token.json');
    const cache = createTokenCache({ TEAMS_TOKEN_FILE: file } as NodeJS.ProcessEnv)!;
    cache.save({ accessToken: '', expiresAt: Date.now() + 1000 });
    expect(cache.load()).toBeNull();
    cache.save({ accessToken: 'a', expiresAt: 'soon' as unknown as number });
    expect(cache.load()).toBeNull();
  });

  it('reports a failed cache write without throwing', () => {
    const spy = vi.spyOn(console, 'error').mockImplementation(() => {});
    reportCacheWriteFailure(new Error('read-only fs'));
    reportCacheWriteFailure('not an error');
    expect(spy).toHaveBeenCalledTimes(2);
    expect(spy.mock.calls[0][0]).toContain('read-only fs');
    expect(spy.mock.calls[0][0]).toContain('microsoft-teams-mcp');
  });
});
