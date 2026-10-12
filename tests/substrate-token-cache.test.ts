import { describe, expect, it } from 'vitest';
import { mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createSubstrateTokenCache, substrateTokenCachePath } from '../src/substrate-token-cache.js';

const dir = () => mkdtempSync(join(tmpdir(), 'teams-substrate-cache-'));

describe('substrate token cache', () => {
  it('resolves ~/.microsoft-teams-mcp/substrate-token.json by default', () => {
    const p = substrateTokenCachePath({ HOME: '/home/x' } as NodeJS.ProcessEnv);
    expect(p).toContain('.microsoft-teams-mcp');
    expect(p.endsWith('substrate-token.json')).toBe(true);
  });

  it('honours TEAMS_SUBSTRATE_TOKEN_FILE', () => {
    const p = join(dir(), 's.json');
    expect(substrateTokenCachePath({ TEAMS_SUBSTRATE_TOKEN_FILE: p } as NodeJS.ProcessEnv)).toBe(p);
  });

  it('is off when explicitly disabled', () => {
    expect(createSubstrateTokenCache({ TEAMS_TOKEN_CACHE: 'false' } as NodeJS.ProcessEnv)).toBeNull();
  });

  it('caches nothing when the bridge — the only Substrate token source — is disabled', () => {
    expect(
      createSubstrateTokenCache({
        TEAMS_DISABLE_FETCHPROXY: '1',
        TEAMS_SUBSTRATE_TOKEN_FILE: join(dir(), 's.json'),
      } as NodeJS.ProcessEnv),
    ).toBeNull();
  });

  it('round-trips a token, separately from the chat-service and SharePoint caches', () => {
    const file = join(dir(), 'substrate-token.json');
    const cache = createSubstrateTokenCache({ TEAMS_SUBSTRATE_TOKEN_FILE: file } as NodeJS.ProcessEnv);
    const tok = { accessToken: 'opaque', expiresAt: Date.now() + 1000 };
    cache!.save(tok);
    expect(cache!.load()).toEqual(tok);
    expect(readFileSync(file, 'utf8')).toContain('opaque');
    cache!.clear();
    expect(cache!.load()).toBeNull();
  });

  it('treats a malformed file as empty', () => {
    const file = join(dir(), 'substrate-token.json');
    const cache = createSubstrateTokenCache({ TEAMS_SUBSTRATE_TOKEN_FILE: file } as NodeJS.ProcessEnv);
    cache!.save({ accessToken: '', expiresAt: 1 } as never);
    expect(cache!.load()).toBeNull();
  });
});
