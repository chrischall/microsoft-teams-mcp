import {
  createFileStatePersistence,
  isBearerTokens,
  resolveStateFile,
  type BearerTokens,
  type SyncStatePersistence,
} from '@chrischall/mcp-utils/session';
import { parseBoolEnv } from '@chrischall/mcp-utils';
import { reportCacheWriteFailure } from './token-cache.js';

/**
 * Where the captured meeting-recap (MCPS) token is cached between runs: its
 * own file beside the other three, since it has its own audience and lifetime.
 */
export function recapTokenCachePath(env: NodeJS.ProcessEnv = process.env): string {
  return resolveStateFile({
    env,
    envVar: 'TEAMS_RECAP_TOKEN_FILE',
    subdir: '.microsoft-teams-mcp',
    fileName: 'recap-token.json',
  });
}

/**
 * The recap token cache, or `null` when it is off (`TEAMS_TOKEN_CACHE=false`)
 * or there is nothing to cache (`TEAMS_DISABLE_FETCHPROXY` — the bridge is
 * its only source). A failed write is reported, never thrown.
 */
export function createRecapTokenCache(env: NodeJS.ProcessEnv = process.env): SyncStatePersistence<BearerTokens> | null {
  if (!parseBoolEnv('TEAMS_TOKEN_CACHE', { env, default: true })) return null;
  if (parseBoolEnv('TEAMS_DISABLE_FETCHPROXY', { env })) return null;

  const file = createFileStatePersistence<BearerTokens>({
    filePath: recapTokenCachePath(env),
    boundTo: 'fetchproxy',
    validate: (raw) => (isBearerTokens(raw) ? raw : null),
  });
  return {
    load: () => file.load(),
    save(tokens) {
      try {
        file.save(tokens);
      } catch (err) {
        reportCacheWriteFailure(err);
      }
    },
    clear() {
      try {
        file.clear();
      } catch (err) {
        reportCacheWriteFailure(err);
      }
    },
  };
}
