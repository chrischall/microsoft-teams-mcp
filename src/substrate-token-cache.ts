import {
  createFileStatePersistence,
  resolveStateFile,
  type BearerTokens,
  type SyncStatePersistence,
} from '@chrischall/mcp-utils/session';
import { parseBoolEnv } from '@chrischall/mcp-utils';
import { reportCacheWriteFailure } from './token-cache.js';

/**
 * Where the captured Substrate search token is cached between runs.
 *
 * Its own file, beside `token.json` (chat service) and `sharepoint-tokens.json`:
 * a different audience with its own lifetime, and one token overall (Substrate
 * has a single host), so a plain `BearerTokens`, not a per-host map.
 */
export function substrateTokenCachePath(env: NodeJS.ProcessEnv = process.env): string {
  return resolveStateFile({
    env,
    envVar: 'TEAMS_SUBSTRATE_TOKEN_FILE',
    subdir: '.microsoft-teams-mcp',
    fileName: 'substrate-token.json',
  });
}

/** Same `BearerTokens` guard as the other caches; copied, as it is module-private there. */
function isToken(raw: unknown): raw is BearerTokens {
  if (raw === null || typeof raw !== 'object') return false;
  const t = raw as Partial<BearerTokens>;
  return (
    typeof t.accessToken === 'string' &&
    t.accessToken !== '' &&
    typeof t.expiresAt === 'number' &&
    (t.refreshToken === undefined || typeof t.refreshToken === 'string')
  );
}

/**
 * The Substrate token cache, or `null` when it is off or there is nothing to
 * cache. Gated like the SharePoint cache: `TEAMS_TOKEN_CACHE` (default on), and
 * off under `TEAMS_DISABLE_FETCHPROXY` — the bridge is the only source of a
 * Substrate token. A failed write is reported, never thrown.
 */
export function createSubstrateTokenCache(
  env: NodeJS.ProcessEnv = process.env,
): SyncStatePersistence<BearerTokens> | null {
  if (!parseBoolEnv('TEAMS_TOKEN_CACHE', { env, default: true })) return null;
  if (parseBoolEnv('TEAMS_DISABLE_FETCHPROXY', { env })) return null;

  const file = createFileStatePersistence<BearerTokens>({
    filePath: substrateTokenCachePath(env),
    boundTo: 'fetchproxy',
    validate: (raw) => (isToken(raw) ? raw : null),
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
