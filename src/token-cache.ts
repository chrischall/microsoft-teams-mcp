import {
  createFileStatePersistence,
  isBearerTokens,
  resolveStateFile,
  type BearerTokens,
  type SyncStatePersistence,
} from '@chrischall/mcp-utils/session';
import { readEnvVar, parseBoolEnv } from '@chrischall/mcp-utils';

/**
 * Where the captured chat-service access token is cached between runs.
 *
 * Deliberately its OWN directory, so nothing else that keeps Teams
 * credentials (the `teams-fpx` access skill, say) can corrupt it by writing a
 * different format to the same file.
 */
export function tokenCachePath(env: NodeJS.ProcessEnv = process.env): string {
  return resolveStateFile({
    env,
    envVar: 'TEAMS_TOKEN_FILE',
    subdir: '.microsoft-teams-mcp',
    fileName: 'token.json',
  });
}

/**
 * What a cached token is bound to, or `null` when there is nothing worth
 * caching.
 *
 * A `TEAMS_ACCESS_TOKEN` IS the environment variable — there is no capture to
 * skip, so caching it would only copy a credential onto disk. A captured token
 * is worth caching most of all: it lets a cold start proceed with no browser
 * present for the token's ~23h lifetime.
 */
function bindingFor(env: NodeJS.ProcessEnv): string | null {
  if (readEnvVar('TEAMS_ACCESS_TOKEN', { env }) !== undefined) return null;
  if (parseBoolEnv('TEAMS_DISABLE_FETCHPROXY', { env })) return null;
  return 'fetchproxy';
}

/** The token cache, or `null` when it is off or there is nothing to cache. */
export function createTokenCache(
  env: NodeJS.ProcessEnv = process.env,
): SyncStatePersistence<BearerTokens> | null {
  if (!parseBoolEnv('TEAMS_TOKEN_CACHE', { env, default: true })) return null;
  const boundTo = bindingFor(env);
  if (boundTo === null) return null;

  return createFileStatePersistence<BearerTokens>({
    filePath: tokenCachePath(env),
    boundTo,
    validate: (raw) => (isBearerTokens(raw) ? raw : null),
  });
}

/**
 * Report a cache write that failed. Not fatal — the token is re-capturable —
 * but worth saying, because a read-only data dir otherwise looks exactly like a
 * server that simply never caches.
 *
 * stderr only; stdout is the JSON-RPC channel.
 */
export function reportCacheWriteFailure(err: unknown): void {
  const detail = err instanceof Error ? err.message : String(err);
  console.error(
    `[microsoft-teams-mcp] could not cache the access token (${detail}); continuing ` +
      'without the cache — every restart will re-capture from the browser.',
  );
}
