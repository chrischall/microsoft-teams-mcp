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
 * Where the captured SharePoint access tokens are cached between runs.
 *
 * A SEPARATE file from the chat-service `token.json`: a SharePoint token is a
 * different audience with a different lifetime. One token serves every
 * SharePoint host of the tenant (see {@link SharePointTokenCache.shared}), held
 * under the `*` key; older releases kept one entry per host, which `shared()`
 * still reads. Deliberately its own file for the same reason `tokenCachePath` is.
 */
export function sharePointTokenCachePath(env: NodeJS.ProcessEnv = process.env): string {
  return resolveStateFile({
    env,
    envVar: 'TEAMS_SHAREPOINT_TOKEN_FILE',
    subdir: '.microsoft-teams-mcp',
    fileName: 'sharepoint-tokens.json',
  });
}

type HostTokenMap = Record<string, BearerTokens>;

/**
 * Narrow the stored blob to a map of valid tokens.
 *
 * A non-object (or array) blob is `null` — the whole file is unusable. Otherwise
 * each host's slice is validated independently and a MALFORMED one is dropped
 * rather than discarding the file: a single corrupt entry must not cost every
 * other host its cached token.
 */
function validateMap(raw: unknown): HostTokenMap | null {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null;
  const out: HostTokenMap = {};
  for (const [host, value] of Object.entries(raw as Record<string, unknown>)) {
    if (isBearerTokens(value)) out[host] = value;
  }
  return out;
}

/** The key the one shared token is stored under. */
const SHARED_KEY = '*';

/** The SharePoint token cache. */
export interface SharePointTokenCache {
  /**
   * The one token every SharePoint host of the tenant accepts (verified live
   * 2026-11-09: a token captured from one OneDrive host returned 200 on another
   * OneDrive region host and on the team-site root), as the
   * `SyncStatePersistence<BearerTokens>` a `TokenManager` takes.
   *
   * `load` falls back to the latest-expiring per-host entry an older release
   * left, so an upgrade does not cost a capture. `save` writes the shared entry
   * alone, dropping those leftovers, and `clear` empties the file.
   */
  shared(): SyncStatePersistence<BearerTokens>;
}

/**
 * The SharePoint token cache, or `null` when it is off or there is nothing to
 * cache.
 *
 * Gated by the same `TEAMS_TOKEN_CACHE` (default on) as the chatsvc cache, and
 * `null` under `TEAMS_DISABLE_FETCHPROXY`: a SharePoint token is only ever
 * captured from the browser bridge — unlike chatsvc there is no direct-token
 * env var — so with the bridge off there is nothing to cache.
 */
export function createSharePointTokenCache(
  env: NodeJS.ProcessEnv = process.env,
): SharePointTokenCache | null {
  if (!parseBoolEnv('TEAMS_TOKEN_CACHE', { env, default: true })) return null;
  if (parseBoolEnv('TEAMS_DISABLE_FETCHPROXY', { env })) return null;

  // One file-backed persistence for the map; the shared view below reads and
  // writes it. No `boundTo`: there is no single credential the whole map is
  // minted from — each host's token stands or falls on its own shape.
  const file = createFileStatePersistence<HostTokenMap>({
    filePath: sharePointTokenCachePath(env),
    validate: validateMap,
  });

  return {
    shared(): SyncStatePersistence<BearerTokens> {
      return {
        load() {
          const map = file.load();
          if (!map) return null;
          if (map[SHARED_KEY]) return map[SHARED_KEY];
          const legacy = Object.values(map).sort((a, b) => (b.expiresAt ?? 0) - (a.expiresAt ?? 0));
          return legacy[0] ?? null;
        },
        save(tokens: BearerTokens) {
          // A failed write is reported, never thrown: a SharePoint token is
          // cheap to re-capture, and a read-only data dir must not take the
          // tool down — it just loses the between-run cache.
          try {
            file.save({ [SHARED_KEY]: tokens });
          } catch (err) {
            reportCacheWriteFailure(err);
          }
        },
        clear() {
          try {
            if (file.load()) file.save({});
          } catch (err) {
            reportCacheWriteFailure(err);
          }
        },
      };
    },
  };
}
