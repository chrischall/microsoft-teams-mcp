import {
  createFileStatePersistence,
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
 * different audience with a very different lifetime (~53 min vs chatsvc's ~23h),
 * and there is one per SharePoint host rather than one overall — so this file
 * holds a `Record<host, BearerTokens>`, not a single token. Deliberately its own
 * directory for the same reason `tokenCachePath` is.
 */
export function sharePointTokenCachePath(env: NodeJS.ProcessEnv = process.env): string {
  return resolveStateFile({
    env,
    envVar: 'TEAMS_SHAREPOINT_TOKEN_FILE',
    subdir: '.microsoft-teams-mcp',
    fileName: 'sharepoint-tokens.json',
  });
}

/** Same `BearerTokens` guard as the chatsvc cache; copied, as it is module-private there. */
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
    if (isToken(value)) out[host] = value;
  }
  return out;
}

/** A per-host cache: hand each SharePoint host its own single-token view. */
export interface SharePointTokenCache {
  /**
   * A `SyncStatePersistence<BearerTokens>` over one host's slice of the shared
   * map — the exact shape a `TokenManager` takes, so each host gets its own
   * manager over one file. `save` merges (it never clobbers a sibling host) and
   * `clear` deletes just this host's key.
   */
  forHost(host: string): SyncStatePersistence<BearerTokens>;
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

  // One file-backed persistence for the whole map; the per-host views below read
  // and write it. No `boundTo`: there is no single credential the whole map is
  // minted from — each host's token stands or falls on its own shape.
  const file = createFileStatePersistence<HostTokenMap>({
    filePath: sharePointTokenCachePath(env),
    validate: validateMap,
  });

  return {
    forHost(host: string): SyncStatePersistence<BearerTokens> {
      return {
        load() {
          return file.load()?.[host] ?? null;
        },
        save(tokens: BearerTokens) {
          // Read-modify-write the shared map so a sibling host's token survives.
          // A failed write is reported, never thrown: a SharePoint token is
          // cheap to re-capture, and a read-only data dir must not take the
          // tool down — it just loses the between-run cache.
          try {
            const map = file.load() ?? {};
            map[host] = tokens;
            file.save(map);
          } catch (err) {
            reportCacheWriteFailure(err);
          }
        },
        clear() {
          try {
            const map = file.load();
            if (map && host in map) {
              delete map[host];
              file.save(map);
            }
          } catch (err) {
            reportCacheWriteFailure(err);
          }
        },
      };
    },
  };
}
