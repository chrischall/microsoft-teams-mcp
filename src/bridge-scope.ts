/**
 * The ONE scope every browser-bridge capture of this server declares.
 *
 * The server captures three tokens through the ContextMint Bridge — the chat
 * service's (`auth-fetchproxy.ts`), Substrate search's (`substrate-auth.ts`) and
 * a SharePoint host's (`sharepoint-auth.ts`) — all under one identity. The
 * bridge keys its trust record on that identity and:
 *
 *   - forces a fresh pair whenever a hello's DOMAIN SET differs from the
 *     approved one (`sameDomainSet` in contextmint-bridge's `background/hello.ts`);
 *   - offers a "grant" whenever the declared CAPTURE HEADERS exceed the approved
 *     ones, and approving it REPLACES the stored list with the declared one.
 *
 * So captures that each declared only their own host thrashed the record: a
 * chat-service capture after a search re-paired, and so on — a pair prompt or a
 * grant prompt on nearly every use. Every capture therefore declares the same
 * domains and the same capture-header list, and captures only the one header it
 * needs.
 *
 * The SharePoint capture host is concrete and per-tenant (the capture layer
 * rejects a wildcard host), so it cannot be known up front. Each one is
 * remembered in `bridge-scope.json` beside the SharePoint token cache once a
 * capture of it SUCCEEDS — i.e. once the user has granted it — and every later
 * hello declares it. A host is never remembered before that: a capture the
 * user has not approved would otherwise grow every later hello and keep
 * offering the same grant.
 */
import { dirname, join } from 'node:path';
import { parseBoolEnv } from '@chrischall/mcp-utils';
import { createFileStatePersistence } from '@chrischall/mcp-utils/session';
import { sharePointTokenCachePath } from './sharepoint-token-cache.js';
import { reportCacheWriteFailure } from './token-cache.js';

/** One header capture, as the bridge hello declares it. */
export interface CaptureDecl {
  host: string;
  path: string;
  headerName: string;
}

/**
 * The chat-service capture: any chat-service call on the New Teams host. An
 * idle tab may make none — opening a chat is what makes the worker issue one.
 */
export const CAPTURE_DECL = {
  host: 'teams.cloud.microsoft',
  path: '/api/chatsvc/*',
  headerName: 'authorization',
} as const;

/** The search capture: a search-service call the Teams tab makes. */
export const SUBSTRATE_CAPTURE_DECL = {
  host: 'substrate.office.com',
  path: '/searchservice/*',
  headerName: 'authorization',
} as const;

/**
 * Every domain the server captures from. `sharepoint.com` is the family domain:
 * the bridge accepts any `*.sharepoint.com` host under it.
 */
export const BRIDGE_TRUST_DOMAINS = [CAPTURE_DECL.host, 'sharepoint.com', SUBSTRATE_CAPTURE_DECL.host] as const;

/**
 * True for a concrete SharePoint tenant host (`{label}.sharepoint.com`), such
 * as `contoso-my.sharepoint.com` or `contoso.sharepoint.com`. Rejects a bare
 * `sharepoint.com`, a suffix spoof (`sharepoint.com.evil.com`), and any
 * non-SharePoint host. Host is lowercased before matching.
 */
export function isSharePointHost(host: string): boolean {
  return /^[a-z0-9][a-z0-9-]*\.sharepoint\.com$/.test(host.toLowerCase());
}

/** The capture of a SharePoint token off any browser request to `host`. */
export function sharePointCaptureDecl(host: string): CaptureDecl {
  return { host: host.toLowerCase(), path: '/*', headerName: 'authorization' };
}

/**
 * The capture headers every hello declares: the chat-service and search
 * captures, then one per granted SharePoint host (deduplicated, sorted, and only
 * concrete SharePoint hosts), so the list is the same whichever capture runs.
 */
export function bridgeCaptureDecls(sharePointHosts: Iterable<string>): CaptureDecl[] {
  const hosts = [...new Set([...sharePointHosts].map((h) => h.toLowerCase()).filter(isSharePointHost))].sort();
  return [{ ...CAPTURE_DECL }, { ...SUBSTRATE_CAPTURE_DECL }, ...hosts.map(sharePointCaptureDecl)];
}

/** Where the remembered SharePoint hosts are kept: beside the SharePoint token cache. */
export function bridgeScopePath(env: NodeJS.ProcessEnv = process.env): string {
  return join(dirname(sharePointTokenCachePath(env)), 'bridge-scope.json');
}

interface StoredScope {
  sharePointHosts: string[];
}

function validateScope(raw: unknown): StoredScope | null {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null;
  const hosts = (raw as { sharePointHosts?: unknown }).sharePointHosts;
  if (!Array.isArray(hosts)) return null;
  return { sharePointHosts: hosts.filter((h): h is string => typeof h === 'string' && isSharePointHost(h)) };
}

function scopeFile(env: NodeJS.ProcessEnv) {
  return createFileStatePersistence<StoredScope>({ filePath: bridgeScopePath(env), validate: validateScope });
}

/** The SharePoint hosts a capture has succeeded for; empty when none, or the file is unusable. */
export function knownSharePointHosts(env: NodeJS.ProcessEnv = process.env): string[] {
  try {
    return scopeFile(env).load()?.sharePointHosts.map((h) => h.toLowerCase()) ?? [];
  } catch {
    return [];
  }
}

/**
 * Remember `host` once a capture of it has succeeded. Off with
 * `TEAMS_TOKEN_CACHE=false` (nothing of this server's is written to disk then);
 * a failed write is reported, never thrown — it costs at most one more grant
 * prompt.
 */
export function rememberSharePointHost(host: string, env: NodeJS.ProcessEnv = process.env): void {
  if (!isSharePointHost(host)) return;
  if (!parseBoolEnv('TEAMS_TOKEN_CACHE', { env, default: true })) return;
  const known = knownSharePointHosts(env);
  const h = host.toLowerCase();
  if (known.includes(h)) return;
  try {
    scopeFile(env).save({ sharePointHosts: [...known, h].sort() });
  } catch (err) {
    reportCacheWriteFailure(err);
  }
}

/**
 * The hello options of a capture transport, widened to the one scope: the
 * domains and capture headers mcp-utils derived from the single capture are
 * replaced with {@link BRIDGE_TRUST_DOMAINS} and {@link bridgeCaptureDecls}
 * (plus `extraSharePointHosts`, the hosts a SharePoint capture is about to try).
 * The verb set is unchanged — every capture is `capture_request_header`.
 */
export function withBridgeScope<T extends object>(
  opts: T,
  env: NodeJS.ProcessEnv,
  extraSharePointHosts: readonly string[] = [],
): T & { domains: string[]; tabDomains: string[]; captureHeaders: CaptureDecl[] } {
  const hosts = [...knownSharePointHosts(env), ...extraSharePointHosts];
  return {
    ...opts,
    domains: [...BRIDGE_TRUST_DOMAINS],
    // Only Teams needs a page: the SharePoint and search captures read requests
    // the Teams tab makes, so a tab on those domains would do nothing.
    tabDomains: [CAPTURE_DECL.host],
    captureHeaders: bridgeCaptureDecls(hosts),
  };
}
