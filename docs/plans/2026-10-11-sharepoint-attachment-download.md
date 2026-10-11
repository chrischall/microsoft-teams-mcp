# SharePoint attachment download for `teams_get_attachment`

**Date:** 2026-10-11
**Status:** approved, ready to execute
**Follows:** `docs/plans/2026-10-10-chatsvc-api.md` (the 0.4 chat-service rewrite)

## Goal

Make `teams_get_attachment` actually return the **contents** of a shared file,
not just its name and open link. Inline images already download (Teams' AMS
object store); shared files live in SharePoint and were left as metadata with
the note "this MCP cannot download their contents yet."

## What the live spike proved (2026-10-11)

A 180s fetchproxy capture over a real Teams tenant, then four candidate routes
tried against every file found in chat. One route works end to end:

```
GET https://{host}/_api/v2.1/shares/u!{base64url(objectUrl)}/driveItem
    ?$select=id,size,@content.downloadUrl          [Authorization: Bearer <SP>]
GET {@content.downloadUrl}                          [no auth — pre-signed]
→ real bytes   (verified: 326KB and 165KB image/png came down intact)
```

Hard facts that shape the design:

- **The working token is a SharePoint token, not chatsvc and not Graph.**
  `aud=00000003-0000-0ff1-ce00-000000000000` (the SharePoint/Office 365 resource
  GUID), same Teams `appid=5e3ce6c0-2b1f-4285-8d4b-75ee78787346`, scopes
  `MyFiles.Write Sites.FullControl.All Sites.Manage.All User.ReadWrite.All`.
  Graph `/shares/{id}/driveItem` returned **401** every time (consistent with the
  earlier 403 on Graph `/shares`); chatsvc's token does not work on SharePoint.
- **It is short-lived: ~53 min TTL** (vs chatsvc's ~23h). It cannot be
  captured-once like chatsvc — it must refresh. `decodeJwtExp` gives the real
  expiry, so `TokenManager` can drive the refresh exactly as chatsvc's does.
- **It is host-family-scoped.** The `transunion-my.sharepoint.com` token also
  downloaded files on `transunion.sharepoint.com` (same tenant root), but got
  **403 on `transunionind-my.sharepoint.com`** (a separate OneDrive region host),
  and 403 on one `transunion-my` item (per-item permission). So each distinct SP
  host needs its own captured token, and a 403 is expected and must degrade
  gracefully, never crash the tool.
- **The capture layer cannot wildcard the host.** `validateCaptureHeaderDecls`
  runs each capture host through `HOSTNAME_RE` (`protocol/dist/validate.js`),
  which rejects `*.sharepoint.com`. The capture host must be a concrete hostname.
  We therefore derive the host from the file's own `objectUrl` at download time
  and declare a capture entry for exactly that host. The pairing **trust scope**
  (`domains`) may be the broad `sharepoint.com` — `hostMatchesAnyDomain` accepts
  any `*.sharepoint.com` subdomain via `endsWith('.' + domain)` — so one domain
  covers every tenant SP host without re-pairing per host.
- `sharepointIds` on each file carries `driveId/siteId/listId/listItemUniqueId/
  webId/siteUrl` — a documented fallback addressing path if the shares route ever
  404s (not required for the happy path; keep as optional, see Task 3).

## Decision (owner, 2026-10-11)

Approved mechanism: **SP token + v2.1 shares API + plain server-side fetch**
(the proven route), chosen over the browser `download` verb because it works
over the hosted/WAN bridge (plain fetch, no same-machine disk requirement) and
matches the existing chatsvc token model. Accepts a one-time re-pair to add
`sharepoint.com` to the extension trust scope, and a one-time "open the file
once in Teams" per SP host per ~50-min window to mint that host's token.

## Architecture fit

- `src/auth-fetchproxy.ts` — chatsvc capture. Reuse `stripBearer` and
  `getWsPort` from here; do not duplicate them.
- `src/chatsvc.ts` — `toBearerTokens` + `TokenManager` pattern to copy for the
  SP per-host token; `fetchBounded` for bounded byte reads; the `behavioroverride:
  redirectAs404` + `redirect: 'error'` discipline on the authed call.
- `src/token-cache.ts` — `createFileStatePersistence` pattern to copy for the
  per-host SP cache.
- `src/model.ts` — `FileAttachment` already carries `url` (the `objectUrl`) and
  `openUrl`; add optional `sharepointIds` (Task 3, fallback only).
- `src/deps.ts` / `src/index.ts` — add one long-lived `SharePointClient`,
  constructed in `index.ts` beside `ChatServiceClient` (never in a registrar).
- `src/tools/attachment.ts` — the only tool touched.

Environment knobs to mirror chatsvc's: `TEAMS_DISABLE_FETCHPROXY` (off ⇒ no SP
capture either), `TEAMS_TOKEN_CACHE` (gates the SP cache too). No new direct-token
env var — a SharePoint token is never supplied out of band in this MCP.

## Tasks

Each task is TDD: write failing tests first, then code, then green. Run
`npm run typecheck && vitest run` (i.e. `npm test`) — a green vitest alone is not
a green typecheck (CLAUDE.md). Keep every comment truthful to verified behaviour.

### Task 1 — `src/sharepoint-auth.ts`: capture a SharePoint bearer per host

Mirror `src/auth-fetchproxy.ts`. Export:

- `SHAREPOINT_TRUST_DOMAINS = ['sharepoint.com'] as const` — the trust scope.
- `isSharePointHost(host: string): boolean` — lowercased host matches
  `/^[a-z0-9][a-z0-9-]*\.sharepoint\.com$/` (so `transunion-my.sharepoint.com`,
  `transunion.sharepoint.com`, `transunionind-my.sharepoint.com` pass; reject
  anything not ending in `.sharepoint.com`, reject a bare `sharepoint.com`).
- `captureSharePointTokenViaFetchproxy(host: string, env = process.env):
  Promise<string>` — validate `host` with `isSharePointHost` (throw a clear
  `Error` otherwise); build the transport with
  `createBootstrapOpts({ domains: [...SHAREPOINT_TRUST_DOMAINS], bootstrap: {
  captureHeaders: [{ host, path: '/*', headerName: 'authorization' }] } })`,
  `port: getWsPort()` (imported from `auth-fetchproxy.js`), window from
  `captureWindowMs(env)` and `fetchTimeoutMs: bridgeDeadlineMs(windowMs)`,
  `serverName: PACKAGE_NAME`, `version: VERSION`. `start()`, then
  `captureRequestHeader({ host, path: '/*', headerName: 'authorization',
  timeoutMs: windowMs })`, `stripBearer` (import from `auth-fetchproxy.js`), the
  same `MIN_TOKEN_LENGTH` guard, return the bare token, `close()` in `finally`.
  On capture failure throw with guidance: "open the shared file once in Teams —
  that makes the browser request `{host}`, which is what the SharePoint token is
  captured from — then retry."

Tests (`tests/sharepoint-auth.test.ts`), mocking
`@chrischall/mcp-utils/fetchproxy` exactly as `tests/auth-fetchproxy.test.ts`
does: `isSharePointHost` accepts the three real hosts and rejects
`evil.com`, `sharepoint.com.evil.com`, `sharepoint.com`, `graph.microsoft.com`;
the capture decl passed to `captureRequestHeader` is `{host, path:'/*',
headerName:'authorization', timeoutMs}` for the host given; the transport's
`domains` is `['sharepoint.com']`; the `Bearer ` prefix is stripped; a too-short
header throws.

### Task 2 — `src/sharepoint-token-cache.ts`: a per-host SP token cache

SP tokens are per host and ~53 min, so cache a **map** keyed by host, not a
single token. Copy `src/token-cache.ts`'s use of `createFileStatePersistence` /
`resolveStateFile`.

- `sharePointTokenCachePath(env)` → `resolveStateFile({ env, envVar:
  'TEAMS_SHAREPOINT_TOKEN_FILE', subdir: '.microsoft-teams-mcp', fileName:
  'sharepoint-tokens.json' })`.
- `createSharePointTokenCache(env)` → a persistence over
  `Record<string, BearerTokens>`, gated by the same `TEAMS_TOKEN_CACHE`
  (default true) and returning `null` when `TEAMS_DISABLE_FETCHPROXY` is set
  (there is no direct SP token to bind to). `validate` must drop any entry whose
  shape is not a valid `BearerTokens` and return `null` if the whole blob is not
  a plain object.
- Provide a tiny per-host persistence **adapter** so each host gets a
  `SyncStatePersistence<BearerTokens>` view over its slice of the map
  (`load` reads `map[host]`, `save` merges, `clear` deletes the key), reusing the
  file-level persistence underneath. This is what each host's `TokenManager`
  binds to in Task 3. A failed write must not throw into the caller — wrap and
  report via a `reportCacheWriteFailure`-style stderr note (reuse the existing
  one if exported, else a local copy with the same message shape).

Tests (`tests/sharepoint-token-cache.test.ts`): path honours the env override;
disabled when `TEAMS_DISABLE_FETCHPROXY`; the per-host adapter round-trips one
host without clobbering another host's entry; a corrupt entry validates to
`null` without nuking the sibling.

### Task 3 — `src/sharepoint.ts`: the download client

`export class SharePointClient` built like `ChatServiceClient`:

- Constructor opts: `{ captureToken?: (host: string) => Promise<string>;
  fetchImpl?: typeof fetch; env?: NodeJS.ProcessEnv }`. Default `captureToken`
  lazily imports `./sharepoint-auth.js` and calls
  `captureSharePointTokenViaFetchproxy(host)`. Never throw from the constructor.
- A `Map<string, TokenManager>` — one `TokenManager` per host, built on first
  use with `initial`/`refresh` = `toBearerTokens(await captureToken(host))`
  (copy `toBearerTokens`/`decodeJwtExp` from chatsvc, or lift it to a shared
  helper), and `persistence` = Task 2's per-host adapter (omit when the cache is
  `null`). This gives the same 401 → re-capture-once → replay-once behaviour
  chatsvc gets, now per host.
- `async downloadFile(objectUrl: string, maxBytes: number): Promise<{ mimeType:
  string; data: Buffer; size: number }>`:
  1. Parse `objectUrl`; require `https:` and `isSharePointHost(host)` — else a
     `McpToolError` refusing (never send a bearer to an unvetted host).
  2. `share = 'u!' + base64url(objectUrl)` with **no padding**
     (`=`/`+`/`/` → ``/`-`/`_`, matching the spike).
  3. Authed GET
     `https://{host}/_api/v2.1/shares/{share}/driveItem?$select=id,size,@content.downloadUrl`
     via the host's `TokenManager.withAuth`, headers `authorization: Bearer`,
     `accept: application/json`, `behavioroverride: redirectAs404`,
     `redirect: 'error'`, bounded + timed like chatsvc's `#send`. Map **403 →**
     `McpToolError({ kind: 'forbidden', status: 403 })` with hint ("your account
     cannot open this file, or it is on a SharePoint host your captured token
     does not cover"); **404 → not found**; parse JSON otherwise.
  4. Read `@content.downloadUrl` (pre-signed). GET it with **no Authorization**,
     `accept: */*`, `redirect: 'follow'` (it 302s to a CDN/blob host — safe
     because no bearer is attached), bounded to `maxBytes`. Require `https:`.
  5. Return `{ mimeType: content-type sans params, data, size: data.length }`.
- Expose `isConfigured()` / a deferred-config pattern consistent with chatsvc so
  `TEAMS_DISABLE_FETCHPROXY` with no capture path yields a clean
  `no_credential` `McpToolError` rather than a bridge attempt.
- Optional (fallback): accept the file's `sharepointIds` and, only when the
  shares route 404s, try
  `https://{host}/_api/v2.1/sites/{siteId}/lists/{listId}/items/{listItemUniqueId}/driveItem`.
  Keep this behind a clear branch; the happy path does not need it. If it adds
  risk or time, omit it and leave a `// fallback: sharepointIds route` note — the
  shares route is the proven one.

Tests (`tests/sharepoint.test.ts`) with an injected `fetchImpl` + stub
`captureToken`: builds the exact shares URL incl. unpadded base64url of the
`objectUrl`; sends `Bearer` on the driveItem call and **no** auth header on the
downloadUrl call; returns `{mimeType,data,size}` from the second response;
`objectUrl` on a non-SharePoint host is refused before any fetch; a 403 on the
driveItem call surfaces as a `forbidden` `McpToolError`; `maxBytes` is passed to
the bounded read. Do not put real tokens or real tenant file names in fixtures.

### Task 4 — wire `SharePointClient` into `teams_get_attachment`

- `src/deps.ts`: add `sharepoint: SharePointClient` to `TeamsDeps`.
- `src/index.ts`: `const deps = { chatsvc: new ChatServiceClient(), sharepoint:
  new SharePointClient() }`; attachment registrar now takes
  `Pick<TeamsDeps, 'chatsvc' | 'sharepoint'>`.
- `src/tools/attachment.ts`: when a `file` is matched (and it has a `url`), try
  `sharepoint.downloadFile(file.url, MAX_FILE_BYTES)`:
  - **success:** if `mimeType` starts `image/`, return image content (as the
    inline-image branch does); otherwise return the file as MCP **embedded
    resource** content `{ type: 'resource', resource: { uri: file.openUrl ??
    file.url, mimeType, blob: data.toString('base64') } }`. Append the
    `untrustedResult` caption (name, type, mimeType, bytes, sender, time). If the
    server SDK rejects a `resource` block, fall back to returning the base64 in a
    clearly-labelled text block plus the caption — verify which the
    `@modelcontextprotocol/server` version in use accepts and comment the choice.
  - **graceful degrade:** on a capture failure, `no_credential`, a `forbidden`
    (403), or any `SharePointClient` error, return the **current** metadata +
    `openUrl` result (today's `untrustedResult({...file, content_available:false,
    hint})`) with a hint that says why (e.g. "couldn't mint a SharePoint token
    for `{host}` — open the file once in Teams and retry" / "your account cannot
    open this file"). The tool must never hard-fail just because a file could not
    be fetched; inline images and the no-match error path stay exactly as they
    are.
  - Add `MAX_FILE_BYTES` (e.g. `25 * 1024 * 1024`) near `MAX_RAW_MESSAGES`.
- Update the tool **description** and the module docblock: drop "cannot be
  downloaded by this MCP yet"; state that a shared file now returns its contents
  (image as image, else as an embedded resource), that the first file from a
  given SharePoint host in a session may require opening it once in Teams to mint
  a short-lived token, and that a file the account cannot open returns metadata +
  link. Update `FILE_HINT` accordingly.

Tests (extend `tests/` for the attachment tool, or add one): a stub `sharepoint`
whose `downloadFile` resolves bytes → tool returns resource/image content + the
untrusted caption; whose `downloadFile` rejects (forbidden / capture failure) →
tool returns the metadata-and-link fallback, not a thrown error; an inline image
still returns image content; an unknown `ref` still throws the "no attachment"
error. Keep stubs minimal (`Pick` the two deps).

### Task 5 — docs, trust scope, healthcheck, manifests

- `src/auth-fetchproxy.ts` / `README`: the extension trust scope now includes
  `sharepoint.com` for file download. Document the **one-time re-pair** and that
  it is additive to the chatsvc capture. If `TRUST_DOMAINS` is surfaced anywhere
  for pairing, make sure `sharepoint.com` is represented for the SP path (without
  widening the chatsvc capture decl itself).
- Healthcheck (`src/tools/healthcheck.ts`): if it enumerates capabilities or
  token sources, mention the SharePoint token is captured on demand per host and
  is short-lived; do not make a successful healthcheck depend on an SP token
  existing (there may be none until a file is opened).
- `docs/plans/2026-10-10-chatsvc-api.md` cross-link and the 0.4 output-changes
  doc: note that `teams_get_attachment` now downloads shared-file contents.
- Keep tool-count/manifest sync green (`tests/packaging.test.ts`,
  `tests/tool-annotations.test.ts`) — no new tool is added, so counts should be
  unchanged; fix them only if a manifest actually drifts.

## Release

A single user-facing feature: shared-file download. Title the squash
`feat: download shared-file attachment contents (SharePoint)` so release-please
cuts a minor (0.x `bump-minor-pre-major`). Does **not** block on the pending
0.4.0 release PR (#44) — this lands on top. Confirm `npm view
@chrischall/microsoft-teams-mcp version` after any release.

## Out of scope

- Office/PDF text extraction (return raw bytes; the host/model extracts).
- Uploading or sending files.
- The browser `download` verb path (rejected: same-machine only).
