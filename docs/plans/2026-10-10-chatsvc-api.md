# Teams via the chat-service API — implementation plan

**Goal.** Let the MCP read ANY chat or channel by name/id (no "open it in the
browser first"), return exactly the messages the user has not read yet, and
read attachments — by moving the chat/channel tools off DOM-scraping onto
Teams' own chat-service REST API, authenticated by a bearer token captured once
from the signed-in tab (the `office-outlook-mcp` bootstrap-capture archetype).

**Architecture.** `auth-fetchproxy.ts` captures `Authorization: Bearer …` off
a request the Teams worker already makes to
`https://teams.cloud.microsoft/api/chatsvc/*` (fetchproxy
`capture_request_header`). From then on every call is a plain Node `fetch` —
the bridge is touched once per token (~23h), never in the hot path. A 401
triggers one re-capture and replay. `teams_get_activity` stays DOM-based (no
API equivalent verified) and keeps the `read_dom_list` transport.

**Tech.** TypeScript (ESM), `@chrischall/mcp-utils` (`TokenManager`,
`createApiClient`-style patterns, `untrustedResult`, `McpToolError`,
`toolAnnotations`), `@chrischall/mcp-utils/fetchproxy`, zod 4, vitest. TDD:
every task writes failing tests first. `npm test` runs typecheck + vitest and
must be green at the end of every task (a green vitest is not a green
typecheck — `npm test` runs both).

## Verified live facts (2026-10-10, real tenant) — build to these

- **Auth.** Worker requests to `teams.cloud.microsoft/api/chatsvc/*` carry
  `Authorization: Bearer <JWT>`; `aud=https://ic3.teams.office.com`,
  `scp=Teams.AccessAsUser.All`, lifetime ~23.5h (`exp` claim). A plain
  server-side request with that header + `behavioroverride: redirectAs404`
  returns 200. No `authentication`/`x-skypetoken` header is sent.
  The JWT's `oid` claim is the signed-in user; their MRI is `8:orgid:<oid>`.
- **Base URL.** `https://teams.cloud.microsoft/api/chatsvc/{region}/v1`, region
  `amer` for this tenant. Make it `TEAMS_CHATSVC_REGION` (default `amer`) and
  `TEAMS_CHATSVC_BASE` (full override).
- **List conversations.**
  `GET {base}/users/ME/conversations?view=msnp24Equivalent&pageSize=200` →
  `{ conversations: [...], _metadata: { backwardLink, forwardLink, syncState, totalCount } }`.
  Follow `_metadata.backwardLink` (absolute URL) for more pages until absent or
  a page is empty. Conversation keys: `id, lastMessage, memberProperties,
  properties, threadProperties, type, version, targetLink, messages, …`.
  - `properties`: `consumptionhorizon, favorite, isemptyconversation,
    lastimreceivedtime, lastimportantimreceivedtime, addedBy, …`.
  - `threadProperties`: `topic, threadType, productThreadType, spaceId,
    groupId, createdat, lastjoinat, privacy, …` (all strings).
  - `lastMessage`: `id, composetime, originalarrivaltime, content, from,
    imdisplayname, messagetype, conversationid, sequenceId, …`.
  - Id suffixes seen: `@thread.v2` (group/meeting chats, 145),
    `@thread.tacv2` (channels, 30), `@unq.gbl.spaces` (1:1, 15), and `48:*`
    system ids (notes/calendar/etc — exclude from results).
- **Read marker.** `properties.consumptionhorizon` =
  `"<lastReadArrivalMs>;<clientMs>;<messageId>"` (seen from Teams' own
  `PUT …/properties?name=consumptionhorizon`). A message is unread when its
  `id` (numeric string = arrival ms) > the first segment and it is not from
  the user. 19 of 194 conversations were unread by `lastMessage.id >
  horizon[0]` on the live tenant.
- **Messages.**
  `GET {base}/users/ME/conversations/{encodeURIComponent(id)}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=200&startTime=1`
  → `{ messages: [...] (newest first), _metadata: { backwardLink,
  lastCompleteSegmentStartTime, lastCompleteSegmentEndTime, syncState }, tenantId }`.
  Follow `backwardLink` for older pages. Message keys: `id, clientmessageid,
  composetime, originalarrivaltime, content, contenttype, messagetype, from,
  imdisplayname, conversationLink, conversationid, properties, sequenceId,
  version, type, …`. `messagetype` seen: `Text`, `RichText/Html`; others
  (`ThreadActivity/*`, `Event/Call`, …) are system events. `from` is a URL
  ending in the sender MRI. In a channel, a reply's `conversationLink` ends
  with `;messageid=<rootId>`; a root post's root id is its own `id`.
  `properties` keys seen: `files, subject, mentions, emotions, edittime,
  deletetime, cards, links, qtdMsgs, atp, …`. `deletetime` set ⇒ deleted.
- **File attachments.** `properties.files` is a JSON STRING (parse it) of
  objects with `fileName, fileType, title, objectUrl, openUrl, baseUrl, id,
  itemid, type, state, fileInfo{fileUrl,itemId,serverRelativeUrl,shareId,shareUrl,siteUrl},
  sharepointIds, site, …`. They live in SharePoint (several tenant hosts).
  The chatsvc token CANNOT download them (Graph /shares → 403, SharePoint
  direct → 401). **Scope for this plan: return file metadata + `openUrl`/
  `objectUrl`, never content.** Content download is a follow-up spike.
- **Inline images.** `<img src="https://us-api.asm.skype.com/v1/objects/{id}/views/imgo">`
  in `content`. `GET` with the chatsvc bearer → 200 `image/jpeg` (8–18 KB
  seen). Host varies by region (`us-api.asm.skype.com`); only fetch URLs whose
  host ends in `.asm.skype.com` (never send the token anywhere else).

## Tool surface after this plan

| Tool | Change |
|---|---|
| `teams_list_chats` | **Reimplemented on the API** — every chat (group/meeting/1:1), no browser view needed. Adds `unread_only`, `query`. |
| `teams_list_teams_and_channels` | **Reimplemented on the API** — every channel the user is in, grouped by team where determinable. Adds `unread_only`, `query`. |
| `teams_read_conversation` | **New.** Read any chat or channel by id or name; `unread_only`, `since`, `limit`. Channel replies grouped under their root post. |
| `teams_get_unread` | **New.** Every unread message across all chats/channels (or a filter), grouped by conversation. Never marks anything read. |
| `teams_get_attachment` | **New.** Inline image → MCP image content; file → metadata + open link (content not fetchable yet, said explicitly). |
| `teams_get_open_chat_messages`, `teams_get_open_channel_posts` | **Removed** — superseded by `teams_read_conversation`. |
| `teams_get_activity` | Unchanged (DOM). |
| `teams_healthcheck` | Probes the API (token capture + one conversations page) instead of the DOM chat list. |

All Teams-text results stay wrapped in `untrustedResult` with
`UNTRUSTED_DESCRIPTION_SUFFIX` on descriptions. All tools `readOnly: true,
openWorld: true`. No tool writes the read marker.

---

## Task 1 — token capture + cache

**Files:** create `src/auth-fetchproxy.ts`, `src/token-cache.ts`,
`tests/auth-fetchproxy.test.ts`, `tests/token-cache.test.ts`.

Port `~/git/office-outlook-mcp/src/auth-fetchproxy.ts` and
`src/token-cache.ts` (read both first; match their style and comment density).
Differences:

- One capture declaration: `{ host: 'teams.cloud.microsoft', path: '/api/chatsvc/*', headerName: 'authorization' }`;
  `TRUST_DOMAINS = ['teams.cloud.microsoft']`. No race across hosts needed —
  keep a single-attempt version (do not port `raceCaptures`).
- Env: `TEAMS_CAPTURE_TIMEOUT` (default 30s), `TEAMS_WS_PORT` (reuse existing
  `getWsPort`), cache file `~/.microsoft-teams-mcp/token.json` via
  `resolveStateFile` with `TEAMS_TOKEN_FILE` override, `TEAMS_ACCESS_TOKEN`
  direct-token override, `TEAMS_DISABLE_FETCHPROXY`.
- Error text on capture failure: tell the user to open a signed-in
  teams.cloud.microsoft tab and click into any chat (that is what makes the
  worker issue a chatsvc request — an idle tab never matches).
- Lazy import of the bridge module exactly as Outlook does (the `.mcpb`
  bundle must boot without `@fetchproxy/server` installed).

Tests: header normalisation (`stripBearer`), too-short token rejected, cache
round-trip, cache ignored when the env token is set, capture uses the declared
host/path/header (mock `createFetchproxyTransport` / inject a fake capture).

## Task 2 — `ChatServiceClient`

**Files:** create `src/chatsvc.ts`, `tests/chatsvc.test.ts`,
`tests/fixtures/` (hand-written JSON fixtures using ONLY the key shapes in
"Verified live facts" — invented names/text, never real data).

- Constructor `{ env?, captureToken?, fetchImpl? }` (inject both for tests;
  receiver-safe `fetch` wrapper as in Outlook's client).
- Token via `TokenManager` with the recapture sentinel pattern from Outlook's
  `client.ts`; expiry from the JWT `exp`. `me` MRI from the JWT `oid`
  (`8:orgid:<oid>`); expose `meMri()`.
- Every request sends `authorization: Bearer …`, `behavioroverride:
  redirectAs404`, `accept: application/json`. On 401: invalidate, re-capture
  once, replay once; a second 401 → `McpToolError` with a re-open-Teams hint.
  429/5xx → `McpToolError` carrying status (no retry loop).
- `listConversations(): Promise<RawConversation[]>` — follows `backwardLink`
  pages (cap 10 pages), drops `48:` ids.
- `getMessages(id, { stopAt?: (m) => boolean, max }): Promise<RawMessage[]>`
  — newest-first pages via `backwardLink` until `stopAt` returns true for a
  message, `max` reached, or no more pages (cap 20 pages). Returns
  newest-first.
- `fetchImage(url): Promise<{ mimeType, data: Buffer }>` — rejects (throws
  `McpToolError`) unless `new URL(url).hostname` ends with `.asm.skype.com`
  and protocol is https; size cap 5 MB.
- Only `listConversations`/`getMessages`/`fetchImage` — no write methods.

Tests: paging stops on each condition; 401 → one recapture → success; second
401 → error; host allow-list for images (reject `evil.com`,
`asm.skype.com.evil.com`, `http://`); `48:` ids dropped; headers sent.

## Task 3 — normalisation (pure functions)

**Files:** create `src/model.ts`, `tests/model.test.ts`.

- `conversationKind(c)`: `'channel'` (`@thread.tacv2`), `'meeting'`
  (`threadProperties.productThreadType`/`threadType` says meeting, or id starts
  `19:meeting_`), `'oneOnOne'` (`@unq.gbl.spaces`), else `'group'`.
- `readHorizon(c)`: first `;` segment of `consumptionhorizon` as a number, 0
  when absent/unparseable.
- `isUnread(msg, horizon, meMri)`: numeric `id` > horizon, sender MRI ≠ me,
  `messagetype` is `Text` or `RichText/Html`, not deleted.
- `conversationSummary(c, meMri)` → `{ id, kind, title, team?, unread,
  lastActivity (ISO), lastMessagePreview, lastSender }`. Title: `topic`;
  else for 1:1/group: other participant's `imdisplayname` from `lastMessage`
  when it is not from me, else `"(untitled chat)"`. Channel `title` = topic or
  `"General"`; `team` = topic of the conversation whose id equals
  `threadProperties.spaceId` when present in the list (pass the full list in).
- `htmlToText(html)`: strip tags, `<br>`/`</p>`/`</div>` → newline, decode
  entities (`&amp; &lt; &gt; &quot; &#39; &nbsp; &#NNN;`), keep link targets as
  `text (url)`, replace `<img>` with `[image]`, collapse blank lines. No new
  dependency.
- `normaliseMessage(m, meMri)` → `{ id, time (ISO from originalarrivaltime or
  composetime), sender, fromMe, text, subject?, replyTo? (channel root id when
  it differs from id), edited, attachments: { files: [{ name, type, openUrl,
  url }], images: [{ ref, url }] } }` where image `ref` is a stable id the
  attachment tool accepts (the asm object id).
- `matchConversations(summaries, query)`: case-insensitive substring on title
  and team; exact id match wins.

Tests cover every branch above with fixtures, including a quoted-reply HTML
body, a deleted message, a system message, a channel reply, a 1:1 from me.

## Task 4 — list tools on the API

**Files:** modify `src/tools/chat.ts`, `src/tools/channels.ts`, their tests;
`src/index.ts` wiring (build ONE `ChatServiceClient` beside the existing
`TeamsClient`, pass both via `deps`; update registrar signatures).

- `teams_list_chats({ unread_only?, query?, limit? })` → chat/meeting/1:1
  summaries sorted by `lastActivity` desc. Description must say it covers
  every chat, not just the visible sidebar, and that ids feed
  `teams_read_conversation`.
- `teams_list_teams_and_channels({ unread_only?, query?, limit? })` →
  channels sorted by team then title, each with `unread`.
- Remove `teams_get_open_chat_messages` and `teams_get_open_channel_posts`,
  `tools/conversation.ts`, and the now-unused DOM selectors/methods
  (`CHAT_MESSAGES_SELECTOR`, `CHAT_LIST_SELECTOR`,
  `TEAMS_AND_CHANNELS_SELECTOR`, `CHANNEL_POSTS_SELECTOR`,
  `OPEN_CONVERSATION_SELECTOR`) — keep `ACTIVITY_FEED_SELECTOR` and the DOM
  transport for `teams_get_activity`. Update `client.ts`'s module doc to say
  only the Activity feed is DOM-read now and why.
- Errors through a new `wrapApiError` (in `tools/errors.ts`) that turns
  capture failures into the "open a signed-in Teams tab and click a chat"
  hint; keep `wrapBridgeError` for activity.

## Task 5 — `teams_read_conversation`

**Files:** create `src/tools/conversation.ts` (replacing the old one),
`tests/tools/conversation.test.ts`; register in `index.ts`.

Input: `{ conversation: string (id or name), unread_only?: boolean (default
false), since?: ISO, limit?: 1..500 (default 50) }`.

- Resolve: id match, else `matchConversations`; 0 matches → error listing
  nothing; >1 → `McpToolError` naming up to 10 candidates (title, kind, id)
  and asking for the id.
- Fetch with `stopAt` = message id ≤ horizon (unread_only) / time < since /
  `limit` reached. Drop system messages. Return oldest-first.
- Channels: group into `threads: [{ root, replies[] }]`; a reply whose root is
  outside the window still appears, with `root: { id, missing: true }`.
- Result: `{ conversation: summary, unread_count, messages | threads,
  truncated? }` via `untrustedResult`. Description: works for any chat or
  channel without opening it; does NOT mark messages read; attachments are
  listed per message and fetched with `teams_get_attachment`.

## Task 6 — `teams_get_unread`

**Files:** create `src/tools/unread.ts`, `tests/tools/unread.test.ts`.

Input: `{ kinds?: ('chat'|'channel'|'meeting'|'oneOnOne')[], query?,
per_conversation_limit?: 1..200 (default 50), max_conversations?: 1..50
(default 20) }`.

- List conversations, keep unread (`lastMessage` newer than horizon and not
  from me), most recent first, cap `max_conversations` (report the rest as
  `more_unread_conversations: [{ id, title }]`).
- Fetch each with `stopAt` at the horizon, concurrency 4.
- Result: `{ total_unread, conversations: [{ conversation, messages |
  threads }] }` — never touches the read marker; description says so (the
  user's Teams badges are unaffected).

## Task 7 — `teams_get_attachment`

**Files:** create `src/tools/attachment.ts`, `tests/tools/attachment.test.ts`.

Input: `{ conversation: string (id), message_id: string, ref: string }` (`ref`
from `attachments.images[].ref` or a file's `name`).

- Re-fetch that message (page until found, cap), locate the attachment.
- Image → `fetchImage` → return `{ content: [{ type: 'image', data: base64,
  mimeType }, { type: 'text', … caption with untrusted note }] }`.
- File → `untrustedResult({ name, type, openUrl, url, content_available:
  false, hint: 'Shared files live in SharePoint; this MCP cannot download
  their contents yet — open the link.' })`.

## Task 8 — healthcheck, manifests, docs

- `teams_healthcheck`: report token source/expiry (never the token) and one
  conversations page count; keep the bridge status section for the activity
  transport.
- `manifest.json` `tools` list, `server.json`, plugin metadata descriptions,
  README ("How it works" — rewrite the one-real-limitation section: chats and
  channels no longer need to be open; only Activity reads the page),
  `skills/teams-fpx/SKILL.md` (document the chatsvc capture + curl recipe;
  keep the DOM recipe only for Activity), `tests/served-tools.ts`,
  `tests/tool-annotations.test.ts`, `tests/server-boot.test.ts`.
- Packaging tests must pass (manifest ↔ served tools sync).

## Live verification (orchestrator, after all tasks — needs the user)

Build, then with the user's Teams tab open: `teams_list_chats unread_only`
count ≈ Teams' unread badge; `teams_read_conversation` on a named chat and a
channel; `teams_get_unread`; `teams_get_attachment` on an inline image; a
shared file returns metadata + link.

## Out of scope (follow-ups)

- Downloading shared-file content (SharePoint token per host, captured on
  demand) — spike.
- Mark-as-read, sending messages.
- Hosting on mcp-host.

---

## Task 8b — review follow-ups from Part 1

- `conversationSummary.unread`: compare the conversation's horizon against the
  newest NON-system, non-deleted message id. `lastMessage` alone is wrong when it
  is a system event (`Event/Call`, `ThreadActivity/*`) or deleted: treat such a
  conversation as unread when `lastMessage.id > horizon` and it is not from me,
  and mark `unread_uncertain: true` so the read tools confirm by fetching.
  `lastMessagePreview`/`lastSender`: skip system events (preview `"(system event)"`),
  and cap the preview at 200 chars.
- Team naming: when the team thread's topic is empty, fall back to the
  `threadProperties.spaceThreadTopic`/`topicThreadTopic` if present, else
  `"(team)"`; a team's own thread lists with title `"General"`.
- `teams_get_unread`: a failure on one conversation must not fail the call —
  record `{ conversation, error }` for it and carry on. `total_unread` must say
  whether it is exact (`total_unread_exact: false` when conversations were
  capped or per-conversation limits truncated).
- Share one `makeLimitedStopAt(horizon, wanted, limit)` helper between
  `teams_read_conversation` and `teams_get_unread`. In `teams_read_conversation`
  set `truncated` whenever the client stopped at its page/message cap too, and
  label `unread_count` as window-scoped when `limit`/`since` cut it.
- `skills/teams-fpx/references/teams-env.sh`: never change the caller's umask —
  use a subshell `(umask 077; …)`.
- healthcheck: tolerate non-JSON text from `runBridgeHealthcheck`.
- `getMessages`: guard `max <= 0`.

---

# Part 2 — writes: send (quote-reply, mentions, emoji) and reactions

Added 2026-10-10 after the read tasks landed. Shapes below were captured live
from Teams' own requests (ids replaced). The token, base URL and client from
Part 1 are reused; writes go through the same `ChatServiceClient`.

## Verified write shapes

- **Send.** `POST {base}/users/ME/conversations/{encodeURIComponent(convId)}/messages`
  → 201; `Location` header ends `/messages/<newId>`. Body:
  `{ id: "-1", type: "Message", conversationid, conversationLink: "blah/<convId>",
  from: meMri, fromUserId: meMri, composetime: ISO, originalarrivaltime: ISO,
  content: "<p>…</p>", messagetype: "RichText/Html", contenttype: "Text",
  imdisplayname: <my display name — from the JWT \`name\` claim>,
  clientmessageid: <random 19-digit numeric string>, callId: "", state: 0,
  version: "0", amsreferences: [], properties: { importance: "", subject: "",
  title: "", cards: "[]", links: "[]", mentions: "[]", onbehalfof: null,
  files: "[]", policyViolation: null, formatVariant: "TEAMS" },
  crossPostChannels: [] }`. Note `properties.*` values are JSON STRINGS.
- **Channel reply** (unverified — verify live): post to
  `encodeURIComponent(channelId + ";messageid=" + rootId)` as the conversation
  id in the URL path; a new top-level channel post may set
  `properties.subject`.
- **Quote reply.** Same POST; content is
  `<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="{qid}"><strong itemprop="mri" itemid="{qSenderMri}">{qSenderName}</strong><span itemprop="time" itemid="{qid}"></span><p itemprop="preview">{escaped preview ≤ 200 chars}</p></blockquote><p>{reply}</p>`
  and `properties.qtdMsgs = JSON.stringify([{ messageId: qid, sender: qSenderMri, time: Number(qid) }])`.
- **Inline emoji in text:** `<span contenteditable="false" title="{Title}" type="({key})" class="animated-emoticon-20-{key}" itemscope><img itemscope itemtype="http://schema.skype.com/Emoji" itemid="{key}" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/{key}/default/20_f.png" title="{Title}" alt="{unicode}" style="width:20px;height:20px;"></span>`.
  Plain unicode emoji in the text are also acceptable; only emit the span form
  for shortcodes the caller passes as `:key:`.
- **Mention** (documented format, NOT captured — verify live): in content
  `<span itemscope itemtype="http://schema.skype.com/Mention" itemid="{n}">{displayName}</span>`,
  and `properties.mentions = JSON.stringify([{ "@type": "http://schema.skype.com/Mention", itemid: "{n}", mri: "{personMri}", mentionType: "person", displayName: "{displayName}" }])`,
  `n` = 0,1,… in order of appearance.
- **Reaction.** `PUT {base}/users/ME/conversations/{encodeURIComponent(convId)}/messages/{msgId}/properties?name=emotions`
  body `{ "emotions": { "key": "<key>", "value": <epoch ms> } }` → 2xx.
  👍 = `like`. Known named keys: `like, heart, laugh, surprised, sad, angry,
  yes, party, fire, smile, cool, wink, hearteyes`; unicode-style keys look like
  `1f389_partypopper`. Removal (unverified): `DELETE` same URL with body
  `{ "emotions": { "key": "<key>" } }`.
- **Read-back.** Existing reactions appear on read messages as
  `properties.emotions` (JSON string or array of `{ key, users: [{ mri, time }] }`)
  — Part 1's `normaliseMessage` should surface them as
  `reactions: [{ key, count, byMe }]` (add this in Task 9 if Part 1 did not).

## Safety rules for every write tool

- Use `confirmWrite`/`confirmTokenParam` from `@chrischall/mcp-utils` exactly
  as `~/git/office-outlook-mcp/src/tools/writes.ts` does: the first call returns
  a preview (conversation title + kind, the rendered text, mentions, quoted
  message preview, reaction) and a token; only the second call with the token
  posts. Read that file first and match it.
- Non-idempotent sends go through a `writeOrUnknown` equivalent: a timeout or
  dropped connection after the request was sent returns `status: "unknown"`
  with "check the conversation before retrying", never a plain error.
- After a send, read back the newest messages and confirm a message with our
  `clientmessageid` exists; warn if not. After a reaction, read the message
  and confirm the key is present with `byMe`.
- Annotations: send `destructiveHint: true, idempotentHint: false,
  openWorldHint: true`; react `destructiveHint: false, idempotentHint: true`.
- Never send to `48:*` system conversations except `48:notes` (self-chat —
  useful for testing; say so in the description).

## Task 9 — `teams_send_message`

**Files:** `src/tools/send.ts`, `src/compose.ts` (pure HTML builders),
tests for both; `ChatServiceClient.sendMessage(convId, body)` and
`getMessage(convId, id)` in `src/chatsvc.ts` (+ tests).

Input `{ conversation (id or name, resolved like teams_read_conversation),
text (plain text; newlines → <br>; HTML-escaped), reply_to?: message id
(channel: posts as a threaded reply to that root; chat: quote-reply),
quote?: message id (explicit quote-reply in any conversation),
mentions?: [{ name_in_text: string, person: string (display name or MRI) }],
confirmToken? }`.

- Resolve mention people: MRI as-is; else match display names among recent
  senders of that conversation (`imdisplayname` → `from` MRI); ambiguous or
  unknown → `McpToolError` listing candidates. Replace the FIRST occurrence of
  `@{name_in_text}` (or `name_in_text`) in the text with the mention span.
- `:key:` shortcodes → emoji spans; everything else escaped.
- Quote: fetch the quoted message for sender MRI/name and preview text.
- Return `{ status: 'sent', conversation, message_id, verified: boolean }`.

## Task 10 — `teams_react`

**Files:** `src/tools/react.ts` (+ tests); `ChatServiceClient.setReaction` /
`removeReaction`.

Input `{ conversation (id or name), message_id, reaction: string, remove?:
boolean, confirmToken? }`. `reaction` accepts aliases → key: `thumbsup|👍|like`
→ `like`, `heart|❤️` → `heart`, `laugh|😆` → `laugh`, `surprised|😮` →
`surprised`, `sad|😢` → `sad`, `angry|😡` → `angry`, plus any raw key passed
through. `checkmark|✅` → read from a constant `CHECKMARK_KEY` that defaults to
`2705_whiteheavycheckmark` with a comment that it is UNVERIFIED and must be
confirmed live (the orchestrator will fix it after reading a real ✅ reaction).
Confirm-gated (it is visible to everyone in the chat). Read-back verification.

## Task 11 — manifests/docs for the write tools

manifest tools list, README (writes section: confirm flow, self-chat for
testing), skill doc, served-tools/annotation tests.

## Live verification for Part 2 (orchestrator + user, explicit go-ahead per send)

In `48:notes`: plain send, quote-reply, :emoji:, 👍 react + remove. Read a
message carrying ✅ to learn its key and fix `CHECKMARK_KEY`. Mention: only in
a chat the user picks, with their explicit go-ahead.

## Addendum to Task 10 — reaction key rule (from captured recentReactions, 2026-10-10)

Any emoji is a valid reaction key, not just the named ones: 😬 reacted as
`1f62c_grimacingface`. Unicode-style keys are `<lowercase hex code point(s)>_<Unicode
character NAME, lowercased, spaces/punctuation removed>` — e.g. 💯 `1f4af_hundredpointssymbol`,
🛫 `1f6eb_airplanedeparture`, 🌍 `1f30d_earthglobeeuropeafrica`, 👀 `1f440_eyes`. Hence ✅
should be `2705_whiteheavycheckmark` (still confirm live). Teams also keeps short legacy
names for common ones (`like`, `heart`, `laugh`, `surprised`, `sad`, `angry`, `yes`, `no`,
`party`, `fire`, `smile`, `cool`, `wink`, `hearteyes`, `manshrug`, `fistbump`, `handsinair`).
`teams_react` must: accept a raw key; map the alias table; and for a single unicode emoji not
in the table, return an McpToolError asking for the key (no Unicode-name database is bundled)
with the rule above as the hint. Do NOT write the user's `recentReactions` property (Teams does
that as a UI nicety; we must not touch preferences).
