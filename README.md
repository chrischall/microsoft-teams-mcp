# microsoft-teams-mcp

MCP server for **Microsoft Teams** — list every chat and channel you are in,
read any of them by id or name, pull everything you have not read yet,
read your Activity feed, fetch inline images, and — each one confirmed by you first — post messages
and reactions, using your signed-in browser session.

> This project was developed and is maintained by AI (Claude Code). Use at your
> own discretion.

## How it works

Chats and channels are read through Teams' own **chat-service REST API**
(`https://teams.cloud.microsoft/api/chatsvc/<region>/v1`). The server needs a
bearer token for it, and the only place one exists is your signed-in
`teams.cloud.microsoft` tab: through the
[fetchproxy](https://github.com/chrischall/fetchproxy) browser bridge it
captures the `Authorization` header off a request the Teams web app makes to
its chat service — **once**. From then on every call is a plain server-side
request; the token is cached in `~/.microsoft-teams-mcp/token.json` for its
~23h lifetime (so a restart needs no browser), and a 401 re-captures once and
replays. `teams.microsoft.com` (the old host) is not supported.

So **chats and channels no longer need to be open in the browser** — any of
them can be read by id or name, and listing covers every chat, not just the
ones the sidebar shows. To capture a token, an idle tab is not enough: click
into any chat so Teams calls its chat service.

**The Activity feed comes from the same API.** It is a chat-service
conversation of its own (`48:notifications`), so `teams_get_activity` reads
it like any chat — no Activity view (or any view) needs to be open, and
nothing reads the page. The bridge is used for one thing only: capturing the
token.

Nothing here marks anything read. Reading a chat through this server leaves
your Teams unread badges exactly as they were. The two writes —
`teams_send_message` and `teams_react` — go through the same chat-service
API and post nothing until you confirm (see [Writes](#writes)).

## SharePoint file download

A **shared file** in a chat or channel is not in Teams' own object store
(where inline images live) — it is in **SharePoint**, which does not accept the
chat-service token. So `teams_get_attachment` captures a **separate SharePoint
token** and fetches the file through SharePoint's `v2.1` shares API, then
downloads the pre-signed link it returns (with no token attached).

Two things follow from that, both one-time and additive — they do **not** change
how the chat-service token is captured:

- **A one-time re-pair.** The SharePoint path declares a second trust scope,
  `sharepoint.com`, so the ContextMint Bridge grant has to cover it. If you
  paired `microsoft-teams-mcp` before this feature, revoke and re-pair it in the
  bridge popup once to add the SharePoint scope (a `not in declared set` error
  on a file download is this, nothing more). The chat-service capture is
  untouched — it still trusts only `teams.cloud.microsoft`.
- **Open the file once per SharePoint host.** The SharePoint token is captured
  from the browser's own request to the file's SharePoint host, is short-lived
  (~53 minutes), and is scoped to one host family. So the first file from a
  given host in a session may need you to open it once in a signed-in Teams tab
  (that browser request is what the token is read from); after that the token is
  cached and refreshed until it expires. A file your account cannot open, or on
  a SharePoint host your token does not cover, comes back as metadata and its
  open link rather than failing.

## Install

```sh
npm i -g @chrischall/microsoft-teams-mcp
```

Requires the **ContextMint Bridge** browser extension and `@fetchproxy/cli`,
on a matching major version. Install the bridge from
[its releases page](https://github.com/nullnet-app/contextmint-bridge/releases):
in Chrome, unzip the chrome zip and load it unpacked
(`chrome://extensions` → Developer mode → Load unpacked). Safari isn't
available yet (it will ship inside the ContextMint app, which has no public
download), so use Chrome for now.

ContextMint Bridge is the fetchproxy browser extension under its new name,
from the same maintainer — fetchproxy's own README
(https://github.com/chrischall/fetchproxy#extension) points to it. Its source
is public at https://github.com/nullnet-app/contextmint-bridge: build it
yourself, or check a release zip against the `.sha256` file published beside
it (`shasum -a 256 -c contextmint-bridge-chrome-<version>.zip.sha256`).

```sh
npm i -g @fetchproxy/cli
```

The first call prints a 6-digit pair code to approve in the ContextMint
Bridge popup; the grant persists.

### Install in opencode

opencode reads MCP servers from `opencode.json` (project) or
`~/.config/opencode/opencode.json` (global):

```json
{
  "mcp": {
    "servers": {
      "teams": {
        "type": "local",
        "command": ["npx", "-y", "@chrischall/microsoft-teams-mcp"]
      }
    }
  }
}
```

### Configuration

Everything is optional — by default the token is captured from your browser.

| variable | purpose |
| --- | --- |
| `TEAMS_ACCESS_TOKEN` | A chat-service bearer token (the JWT itself) to use instead of capturing one. Never re-captured from the browser: when Teams rejects it, replace it. |
| `TEAMS_DISABLE_FETCHPROXY` | `true` to never capture from the browser (then `TEAMS_ACCESS_TOKEN` is required). |
| `TEAMS_CHATSVC_REGION` | Chat-service region in the API path. Defaults to `amer`. |
| `TEAMS_CHATSVC_BASE` | Full override for the API base (default `https://teams.cloud.microsoft/api/chatsvc/<region>/v1`). |
| `TEAMS_CAPTURE_TIMEOUT` | Seconds to wait for the Teams tab to issue a chat-service request. Defaults to `30`. The bridge deadline is derived from it (window + 15s). |
| `TEAMS_TOKEN_CACHE` | `false` to keep the captured token in memory only. Defaults to on. |
| `TEAMS_TOKEN_FILE` | Where the captured token is cached. Defaults to `~/.microsoft-teams-mcp/token.json`. |
| `TEAMS_TIME_ZONE` | Default IANA time zone (e.g. `America/New_York`) for the `tz` argument of `teams_read_conversation` and `teams_get_unread`: each message then carries `time_local` (ISO 8601 with offset) next to its UTC `time`. Unset means no `time_local` unless `tz` is passed; a name `Intl` does not know is reported as an error. |
| `TEAMS_WS_PORT` | fetchproxy concentrator port. Defaults to `37149`, the fleet-wide shared port. |
| `FETCHPROXY_WS_HOST` | Bind interface for the bridge (default `127.0.0.1`; a literal IP only, anything else is ignored). Keep on loopback — it exists for ONE topology: a hosted child in its own network namespace. |
| `FETCHPROXY_IDENTITY_DIR` | Where to store the long-term fetchproxy identity keypair (default `~/.fetchproxy/identity/`; an absolute path only, anything else is ignored). Exists for a host that runs a separate child per caller, so they can share one identity. |

## Tools

- `teams_list_chats` — every chat you are in (1:1, group and meeting chats,
  not just the visible sidebar, plus your `48:notes` self-chat as kind
  `self` when Teams lists it), favorites and pinned chats first, then most
  recent activity first: id, kind,
  title, unread hint, ISO-8601 last-activity time, last-message preview and
  sender. `unread_only` (exact — see below), `since_days` (with
  `unread_only`; default 14), `older_unread` (with `unread_only`; lists what
  `since_days` left out — see below), `query` (title, a participant's name, or the
  last-message preview), `search_messages`,
  `since` / `until` (ISO 8601 date-times: keep chats whose last activity is
  at or after / at or before them; a chat with no last activity is left out
  when either is given), `sort` and `limit` (default 50, max 500; keeps the
  head of the sort order). `sort` is `activity` (default — favorites, pinned,
  then most recent; with `unread_only`, most recent), `activity_only` (most
  recent first, ignoring favorites and pinned) or `title` (alphabetical,
  case-insensitive — this titles every untitled chat in the listing before
  `limit` applies, so it reads more than the other orders). A group chat with no topic is titled after
  who is in it — `Group chat: <up to three recent senders other than you,
  most recent first>` plus `+N` for more, with every recent sender in
  `participants` — rather than after whoever posted last, so chats that
  share a last sender stay distinguishable. The chat service has no member
  list, so this reads up to 30 recent messages (one small page) of each such
  chat returned (four at a time, cached until a new message arrives); if that
  read fails, the chat keeps its last-sender title. A 1:1 is titled by the
  other person (its id names them) and carries them as its one
  `participants` entry — read from their recent messages when you sent the
  last one, `(1:1 chat)` if they wrote nothing recently. `query` matches such
  chats by their participants, never by the made-up `Group chat:` or
  `(1:1 chat)` titles, so `group` or `chat` does not hit every untitled chat.
  With a `query`, every row carries `match_reason` — the first of `id` (an
  exact id), `title`, `team`, `participant` (an untitled chat's participant,
  including a 1:1's other person) or `message` that applies. `message` means
  the query was found in what was said, not in what the chat is called: by
  default only the last-message preview the listing already carries is
  checked (no extra reads). `search_messages: true` (needs `query`) also
  reads the last 50 messages of every chat no name or preview matched, four
  at a time, and returns the ones where a message a person wrote contains
  the query, with the newest such message as `matched_message` (`id`,
  `time`, `sender`, `text` cut to a window around the match). It is one read
  per chat, so narrow a large listing with `since` / `until` or
  `unread_only`; chats whose read failed are counted in
  `messages_not_searched`.
  Conversations the chat service marks favorite or pinned carry
  `favorite: true` / `pinned: true` and lead the plain listing (favorites,
  then pinned); `unread_only` keeps most-recent-first order but carries the
  flags. Note the new Teams **Favorites section for chats** is kept by a
  different Teams service this MCP does not read, so a chat you put there is
  not flagged (verified live 2026-10-10) — the flags mostly reflect favorited
  channels and pinned conversations. A chat you muted in Teams carries
  `muted: true`.
- `teams_list_teams_and_channels` — every channel you are in, by team then
  title, with the same fields — favorite and pinned channels first, flagged
  as in `teams_list_chats`. Channels you have hidden in Teams are listed
  after the rest with `hidden: true`, and left out of `unread_only` unless
  `include_hidden` is set. `unread_only`, `include_hidden`, `since_days`
  (with `unread_only`; default 14), `query` (channel or team name, or the
  last-message preview; each match carries `match_reason`, as in
  `teams_list_chats`) and
  `limit` (default 50, max 500).
- `teams_read_conversation` — read any chat or channel by id or name,
  oldest first: sender, time (UTC, ISO 8601), text, subject, edited/deleted
  flags and attachments. A channel comes back as threads (root post + replies).
  A thread whose root post is older than the window (a reply came in after
  `since`, `limit` or `after_message_id` cut the read) still shows its root:
  it is fetched by id and marked `outside_window: true` — context only, not
  counted in `unread_count` or `message_count` and never the `next_cursor`.
  Up to 25 roots are fetched per call (shared by every conversation it
  reads, newest threads first); one past that, or one that cannot be
  fetched, is `{ id, missing: true }`.
  `unread_only`, `since` (ISO 8601), `limit` (default 50, max 500, newest
  kept) and `tz` (an IANA time zone such as `America/New_York`, default
  `TEAMS_TIME_ZONE`), which adds `time_local` — the same instant in that zone,
  ISO 8601 with its offset — next to each message's `time`. Message text is
  never truncated, and a table's cells are separated by ` | `, one row per
  line. The resolved `conversation` carries `match_reason` (`id`, `title`,
  `team` or `participant`); a name never resolves by what was said in a
  chat, since `teams_send_message` and `teams_react` name their target the
  same way. An ambiguous name is not an error here: the result is
  `{ambiguous: true, query, match_count, candidates, hint}` — up to 10
  candidates, each with `id`, `title`, `kind`, `team` (channels),
  `lastActivity` and `match_reason` — and no messages, so you can read again
  by one candidate's id (`teams_send_message` and `teams_react` still refuse
  an ambiguous name outright, since a write never guesses); an
  untitled group chat can be named by one of its participants, and a 1:1 by
  the other person (this, and the same naming in `teams_send_message` /
  `teams_react`, titles every untitled chat first).
  `48:notes` — or the name `notes` / `self` — is your notes self-chat.
  `after_message_id` returns only messages newer than that id (the read
  stops at the first message at or before it, alongside `unread_only`,
  `since` and `limit`); every result carries `next_cursor` — the newest
  returned message's id, or the `after_message_id` you passed when nothing
  is new — so passing it back as `after_message_id` reads only what arrived
  since. See [Cursors](#cursors-only-what-is-new).
  `format` sets how much comes back: `full` (default, everything above),
  `compact` (each message only `id`, `time`, `time_local`, `sender`, `text`,
  plus `subject` when set — no reactions, attachments or edited/deleted
  flags — and the conversation only `id`, `kind`, `title`, `team`) or
  `summary` (no messages: `message_count`, `newest_time` and, for a channel,
  one `{ root_id, subject, reply_count, newest_time }` per thread, with
  `root_outside_window: true` or `root_missing: true` for a root from
  outside the window).
  `conversation` may also be an array of up to 20 ids or names: the result
  is then `{ conversations: [...], errors, fetched_at }`, one single-read
  result per conversation (the other arguments apply to each), read four at
  a time; one that cannot be resolved or read is listed in `errors` with its
  reason and the rest still come back. A single string returns exactly the
  shape above.
- `teams_get_unread` — every unread message across your chats and channels,
  grouped by conversation, most recent first. Filter by `kinds` (`chat`,
  `channel`, `meeting`, `oneOnOne`) or `query` (title, channel or team
  name, an untitled chat's participant name, or the last-message preview —
  each conversation then carries `match_reason`, as in `teams_list_chats`);
  `per_conversation_limit`
  (default 50) and `max_conversations` (default 20 returned with their
  messages — the rest are named, with their `unread_count`, in
  `more_unread_conversations`). Hidden channels are left out unless
  `include_hidden`; chats you muted in Teams (`muted: true`) are left out
  unread and counted in `muted_skipped` unless `include_muted`. Only
  conversations with unread messages from the last `since_days` (default 14,
  max 3650) are returned; older ones are counted in `older_unread_skipped`,
  and `older_unread: true` also lists them (see below). An untitled group chat or 1:1 is titled after
  its participants, as in `teams_list_chats` — in `conversations`, in
  `more_unread_conversations` and in `errors`. `tz` (default
  `TEAMS_TIME_ZONE`) adds each message's `time_local`, as in
  `teams_read_conversation`. A channel's unread replies come grouped under
  their root post; a root you had already read is fetched and marked
  `outside_window: true` (not unread, not counted), up to 25 per call as in
  `teams_read_conversation`, else `{ id, missing: true }`. `cursors` — `{ "<conversation id>": "<message id>" }`
  — returns, per conversation, only unread messages newer than its cursor
  (each read stops at the later of the read marker and the cursor; a
  conversation with nothing past its cursor is left out). Each returned
  conversation carries `next_cursor`, and the result carries `next_cursors`:
  the cursors you passed, updated with every returned conversation's
  `next_cursor` — pass it back as `cursors` next time.
  `format` (`full`, `compact` or `summary`) trims each conversation as in
  `teams_read_conversation`; every format keeps the conversation `id`,
  `unread_count` and `next_cursor`.
- `teams_get_changes` — which chats and channels got new messages since a
  cursor: counts, not messages. Each changed conversation comes back with its
  summary, `new_message_count` (real messages only — call recordings and other
  system events, and deleted messages, are not counted),
  `new_from_others_count`, `newest_time` and, for a channel, `threads`: one
  `{ root_id, new_root, new_reply_count }` per thread the new messages landed
  in, including a thread whose root post is older than the cursor. Without
  `cursor` it looks back 24 hours. `next_cursor` is an opaque token (base64url
  of `{"t": <epoch ms>}`; `cursor` also takes a plain ISO 8601 time) — pass it
  back to see only what arrived since. It is at-least-once: a conversation can
  come back twice, but none is skipped, and it stays put when a read failed
  (`errors`). `kinds` and `include_hidden` filter as in `teams_get_unread` —
  use the same ones with a cursor as the call that returned it; `limit`
  (default 50, max 200) caps the conversations returned, the rest are named in
  `more_conversations`. Then read what changed with `teams_read_conversation`:

  ```jsonc
  // first call: the last 24 hours
  teams_get_changes({})
  // → { since, total_new_messages: 5, conversations: [
  //      { conversation: { id: "19:…@thread.tacv2", kind: "channel", title: "Design", … },
  //        new_message_count: 3, new_from_others_count: 3, newest_time: "…",
  //        threads: [{ root_id: "1791803726629", new_root: false, new_reply_count: 2 }, …] }, … ],
  //    next_cursor: "eyJ0IjoxNzkxODE1MjAxMzYxfQ", next_cursor_time: "…", fetched_at }
  // later: only what arrived since
  teams_get_changes({ cursor: "eyJ0IjoxNzkxODE1MjAxMzYxfQ" })
  ```

  A channel reply moves only the listing's `lastUpdatedMessageId` (its
  `lastMessage` stays on the newest root post), so that is what nominates a
  channel; each conversation is then read newest *change* first down to the
  cursor, since an old message edited after it is served ahead of new ones.
  A hidden channel's listing entry stops moving altogether, so with
  `include_hidden` every hidden channel is read.
- `teams_get_attachment` — one attachment of a message: an inline image is
  returned as the image itself; a shared file is downloaded from SharePoint and
  returned as its contents — an image as the image, anything else as an embedded
  resource carrying the bytes. Shared-file download uses a separate, short-lived
  SharePoint token captured per SharePoint host, so the first file from a given
  host may need [opening once in Teams](#sharepoint-file-download) to mint it; a
  file your account cannot open, or one that could not be fetched, comes back as
  metadata and its open link instead of failing.
- `teams_get_activity` — the Activity feed (the bell icon), newest first:
  mentions (including @team / @everyone), reactions, replies, meeting
  notices and missed calls across every chat and channel. Each row has its
  time (ISO 8601), `kind` (`mention`, `reaction`, `reply`, `meeting`, `call`
  or `other`) with Teams' raw type/subtype, whether you have seen it, who
  did it, where (`conversation_id`, title, type), the `message_id` and a
  plain-text preview — pass `where.conversation_id` to
  `teams_read_conversation` to read it in context. Filter with
  `unread_only`, `kinds` and `since_days` (default 14, max 3650); `limit`
  (default 50, max 200). Read from the chat service, so the Activity view
  does not need to be open; it never marks anything read.
- `teams_healthcheck` — resolves the token the way every tool does
  (capturing one on a cold start) and probes the chat service with one
  conversations listing. Reports where the token came from and when it
  expires (both `null` when no token is held, e.g. after a failed capture)
  and how many conversations the listing returned — never the token.

**Write** (both ask you to confirm first — see [Writes](#writes)):

- `teams_send_message` — post as you to any chat or channel, by id or name.
  `text` is plain text: it is escaped, newlines are kept, and `:key:`
  shortcodes (`:smile:`, `:1f389_partypopper:`) become Teams emoji.
  `reply_to` posts into a channel post's thread — it must be a root post in
  that channel, not a reply (in a chat it quote-replies that message); `quote` quote-replies a message anywhere; `mentions`
  @-mentions people by MRI or by a display name matched among the
  conversation's recent senders (an ambiguous name is refused with the
  candidates). The result says whether the post was found again by reading
  the conversation back (`verified`).
- `teams_react` — add or (with `remove: true`) withdraw your reaction on a
  message: an emoji or name (`thumbsup`/👍, `heart`/❤️, `laugh`/😆,
  `surprised`/😮, `sad`/😢, `angry`/😡, `checkmark`/✅) or any raw Teams
  reaction key (`fire`, `1f389_partypopper`). The message is read back to
  verify the change (`verified: null` when it is not found again). It writes only that message's reaction — never your
  recent-reactions list or any other preference.

Every other tool is read-only, and none of them marks anything read.

## Writes

Both writes are visible to everyone in the conversation, so each asks you
first. A client that can show a confirmation prompt (Claude Code) shows one.
On a client that cannot, the first call posts nothing and returns a
`confirmation-required` preview — the conversation, the exact text, any
mentions, the quoted message or the reaction, with everything other people
wrote (titles, names, quoted text) fenced as untrusted — plus a `confirmToken`; only a
repeat call with that token posts. The token is single-use, expires, and is
bound to the exact arguments: change anything between the two calls and it
is refused with a fresh preview.

| variable | default | |
| --- | --- | --- |
| `MCP_CONFIRM_MODE` | `ask-user` | What a write does on a client that cannot show a confirmation prompt (claude.ai, Claude Desktop). `ask-user`: two steps, and the model must get your approval in chat before calling again with the token. `auto`: the same two steps, but the model may use the token after reviewing the preview itself. `refuse`: writes are refused on such clients. An unrecognised value is treated as `refuse`. |
| `MCP_CONFIRM_ELICITATION` | `on` | `off` never shows a confirmation prompt, so every client gets the `MCP_CONFIRM_MODE` behaviour. Set it for a client that says it can show prompts but never does. |
| `MCP_CONFIRM_TTL_SECONDS` | `600` | How long a token stays valid. |
| `MCP_CONFIRM_SECRET` | random per process | Signing key; set it only if tokens must survive a server restart. |

On a client without prompts the token is something the model passes back
itself, so keep `MCP_CONFIRM_MODE=ask-user` and approve each preview in chat
as the real guard on what gets posted.

**Try it in your own notes first.** The conversation `48:notes` (also by the
name `notes` or `self`) is your Teams self-chat ("notes to self"): both writes
accept it, `teams_read_conversation` reads it, nobody else sees it, and
it is the safe place to see what a message or reaction will look like. Every
other `48:` system conversation (calendar and the like) is refused.

**A send that times out is not retried.** If Teams does not confirm a post (a
timeout, a dropped connection or a 5xx after the request went out), the
result is `status: "unknown"` — the message may already be there, so read the
conversation before sending again. A 4xx is a refusal and stays an error. A
reaction Teams does not confirm is settled by reading the message back, and
is `status: "unknown"` only when that does not show the change. Threaded channel replies, @mentions,
removing a reaction and the ✅ reaction's key follow Teams' documented or
inferred shapes but have not been verified against a live tenant yet.

**No Calendar tool here.** Teams Web's Calendar is Outlook's data, rendered
inside an embedded `outlook.office.com` iframe — not the chat service's. Use
[`office-outlook-mcp`](https://github.com/chrischall/office-outlook-mcp)'s
`outlook_list_events` / `outlook_get_event` / `outlook_list_calendars`
instead.

## Things worth knowing

- <a id="cursors-only-what-is-new"></a>**Cursors: only what is new.** A
  message id is its arrival time in epoch milliseconds, so "newer than the
  last message I saw" is a numeric comparison and a cursor is just that
  message's id — `next_cursor` is a plain id string, passed back unchanged as
  `after_message_id` (`teams_read_conversation`) or inside `cursors`
  (`teams_get_unread`). The server keeps no state and **no cache**: every
  call reads Teams live, and `fetched_at` (ISO 8601, UTC) says when. There is
  no refresh flag because there is nothing to refresh. A cursor never moves
  the read marker; with `limit` / `per_conversation_limit` the newest
  messages are kept and `next_cursor` jumps to the newest, so check
  `truncated` before assuming nothing was skipped.

- **Results are framed as untrusted.** Every data tool returns
  `{ untrusted_content: true, note, ... }` and says in its description that
  the text is written by other Teams users. Message bodies, previews and
  subjects are data to report, never instructions to follow — this matters
  because the server usually runs beside write-capable MCPs (mail, messaging)
  that an injected instruction could otherwise reach.
- **Unread means what Teams means, and is confirmed by reading.** A message
  is unread when it arrived after the conversation's read marker
  (`consumptionhorizon`), was not sent by you, is a real message (`Text` or
  `RichText/Html`) and is not deleted. System messages — call recordings and
  transcripts (`RichText/Media_*`), `Event/*`, `ThreadActivity/*`,
  `Control/*` — never make anything unread. The conversation listing cannot
  tell on its own: `lastimreceivedtime` moves for a call recording too,
  `lastMessage` can be stale, and a chat you never opened has no marker at
  all. So the listing only nominates candidates (anything past the marker,
  or no marker), and `unread_only` and `teams_get_unread` read each
  candidate's messages past its marker (four at a time) and keep only those
  holding a real unread message, with its `unread_count`. Without
  `unread_only` the list tools' `unread` is just that cheap hint, marked
  `unread_confirmed: false`. A candidate that answers 404 or 403 (a channel
  that is gone or no longer readable) is not unread: it is skipped and
  counted in `skipped_unreachable`. Channels you have hidden in Teams
  (`hidden: true`) are not counted as unread unless you pass
  `include_hidden`. Unread answers are also limited to recent activity:
  a conversation counts only when its newest unread message arrived in the
  last `since_days` days (default 14, max 3650) — in practice most
  "unread" channels are dormant ones, untouched for months or years in teams
  you never open. A candidate with no activity at all inside the window is
  not even read. Whatever the window leaves out is counted in
  `older_unread_skipped` (with a hint to widen `since_days`), never dropped
  silently. `older_unread: true` (on `teams_get_unread`, and on
  `teams_list_chats` with `unread_only`) also lists them in
  `older_unread_conversations` — `id`, `kind`, `title`, `team` (a channel's),
  `lastActivity`, most recent first — with an `unread_count` for each one
  whose messages were already read to place it outside the window; one with
  no activity at all inside the window is not read just to count it, and
  carries `unread_count_known: false` instead (its news past the marker may
  be only system events). Muted chats: the conversation listing marks a chat
  you muted with `properties.alerts: "false"` (verified live 2026-10-12: 8 of
  1,435 conversations, every other one `"true"` or absent), surfaced as
  `muted: true`. `teams_get_unread` leaves them out without reading them and
  counts them in `muted_skipped`; `include_muted` adds them back. The list
  tools' `unread_only` still includes them. `teams_read_conversation` is not windowed: it takes an explicit
  `since` and `limit`. This server only reads the marker; it never moves it.
- **The token only goes to Teams.** The chat-service token is sent to the
  chat-service base and, for inline images, only to https hosts under
  `.asm.skype.com` (Teams' own image store) — an image link pointing anywhere
  else is refused. Shared-file download uses a **separate** SharePoint token
  (below), sent only to the file's own `*.sharepoint.com` host; the pre-signed
  download link it returns is fetched with **no** token at all.
- **Shared files download from SharePoint.** They live in SharePoint, which
  does not accept the chat-service token, so `teams_get_attachment` captures a
  separate SharePoint token and fetches the file through SharePoint's v2.1
  shares API — see below.
- **Scope grows quietly.** If you see a `not in declared set` error, the
  extension's approved scope is behind the server's declared one — revoke
  and re-pair `microsoft-teams-mcp` in the ContextMint Bridge popup.

## Upgrading from 0.3

0.4 moves every tool onto the Teams chat-service API, and several outputs
changed shape. If you script against them:

- **Removed tools:** `teams_get_open_chat_messages` and
  `teams_get_open_channel_posts` — use `teams_read_conversation`, which reads
  any chat or channel by id or name without it being open.
- **`teams_get_activity`:** rows now come from the Activity feed itself.
  `title` and `location` are replaced by `type`/`subtype`/`kind`, `from`,
  `where: { conversation_id, title, type }`, `message_id`, `reply_chain_id`
  and `preview`; `time` is ISO 8601 (it was prose like "1:19 PM"); `read` is
  new; `empty_hint` is gone (the feed no longer depends on the Activity view).
- **`teams_get_attachment`:** a shared file now returns its **contents** (an
  image as image content, anything else as an MCP embedded resource carrying the
  bytes), not just metadata and a link — see
  [SharePoint file download](#sharepoint-file-download). A file that cannot be
  fetched still returns the old metadata-and-link shape, so a script that only
  reads the metadata keeps working.
- **`teams_healthcheck`:** the top-level `bridge` section is gone — the bridge
  is only used to capture the token now. Token health is under `credential`
  and `chat_service`; a new `sharepoint` note describes the on-demand,
  short-lived SharePoint token (the healthcheck does not require one).
- **Chat titles:** untitled chats are titled `(1:1 chat)` (or the other
  person's name), `Group chat: A; B; C +N` (with a `participants` array), or
  `(untitled group chat)` — no longer the last sender's name or
  `(untitled chat)`. Plain `teams_list_chats` / `teams_list_teams_and_channels`
  listings put `favorite` and then `pinned` conversations first.
- **Unread:** confirmed by reading messages after your read point (call
  recordings and other system events never count), hidden channels are left
  out unless `include_hidden`, and `since_days` defaults to 14 with
  `older_unread_skipped` reporting the rest.

## Development

```sh
npm install
npm run build
npm test              # typecheck + suite
npm run test:coverage # CI's gate
```

## License

MIT
