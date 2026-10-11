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
  `unread_only`; default 14), `query` (title, or a participant's name) and
  `limit` (default 50, max 500). A group chat with no topic is titled after
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
  Conversations the chat service marks favorite or pinned carry
  `favorite: true` / `pinned: true` and lead the plain listing (favorites,
  then pinned); `unread_only` keeps most-recent-first order but carries the
  flags. Note the new Teams **Favorites section for chats** is kept by a
  different Teams service this MCP does not read, so a chat you put there is
  not flagged (verified live 2026-10-10) — the flags mostly reflect favorited
  channels and pinned conversations.
- `teams_list_teams_and_channels` — every channel you are in, by team then
  title, with the same fields — favorite and pinned channels first, flagged
  as in `teams_list_chats`. Channels you have hidden in Teams are listed
  after the rest with `hidden: true`, and left out of `unread_only` unless
  `include_hidden` is set. `unread_only`, `include_hidden`, `since_days`
  (with `unread_only`; default 14), `query` (channel or team name) and
  `limit` (default 50, max 500).
- `teams_read_conversation` — read any chat or channel by id or name,
  oldest first: sender, ISO-8601 time, text, subject, edited/deleted flags
  and attachments. A channel comes back as threads (root post + replies).
  `unread_only`, `since` (ISO 8601) and `limit` (default 50, max 500, newest
  kept). An ambiguous name lists the candidates and asks for the id; an
  untitled group chat can be named by one of its participants, and a 1:1 by
  the other person (this, and the same naming in `teams_send_message` /
  `teams_react`, titles every untitled chat first).
  `48:notes` — or the name `notes` / `self` — is your notes self-chat.
- `teams_get_unread` — every unread message across your chats and channels,
  grouped by conversation, most recent first. Filter by `kinds` (`chat`,
  `channel`, `meeting`, `oneOnOne`) or `query` (title, channel or team
  name, or an untitled chat's participant name); `per_conversation_limit`
  (default 50) and `max_conversations` (default 20 returned with their
  messages — the rest are named, with their `unread_count`, in
  `more_unread_conversations`). Hidden channels are left out unless
  `include_hidden`. Only conversations with unread messages from the last
  `since_days` (default 14, max 3650) are returned; older ones are counted
  in `older_unread_skipped`. An untitled group chat or 1:1 is titled after
  its participants, as in `teams_list_chats` — in `conversations`, in
  `more_unread_conversations` and in `errors`.
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
  silently. `teams_read_conversation` is not windowed: it takes an explicit
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
