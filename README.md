# microsoft-teams-mcp

MCP server for **Microsoft Teams** — list every chat and channel you are in,
read any of them by id or name, pull everything you have not read yet,
fetch inline images, and — each one confirmed by you first — post messages
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

**The one thing that still reads the page is the Activity feed**
(`teams_get_activity`). No chat-service equivalent of it has been verified,
so it reads the DOM your tab has rendered through the bridge — which can read
but never navigate a tab, so it needs the Activity view (the bell icon) open.

Nothing here marks anything read. Reading a chat through this server leaves
your Teams unread badges exactly as they were. The two writes —
`teams_send_message` and `teams_react` — go through the same chat-service
API and post nothing until you confirm (see [Writes](#writes)).

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
| `TEAMS_CAPTURE_TIMEOUT` | Seconds to wait for the Teams tab to issue a chat-service request. Defaults to `30`. The bridge deadline is derived from it (window + 15s), for the capture and for `teams_healthcheck` alike. |
| `TEAMS_TOKEN_CACHE` | `false` to keep the captured token in memory only. Defaults to on. |
| `TEAMS_TOKEN_FILE` | Where the captured token is cached. Defaults to `~/.microsoft-teams-mcp/token.json`. |
| `TEAMS_WS_PORT` | fetchproxy concentrator port. Defaults to `37149`, the fleet-wide shared port. |
| `FETCHPROXY_WS_HOST` | Bind interface for the bridge (default `127.0.0.1`; a literal IP only, anything else is ignored). Keep on loopback — it exists for ONE topology: a hosted child in its own network namespace. |
| `FETCHPROXY_IDENTITY_DIR` | Where to store the long-term fetchproxy identity keypair (default `~/.fetchproxy/identity/`; an absolute path only, anything else is ignored). Exists for a host that runs a separate child per caller, so they can share one identity. |

## Tools

- `teams_list_chats` — every chat you are in (1:1, group and meeting chats,
  not just the visible sidebar, plus your `48:notes` self-chat as kind
  `self` when Teams lists it), most recent activity first: id, kind,
  title, unread hint, ISO-8601 last-activity time, last-message preview and
  sender. `unread_only` (exact — see below), `query` (title) and `limit`
  (default 50, max 500).
- `teams_list_teams_and_channels` — every channel you are in, by team then
  title, with the same fields. Channels you have hidden in Teams are listed
  after the rest with `hidden: true`, and left out of `unread_only` unless
  `include_hidden` is set. `unread_only`, `include_hidden`, `query` (channel
  or team name) and `limit` (default 50, max 500).
- `teams_read_conversation` — read any chat or channel by id or name,
  oldest first: sender, ISO-8601 time, text, subject, edited/deleted flags
  and attachments. A channel comes back as threads (root post + replies).
  `unread_only`, `since` (ISO 8601) and `limit` (default 50, max 500, newest
  kept). An ambiguous name lists the candidates and asks for the id.
  `48:notes` — or the name `notes` / `self` — is your notes self-chat.
- `teams_get_unread` — every unread message across your chats and channels,
  grouped by conversation, most recent first. Filter by `kinds` (`chat`,
  `channel`, `meeting`, `oneOnOne`) or `query`; `per_conversation_limit`
  (default 50) and `max_conversations` (default 20 returned with their
  messages — the rest are named, with their `unread_count`, in
  `more_unread_conversations`). Hidden channels are left out unless
  `include_hidden`.
- `teams_get_attachment` — one attachment of a message: an inline image is
  returned as the image itself; a shared file as its name, type and open
  link only (files live in SharePoint and this server cannot download their
  contents yet — the result says so).
- `teams_get_activity` — the Activity feed (the bell icon in the left nav):
  mentions, replies, reactions, and the like across every chat and channel,
  each with its title, message preview, prose time, and location. Reads the
  page, so it requires the Activity view open.
- `teams_healthcheck` — probes the chat service (one conversations listing,
  which also captures a token on a cold start) and reports where the token
  came from and when it expires (both `null` when no token is held, e.g.
  after a failed capture) — never the token — plus the bridge status the
  capture and the Activity feed rely on.

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
  `include_hidden`. This server only reads the marker; it never moves it.
- **The token only goes to Teams.** It is sent to the chat-service base and,
  for inline images, only to https hosts under `.asm.skype.com` (Teams' own
  image store) — an image link in a message pointing anywhere else is
  refused.
- **Shared files are links, not contents.** They live in SharePoint, which
  does not accept the chat-service token.
- **One Teams tab at a time for the Activity feed.** If more than one
  `teams.cloud.microsoft` tab is open, the bridge reads whichever answers
  first — close extras if `teams_get_activity` looks wrong.
- **Scope grows quietly.** If you see a `not in declared set` error, the
  extension's approved scope is behind the server's declared one — revoke
  and re-pair `microsoft-teams-mcp` in the ContextMint Bridge popup.

## Development

```sh
npm install
npm run build
npm test              # typecheck + suite
npm run test:coverage # CI's gate
```

The Activity feed depends on `read_dom_list`, a fetchproxy capability added
alongside this repo — it needs `@fetchproxy/protocol`/`@fetchproxy/server`
`>= 3.1.0`, the first published versions that include it
([chrischall/fetchproxy#381](https://github.com/chrischall/fetchproxy/pull/381)).

## License

MIT
