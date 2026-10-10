# microsoft-teams-mcp

MCP server for **Microsoft Teams** — list every chat and channel you are in,
read any of them by id or name, pull everything you have not read yet, and
fetch inline images, using your signed-in browser session.

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
your Teams unread badges exactly as they were.

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
| `TEAMS_CAPTURE_TIMEOUT` | Seconds to wait for the Teams tab to issue a chat-service request. Defaults to `30`. |
| `TEAMS_TOKEN_CACHE` | `false` to keep the captured token in memory only. Defaults to on. |
| `TEAMS_TOKEN_FILE` | Where the captured token is cached. Defaults to `~/.microsoft-teams-mcp/token.json`. |
| `TEAMS_WS_PORT` | fetchproxy concentrator port. Defaults to `37149`, the fleet-wide shared port. |
| `FETCHPROXY_WS_HOST` | Bind interface for the bridge (default `127.0.0.1`; a literal IP only, anything else is ignored). Keep on loopback — it exists for ONE topology: a hosted child in its own network namespace. |
| `FETCHPROXY_IDENTITY_DIR` | Where to store the long-term fetchproxy identity keypair (default `~/.fetchproxy/identity/`; an absolute path only, anything else is ignored). Exists for a host that runs a separate child per caller, so they can share one identity. |

## Tools

- `teams_list_chats` — every chat you are in (1:1, group and meeting chats,
  not just the visible sidebar), most recent activity first: id, kind,
  title, unread flag, ISO-8601 last-activity time, last-message preview and
  sender. `unread_only`, `query` (title) and `limit` (default 50, max 500).
- `teams_list_teams_and_channels` — every channel you are in, by team then
  title, with the same fields. `unread_only`, `query` (channel or team name)
  and `limit` (default 50, max 500).
- `teams_read_conversation` — read any chat or channel by id or name,
  oldest first: sender, ISO-8601 time, text, subject, edited/deleted flags
  and attachments. A channel comes back as threads (root post + replies).
  `unread_only`, `since` (ISO 8601) and `limit` (default 50, max 500, newest
  kept). An ambiguous name lists the candidates and asks for the id.
- `teams_get_unread` — every unread message across your chats and channels,
  grouped by conversation, most recent first. Filter by `kinds` (`chat`,
  `channel`, `meeting`, `oneOnOne`) or `query`; `per_conversation_limit`
  (default 50) and `max_conversations` (default 20 — the rest are named in
  `more_unread_conversations`).
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
  came from and when it expires — never the token — plus the bridge status
  the capture and the Activity feed rely on.

All tools are read-only, and none of them marks anything read.

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
- **Unread means what Teams means.** A message is unread when it arrived
  after the conversation's read marker (`consumptionhorizon`) and was not
  sent by you. This server only reads that marker; it never moves it.
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
