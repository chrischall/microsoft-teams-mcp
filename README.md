# microsoft-teams-mcp

MCP server for **Microsoft Teams** — list your chats, teams and channels, and
read whichever chat or channel is currently open, routed through your
signed-in browser tab.

> This project was developed and is maintained by AI (Claude Code). Use at your
> own discretion.

## How it works — and its one real limitation

Teams Web's chat and message data lives only in a client-side cache fed by a
WebSocket/registrar sync channel — there is no REST or GraphQL endpoint that
returns it. So this server does not capture a token and make plain server-side
requests the way an API-backed MCP would; instead every call routes through
the [fetchproxy](https://github.com/chrischall/fetchproxy) browser bridge and
reads the DOM your signed-in `teams.microsoft.com` (or `teams.cloud.microsoft`)
tab has already rendered.

**That has one real consequence: this server can only read whichever chat or
channel is currently displayed in your browser tab.** The bridge can fetch
and read — it cannot navigate a tab to a different conversation.
`teams_list_chats` and `teams_list_teams_and_channels` always work (their
sidebars are always rendered once that nav section is open);
`teams_get_open_chat_messages` and `teams_get_open_channel_posts` read
whatever conversation you have open. To read a different chat or channel,
open it in the browser first.

The bridge is on the request path for **every** call here, not just once to
mint a token — unlike an API-backed MCP, there is nothing to cache between
calls.

## Install

```sh
npm i -g @chrischall/microsoft-teams-mcp
```

Requires the **Transporter** Chrome extension and `@fetchproxy/cli`, on a
matching major version:

```sh
npm i -g @fetchproxy/cli
```

The first call prints a 6-digit pair code to approve in the Transporter
popup; the grant persists.

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

Everything is optional — this server needs no credentials of its own.

| variable | purpose |
| --- | --- |
| `TEAMS_WS_PORT` | fetchproxy concentrator port. Defaults to `37149`, the fleet-wide shared port. |

## Tools

- `teams_list_chats` — the Chat sidebar: title, last-message preview,
  last-activity time, and Teams' own conversation id for each chat.
- `teams_get_open_chat_messages` — the message thread of whichever chat is
  currently open: sender, ISO-8601 time, message id, and text.
- `teams_list_teams_and_channels` — the Teams-and-Channels sidebar (a
  DIFFERENT view from Chat): every team and channel the user belongs to,
  with each channel's parent team name.
- `teams_get_open_channel_posts` — the top-level posts of whichever channel
  is currently open: sender, subject (when present), prose time (not
  ISO-8601 — Teams doesn't expose one for these), message id, and text.
  Threaded replies under a post are not included.
- `teams_healthcheck` — verifies the bridge can reach a signed-in tab.

All four data tools are read-only and take no arguments.

## Things worth knowing

- **No chat or channel selection.** See "How it works" above — there is no
  id parameter on either "open" tool because there is no way to act on one.
  Ask the person to open the chat/channel they mean, or use the matching
  list tool to show them what's available.
- **Chat and Teams-and-Channels are different nav sections.** Reading
  channels needs the Teams-and-Channels view open in the browser (the
  "Teams" icon in the left rail, not "Chat") — if a channel tool returns
  nothing, that's usually why.
- **Grouped messages still resolve a sender.** Teams visually hides the
  sender name on consecutive messages from the same person, but the name
  stays in the DOM, so `sender` is populated on every row regardless.
- **Quoted content concatenates.** A chat message that quotes an earlier one,
  or a channel post whose body includes a reply preview, has that quoted
  text run together with the real content inside `text` — a limitation of
  bulk text extraction, not a bug.
- **Multiple open Teams tabs race.** The bridge reads whichever matching tab
  answers first ("first responsive"), not necessarily the one you meant. If
  results look like they're from the wrong chat/channel, close extra
  `teams.microsoft.com` tabs so only one remains.
- **Scope grows quietly.** If you see a `read_dom_list name not in declared
  set` error, the extension's approved scope is behind the server's declared
  one — this shouldn't happen in a released version, but if it does, revoke
  and re-pair `teams-mcp` in the Transporter popup.

## Development

```sh
npm install
npm run build
npm test              # typecheck + suite
npm run test:coverage # CI's gate
```

This server depends on `read_dom_list`, a fetchproxy capability added
alongside this repo — it needs `@fetchproxy/protocol`/`@fetchproxy/server`
`>= 3.1.0`, the first published versions that include it
([chrischall/fetchproxy#381](https://github.com/chrischall/fetchproxy/pull/381)).

## License

MIT
