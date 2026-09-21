# microsoft-teams-mcp

MCP server for **Microsoft Teams** — list your chats and read the currently
open chat's messages, routed through your signed-in browser tab.

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

**That has one real consequence: this server can only read whichever chat is
currently displayed in your browser tab.** The bridge can fetch and read —
it cannot navigate a tab to a different conversation. `teams_list_chats`
always works (the sidebar is always rendered); `teams_get_open_chat_messages`
reads whatever chat you have open. To read a different chat, open it in the
browser first.

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

- `teams_list_chats` — the chat sidebar: title, last-message preview,
  last-activity time, and Teams' own conversation id for each chat.
- `teams_get_open_chat_messages` — the message thread of whichever chat is
  currently open: sender, ISO-8601 time, message id, and text.
- `teams_healthcheck` — verifies the bridge can reach a signed-in tab.

Both chat tools are read-only and take no arguments.

## Things worth knowing

- **No chat selection.** See "How it works" above — there is no `chatId`
  parameter on `teams_get_open_chat_messages` because there is no way to act
  on one. Ask the person to open the chat they mean, or use
  `teams_list_chats` to show them what's available.
- **Grouped messages still resolve a sender.** Teams visually hides the
  sender name on consecutive messages from the same person, but the name
  stays in the DOM, so `sender` is populated on every row regardless.
- **Quoted replies concatenate.** A message that quotes an earlier one
  includes the quoted preview text inside `text`, run together with the
  reply — a limitation of bulk text extraction, not a bug.
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
versions that include it. Until those are published, development points
`node_modules/@fetchproxy/*` at a local `fetchproxy` checkout via `npm link`.

## License

MIT
