---
name: teams-fpx
description: >-
  Read Microsoft Teams chat data (chat list, and the currently open chat's
  messages) from a shell with the fpx CLI (@fetchproxy/cli) instead of running
  the microsoft-teams-mcp server. Teams' chat data lives only in a
  client-side sync cache with no REST/GraphQL endpoint, so this reads the
  rendered DOM of a signed-in browser tab via fpx's read_dom_list capability
  — no login, no token, just an open tab. Use when you want Teams chat data
  without the MCP, in a script, or on a machine where the MCP isn't
  installed.
---

# Microsoft Teams via fpx (no MCP)

Teams Web has no fetchable API for chat/message content — it's synced over a
private WebSocket channel into a client-side cache, and the web app renders
straight from that cache. There is nothing to `curl` or POST a GraphQL query
to. The only way to read it is to read what the page has already rendered,
which is what `fpx dom-list` does: it asks the Transporter extension to run a
declared `querySelectorAll` + per-item field extraction against your
signed-in `teams.microsoft.com` tab and hands back structured rows.

**This means the same limitation the MCP has applies here: you can only read
whichever chat is currently open in the browser.** There is no way to switch
chats from the CLI — open the one you want in Teams first.

## One-time setup

```sh
npm install -g @fetchproxy/cli   # provides `fpx`, needs 2.1.0+ (read_dom_list)
fpx profile add teams --domain teams.microsoft.com --domain teams.cloud.microsoft
fpx profile declare teams --dom-list-selector \
  'chatList=[data-testid="list-item"]::title:[id^="title-chat-list-item_"],preview:[data-testid="comfy-message"],time:[id^="time-chat-list-item_"],conversationKey:@data-fui-tree-item-value,itemType:@data-item-type&max=200'
fpx profile declare teams --dom-list-selector \
  'chatMessages=.fui-ChatMessage, .fui-ChatMyMessage::sender:[data-tid="message-author-name"],time:time@datetime,messageId:time@id,text:[data-tid="chat-pane-message"]&max=200'
```

Both `--dom-list-selector` flags declare a DIFFERENT named selector on the
same profile — run both, not one. Requirements: the **Transporter** browser
extension installed, with an open `teams.microsoft.com` (or
`teams.cloud.microsoft`) tab you're signed into. The first `dom-list` call
prints a pair code — approve it in the Transporter popup; the grant persists
until the declared scope changes again.

## Core calls

```sh
fpx dom-list chatList -p teams --storage-domain teams.microsoft.com | jq
fpx dom-list chatMessages -p teams --storage-domain teams.microsoft.com | jq
```

`--storage-domain` is required whenever a profile declares more than one
domain (this one does) — it tells fpx which declared domain's tab to target.
Use `teams.cloud.microsoft` instead if that's where your signed-in tab lives.

Ready-to-run recipes (jq filters for both selectors) are in
`references/dom-list-recipes.md`.

## Exit codes

- `0` — success.
- `2` — bridge unavailable: extension not connected, or pairing pending →
  run the command again and approve the printed pair code, or check
  `fpx health -p teams`.
- Error text containing `not in declared set` — the profile's declared scope
  changed since you last paired. Revoke `fpx-teams` in the Transporter popup
  and re-run; you'll be asked to approve the new scope.

## Notes

- Read-only. Nothing here can send a message, open a chat, or interact with
  the page — `read_dom_list` only reads what's already rendered.
- Sender names resolve even on consecutive messages from the same person,
  where Teams visually hides the name — it's still in the DOM.
- A quoted-reply message's `text` field includes the quoted preview
  concatenated in, run together with the reply text.
- This project is developed and maintained by AI (Claude).
