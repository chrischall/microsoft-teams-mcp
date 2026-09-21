---
name: teams-fpx
description: >-
  Read Microsoft Teams data (chats, teams, channels, the Activity feed, and
  whichever chat or channel is currently open) from a shell with the fpx CLI
  (@fetchproxy/cli)
  instead of running the microsoft-teams-mcp server. Teams' data lives only
  in a client-side sync cache with no REST/GraphQL endpoint, so this reads
  the rendered DOM of a signed-in browser tab via fpx's read_dom_list
  capability — no login, no token, just an open tab. Use when you want Teams
  data without the MCP, in a script, or on a machine where the MCP isn't
  installed.
---

# Microsoft Teams via fpx (no MCP)

Teams Web has no fetchable API for chat/message/channel content — it's
synced over a private WebSocket channel into a client-side cache, and the
web app renders straight from that cache. There is nothing to `curl` or POST
a GraphQL query to. The only way to read it is to read what the page has
already rendered, which is what `fpx dom-list` does: it asks the Transporter
extension to run a declared `querySelectorAll` + per-item field extraction
against your signed-in `teams.cloud.microsoft` tab and hands back structured
rows. `teams.microsoft.com` (the old host) is not supported.

**This means the same limitation the MCP has applies here: you can only read
whichever chat or channel is currently open in the browser.** There is no
way to switch conversations from the CLI — open the one you want in Teams
first. Reading channels also needs the **Teams-and-Channels** nav section
open (the "Teams" icon in the left rail), not Chat. The Activity feed
doesn't have this limitation — it's a sidebar list, always rendered once
that nav section (the bell icon) is open, regardless of which item is
selected.

**No Calendar selector.** Teams Web's Calendar renders inside an embedded
`outlook.office.com` iframe — `read_dom_list` only ever reads the top-level
tab's DOM, so it's unreachable from here (and from the MCP) no matter what.
Use the `office-outlook-mcp` repo's `outlook-fpx` skill or MCP instead —
Outlook's calendar is a real REST API, not a DOM read.

## One-time setup

```sh
npm install -g @fetchproxy/cli   # provides `fpx`, needs 2.1.0+ (read_dom_list)
fpx profile add teams --domain teams.cloud.microsoft
fpx profile declare teams --dom-list-selector \
  'chatList=[data-testid="list-item"]::title:[id^="title-chat-list-item_"],preview:[data-testid="comfy-message"],time:[id^="time-chat-list-item_"],conversationKey:@data-fui-tree-item-value,itemType:@data-item-type&max=200'
fpx profile declare teams --dom-list-selector \
  'chatMessages=.fui-ChatMessage, .fui-ChatMyMessage::sender:[data-tid="message-author-name"],time:time@datetime,messageId:time@id,text:[data-tid="chat-pane-message"]&max=200'
fpx profile declare teams --dom-list-selector \
  'teamsAndChannels=[data-item-type="team"], [data-item-type="channel"]::title:[id^="title-"][id*="-list-item-"],teamName:[id^="preview-channel-list-item-"],time:[id^="time-"][id*="-list-item-"],conversationKey:@data-fui-tree-item-value,itemType:@data-item-type&max=300'
fpx profile declare teams --dom-list-selector \
  'channelPosts=[data-tid="channel-pane-message"]::sender:[id^="author-"],subject:[id^="subject-line-"],time:time[data-tid="timestamp"]@aria-label,messageId:time[data-tid="timestamp"]@id,text:[id^="content-"]&max=200'
fpx profile declare teams --dom-list-selector \
  'activityFeed=[data-tid="activity-feed-list-item"]::title:[id^="activity-feed-item-title-"],preview:[id^="activity-feed-item-message-preview-"],time:[id^="activity-feed-item-timestamp-"],location:[id^="activity-feed-item-location-"]&max=200'
```

Each `--dom-list-selector` flag declares a DIFFERENT named selector on the
same profile — run all five, not just one. Requirements: the
**Transporter** browser extension installed, with an open
`teams.cloud.microsoft` tab you're signed into. The first `dom-list` call
prints a pair code — approve it in the Transporter popup; the grant persists
until the declared scope changes again (each `declare` above widens it
once).

## Core calls

```sh
fpx dom-list chatList -p teams --storage-domain teams.cloud.microsoft | jq
fpx dom-list chatMessages -p teams --storage-domain teams.cloud.microsoft | jq
fpx dom-list teamsAndChannels -p teams --storage-domain teams.cloud.microsoft | jq
fpx dom-list channelPosts -p teams --storage-domain teams.cloud.microsoft | jq
fpx dom-list activityFeed -p teams --storage-domain teams.cloud.microsoft | jq
```

Ready-to-run recipes (jq filters for all five selectors) are in
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

- Read-only. Nothing here can send a message, open a chat/channel, or
  interact with the page — `read_dom_list` only reads what's already
  rendered.
- Sender names resolve even on consecutive messages from the same person,
  where Teams visually hides the name — it's still in the DOM.
- A quoted chat reply's `text` field includes the quoted preview
  concatenated in; same for a channel post that includes a reply preview.
- `channelPosts.time` is prose ("Tuesday, August 4, 2026 7:50 AM"), not
  ISO 8601 — Teams' channel-post timestamp carries no `datetime` attribute,
  unlike `chatMessages.time`.
- If your browser has more than one `teams.cloud.microsoft` tab open, the
  bridge reads whichever answers first, not necessarily the one you meant —
  close extras.
- This project is developed and maintained by AI (Claude).
