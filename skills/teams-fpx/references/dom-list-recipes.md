# fpx dom-list recipes for Teams

Both commands need `-p teams --storage-domain teams.microsoft.com` (or
`teams.cloud.microsoft`, matching whichever host your signed-in tab is on).

## Chat list (sidebar)

```sh
fpx dom-list chatList -p teams --storage-domain teams.microsoft.com \
  | jq -r '.[] | "\(.time)\t\(.title)\t\(.preview)"'
```

Fields: `title`, `preview` (last-message text), `time` (relative, e.g. `9:02
AM` or `9/18` — Teams renders it that way, not as an ISO timestamp),
`conversationKey` (Teams' internal compound thread id — not usable to open
the chat; there is no navigate capability), `itemType` (`chat` for every row
in the Chat rail).

## Currently open chat's messages

```sh
fpx dom-list chatMessages -p teams --storage-domain teams.microsoft.com \
  | jq -r '.[] | "\(.time)\t\(.sender // "(same as above)")\t\(.text)"'
```

Fields: `sender`, `time` (ISO 8601, from the `<time datetime>` attribute),
`messageId` (`timestamp-<epoch>`, stable per message), `text`.

Reads whatever chat is CURRENTLY OPEN in the tab — open the one you want in
the browser first, then run this.

## Count unread-looking chats (heuristic)

`chatList` doesn't carry an explicit unread flag, but Teams renders an
unread chat's preview in bold — not something `read_dom_list` captures (it
reads text/attributes, not computed style). Use `teams_list_chats` /
`chatList` for content, not read state.
