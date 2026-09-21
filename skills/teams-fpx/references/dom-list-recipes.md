# fpx dom-list recipes for Teams

Every command needs `-p teams --storage-domain teams.microsoft.com` (or
`teams.cloud.microsoft`, matching whichever host your signed-in tab is on).

## Chat list (sidebar, Chat nav section)

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

## Teams and channels (sidebar, Teams nav section — a DIFFERENT view from Chat)

```sh
fpx dom-list teamsAndChannels -p teams --storage-domain teams.microsoft.com \
  | jq -r '.[] | select(.itemType == "channel") | "\(.teamName)\t\(.title)"'
```

Fields: `title` (team or channel name), `teamName` (the parent team's
display name — present on a channel row, absent on a team row), `time`
(relative, like `chatList.time`), `conversationKey` (internal compound id,
not usable to open the item), `itemType` (`team` or `channel`).

Needs the **Teams-and-Channels** nav section open in the browser (click
"Teams" in the left rail), not Chat — if this returns `[]`, that's the
usual reason.

## Currently open channel's posts

```sh
fpx dom-list channelPosts -p teams --storage-domain teams.microsoft.com \
  | jq -r '.[] | "\(.time)\t\(.sender)\t\(.subject // "")\t\(.text)"'
```

Fields: `sender`, `subject` (present only on a post with a title), `time`
(prose, e.g. `Tuesday, August 4, 2026 7:50 AM` — NOT ISO 8601, unlike
`chatMessages.time`), `messageId` (`timestamp-<epoch>`), `text`. Threaded
replies under a post are not captured, only the top-level posts.

Reads whatever channel is CURRENTLY OPEN — open it in the browser first.

## Multiple tabs

If more than one `teams.microsoft.com` tab is open, the bridge picks
whichever answers first — results can come from a tab you didn't mean.
Close extras before relying on output.
