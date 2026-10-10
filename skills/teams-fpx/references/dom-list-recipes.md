# fpx dom-list recipe for Teams — Activity feed only

Chats and channels are read through the chat-service API — see
`chatsvc-recipes.md`. The Activity feed has no verified API equivalent, so it
is the one thing still read off the rendered page.

Every command needs `-p teams --storage-domain teams.cloud.microsoft`.

## Activity feed (sidebar, Activity nav section — the bell icon)

```sh
fpx dom-list activityFeed -p teams --storage-domain teams.cloud.microsoft \
  | jq -r '.[] | "\(.time)\t\(.title)\t\(.location)"'
```

Fields: `title` (e.g. "Example Person mentioned Release Planning"),
`preview` (message text), `time` (prose, e.g. `1:19 PM` — NOT ISO 8601),
`location` (the team/channel or chat the activity happened in, e.g.
"Example Team > Release Planning").

It is a sidebar list — always populated once the Activity view is open,
regardless of which item is selected. If it returns `[]`, the tab is not on
the Activity view: click the bell icon in the left nav.

## Multiple tabs

If more than one `teams.cloud.microsoft` tab is open, the bridge picks
whichever answers first — results can come from a tab you didn't mean.
Close extras before relying on output.
