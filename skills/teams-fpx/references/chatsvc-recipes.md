# Teams chat-service recipes

All read-only. Paths are relative to the chat-service base
(`https://teams.cloud.microsoft/api/chatsvc/amer/v1`; set
`TEAMS_CHATSVC_REGION` or `TEAMS_CHATSVC_BASE` before sourcing
`teams-env.sh` for another region). Run `source references/teams-env.sh` and
`teams_token_refresh` first.

## List every chat and channel

```sh
teams_get '/users/ME/conversations?view=msnp24Equivalent&pageSize=200' \
  | jq -r '.conversations[]
      | select(.id | startswith("48:") | not)
      | [.id, (.threadProperties.topic // "(untitled)"), (.lastMessage.originalarrivaltime // "")]
      | @tsv'
```

Response: `{ conversations: [...], _metadata: { backwardLink, forwardLink,
syncState, totalCount } }`. For more, pass `_metadata.backwardLink` to
`teams_get_url` until it is absent or a page is empty.

Tell the kinds apart by id suffix:

| suffix | kind |
| --- | --- |
| `@thread.tacv2` | channel (`threadProperties.spaceId` is its team's id) |
| `@thread.v2` | group or meeting chat (`19:meeting_…` ids are meetings) |
| `@unq.gbl.spaces` | 1:1 chat (no topic — name it after the other person) |
| `48:*` | system thread (notes, calendar, …) — not a chat; skip |

## Which conversations are unread

```sh
teams_get '/users/ME/conversations?view=msnp24Equivalent&pageSize=200' \
  | jq -r '.conversations[]
      | select(.id | startswith("48:") | not)
      | ((.properties.consumptionhorizon // "0") | split(";")[0] | tonumber? // 0) as $read
      | select((.lastMessage.id // "0" | tonumber? // 0) > $read)
      | [.id, (.threadProperties.topic // .lastMessage.imdisplayname // "(untitled)")] | @tsv'
```

`properties.consumptionhorizon` is `"<lastReadArrivalMs>;<clientMs>;<messageId>"`.
A message is unread when its `id` (arrival epoch ms) is greater than the first
segment. This rough filter also counts a conversation whose newest message is
your own; drop those by comparing `.lastMessage.from` (a URL ending in the
sender's MRI) with your MRI, `8:orgid:<oid>` from the token's `oid` claim.

## Read one conversation

```sh
CONV='19:example0001@thread.v2'
teams_get "/users/ME/conversations/$(jq -rn --arg v "$CONV" '$v|@uri')/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=200&startTime=1" \
  | jq -r '.messages
      | map(select(.messagetype == "Text" or .messagetype == "RichText/Html"))
      | reverse[]
      | [.originalarrivaltime, .imdisplayname, (.content | gsub("<[^>]+>"; ""))] | @tsv'
```

The id must be URL-encoded (it contains `:` and `@`). Messages come newest
first; older pages via `_metadata.backwardLink`. Other `messagetype`s
(`ThreadActivity/*`, `Event/Call`, …) are system events. A message whose
`properties.deletetime` is set was deleted.

In a **channel**, a reply's `conversationLink` ends `;messageid=<rootId>`; a
root post's root id is its own `id`. Group by that to rebuild threads.

## Attachments

Inline images are `<img src="https://us-api.asm.skype.com/v1/objects/<id>/views/imgo">`
in `content` (the host varies by region):

```sh
teams_get_image 'https://us-api.asm.skype.com/v1/objects/<id>/views/imgo' /tmp/image.jpg
```

Shared files are in `properties.files` — a JSON **string**:

```sh
... | jq -r '.messages[] | select(.properties.files) | .properties.files | fromjson[] | [.fileName, .openUrl] | @tsv'
```

They live in SharePoint and **cannot** be downloaded with the chat-service
token (Graph `/shares` → 403, SharePoint direct → 401) — use the `openUrl`.

## Never

Do not `PUT …/properties?name=consumptionhorizon` — that is how Teams marks a
conversation read, and it would clear the user's unread badges.
