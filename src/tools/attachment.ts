/**
 * `teams_get_attachment` — fetch one attachment of one message, from the
 * Teams chat-service API.
 *
 * The message is re-read (newest first, stopping once paging is past it), so
 * the attachment is located in what the service says now rather than trusted
 * from a ref the caller supplies: a ref only ever selects among the URLs the
 * message itself carries, and `fetchImage` still vets the host before the
 * bearer goes anywhere.
 *
 * Two kinds, verified live:
 * - an inline image (Teams' own image store) downloads with the chat-service
 *   token and comes back as MCP image content;
 * - a shared file lives in SharePoint, downloaded through `SharePointClient`
 *   with a separate per-host SharePoint token (see `sharepoint.ts`): an image
 *   comes back as MCP image content, anything else as an MCP embedded resource
 *   carrying the raw bytes. A file the account cannot open — or the first file
 *   from a SharePoint host before its short-lived token has been minted — comes
 *   back as metadata and its open link, degrading gracefully rather than
 *   failing the tool.
 */
import type { CallToolResult, McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { McpToolError, errorKindOf, errorStatusOf, toolAnnotations } from '@chrischall/mcp-utils';
import type { RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import type { FileAttachment, NormalisedMessage } from '../model.js';
import { normaliseMessage } from '../model.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

/**
 * Raw messages read looking for the one asked for (20 pages of 200). The
 * `stopAt` boundary — paging past the message's arrival time — is the real
 * bound; this only caps an id that is not an arrival time.
 */
const MAX_RAW_MESSAGES = 4000;

/**
 * Largest shared file whose contents are downloaded and returned inline. Past
 * this the bounded read in `SharePointClient` aborts and the file degrades to
 * metadata + link — a file this big belongs behind its link, not in a result.
 */
const MAX_FILE_BYTES = 25 * 1024 * 1024;

const FILE_HINT =
  'A shared file is downloaded from SharePoint and returned as its contents (an image as an image, ' +
  'anything else as an embedded resource). The first file from a given SharePoint host may need ' +
  'opening once in Teams to mint a short-lived token; a file your account cannot open comes back as ' +
  'metadata and its link.';

export function registerAttachmentTools(
  server: McpServer,
  { chatsvc, sharepoint }: Pick<TeamsDeps, 'chatsvc' | 'sharepoint'>,
): void {
  server.registerTool(
    'teams_get_attachment',
    {
      title: 'Get a Teams message attachment',
      description:
        'Fetch one attachment of a Teams chat or channel message, as listed by ' +
        'teams_read_conversation / teams_get_unread. An inline image (attachments.images[].ref) ' +
        'is returned as the image itself. A shared file (attachments.files[].name) is downloaded ' +
        'from SharePoint and returned as its contents: an image as the image, anything else as an ' +
        'embedded resource carrying the bytes. The first file from a given SharePoint host in a ' +
        'session may need opening once in Teams to mint a short-lived token; a file your account ' +
        'cannot open, or one that could not be fetched, comes back as metadata and its open link ' +
        'instead of failing. Does not mark anything read.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe('The conversation id the message is in (from teams_read_conversation or teams_get_unread).'),
        message_id: z.string().min(1).describe('The message id, exactly as a message listing returned it.'),
        ref: z
          .string()
          .min(1)
          .describe("An inline image's ref (attachments.images[].ref) or a shared file's name (attachments.files[].name)."),
      },
    },
    async ({ conversation, message_id, ref }) => {
      try {
        // A message id IS its arrival ms, and pages run newest first: the first
        // message older than it means it was not there.
        const target = Number(message_id);
        const stopAt = Number.isFinite(target) && target > 0 ? (m: RawMessage) => Number(m.id) < target : undefined;
        const fetched = await chatsvc.getMessages(conversation, { stopAt, max: MAX_RAW_MESSAGES });
        const raw = fetched.find((m) => m.id === message_id);
        if (raw === undefined) {
          throw new McpToolError(`Message ${message_id} was not found in conversation ${conversation}.`, {
            hint: 'Pass the conversation id and message id exactly as teams_read_conversation returned them.',
          });
        }
        // Only the attachments are read, so whose message it is does not matter.
        const message = normaliseMessage(raw, '');

        const image = message.attachments.images.find((i) => i.ref === ref);
        if (image !== undefined) {
          const { mimeType, data } = await chatsvc.fetchImage(image.url);
          if (!mimeType.startsWith('image/')) {
            throw new McpToolError(`The attachment ${ref} is not an image (the image store returned ${mimeType}).`);
          }
          const caption = untrustedResult({
            ref,
            message_id,
            conversation,
            sender: message.sender,
            time: message.time,
            mimeType,
            bytes: data.length,
          });
          return {
            content: [{ type: 'image' as const, data: data.toString('base64'), mimeType }, ...caption.content],
          };
        }

        const file = message.attachments.files.find((f) => f.name === ref);
        if (file !== undefined) {
          return downloadSharedFile(sharepoint, file, message, { conversation, message_id });
        }

        const { images, files } = message.attachments;
        throw new McpToolError(`Message ${message_id} has no attachment "${ref}".`, {
          hint:
            `It has ${images.length} inline image(s) and ${files.length} shared file(s); pass an ` +
            'image ref or a file name exactly as teams_read_conversation listed it.',
        });
      } catch (err) {
        throw wrapApiError(err, 'get the attachment');
      }
    },
  );
}

/**
 * Download a shared file's contents, degrading to metadata + link on any
 * failure rather than throwing: a file that cannot be fetched (no token yet, a
 * 403, a transient error) must never fail the tool, so this swallows every
 * `downloadFile` error and returns today's metadata result with a hint saying
 * why. A file with no `url` at all (nothing to download) takes the same path.
 */
async function downloadSharedFile(
  sharepoint: Pick<TeamsDeps['sharepoint'], 'downloadFile'>,
  file: FileAttachment,
  message: NormalisedMessage,
  ids: { conversation: string; message_id: string },
): Promise<CallToolResult> {
  if (file.url === undefined) {
    return untrustedResult({ ...file, content_available: false, hint: FILE_HINT });
  }

  try {
    const { mimeType, data, size } = await sharepoint.downloadFile(file.url, MAX_FILE_BYTES);
    const caption = untrustedResult({
      name: file.name,
      ...(file.type !== undefined ? { type: file.type } : {}),
      ...(file.openUrl !== undefined ? { openUrl: file.openUrl } : {}),
      // The degrade path carries `content_available: false`; a successful
      // download must carry the matching `true` so a caller keying on the field
      // does not read a downloaded file as unavailable.
      content_available: true,
      mimeType,
      bytes: size,
      sender: message.sender,
      time: message.time,
      message_id: ids.message_id,
      conversation: ids.conversation,
    });

    if (mimeType.startsWith('image/')) {
      return { content: [{ type: 'image' as const, data: data.toString('base64'), mimeType }, ...caption.content] };
    }
    // An embedded resource carries the raw bytes; the installed
    // `@modelcontextprotocol/server` (2.3.0) accepts a `{type:'resource'}`
    // content block with a `BlobResourceContents` body ({uri, mimeType, blob}),
    // so no base64-text fallback is needed. The `uri` is the file's open link
    // (its SharePoint object url otherwise) — an identifier, not a re-fetch.
    return {
      content: [
        {
          type: 'resource' as const,
          resource: { uri: file.openUrl ?? file.url, mimeType, blob: data.toString('base64') },
        },
        ...caption.content,
      ],
    };
  } catch (err) {
    // Graceful degrade: a file that could not be fetched comes back as metadata
    // and its link, never as a tool failure (see the docblock).
    return untrustedResult({ ...file, content_available: false, hint: degradeHint(err, hostOf(file.url)) });
  }
}

/** The host of the file's url, for the degrade hint; a stand-in if it will not parse. */
function hostOf(url: string): string {
  try {
    return new URL(url).hostname;
  } catch {
    return 'the SharePoint site';
  }
}

/**
 * Why a shared file degraded to metadata + link, as a user-facing hint.
 * `SharePointClient` throws `McpToolError`s carrying `status`/`kind`/`hint`; a
 * token-capture failure arrives as a plain `Error`. Branch on structure, not
 * message text (fleet convention).
 */
function degradeHint(err: unknown, host: string): string {
  if (errorStatusOf(err) === 403) {
    return (
      `Your account cannot open this file, or it is on a SharePoint host (${host}) your captured ` +
      'token does not cover. Open it once in a signed-in Teams tab, then retry.'
    );
  }
  if (errorKindOf(err) === 'no_credential') {
    return 'SharePoint file download is disabled here (no token source). Unset TEAMS_DISABLE_FETCHPROXY, then retry.';
  }
  if (err instanceof McpToolError && err.hint !== undefined) return err.hint;
  return (
    `Couldn't mint a SharePoint token for ${host} — open the shared file once in a signed-in Teams ` +
    'tab (that browser request is where the token is captured from), then retry.'
  );
}
