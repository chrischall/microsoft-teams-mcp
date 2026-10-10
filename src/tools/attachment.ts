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
 * - a shared file lives in SharePoint, which that token cannot read (Graph
 *   /shares → 403, SharePoint direct → 401), so it comes back as metadata and
 *   its open link, saying so.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { McpToolError, toolAnnotations } from '@chrischall/mcp-utils';
import type { RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import { normaliseMessage } from '../model.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

/**
 * Raw messages read looking for the one asked for (20 pages of 200). The
 * `stopAt` boundary — paging past the message's arrival time — is the real
 * bound; this only caps an id that is not an arrival time.
 */
const MAX_RAW_MESSAGES = 4000;

const FILE_HINT = 'Shared files live in SharePoint; this MCP cannot download their contents yet — open the link.';

export function registerAttachmentTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_get_attachment',
    {
      title: 'Get a Teams message attachment',
      description:
        'Fetch one attachment of a Teams chat or channel message, as listed by ' +
        'teams_read_conversation / teams_get_unread. An inline image (attachments.images[].ref) ' +
        'is returned as the image itself. A shared file (attachments.files[].name) is returned ' +
        'as its name, type and links only: files live in SharePoint and their contents cannot ' +
        'be downloaded by this MCP yet, so the result says so and gives the link to open. ' +
        'Does not mark anything read.' +
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
          return untrustedResult({ ...file, content_available: false, hint: FILE_HINT });
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
