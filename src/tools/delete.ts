/**
 * `teams_delete_message` — delete one of the signed-in user's own messages
 * through the chat-service API (`DELETE …/messages/{id}`).
 *
 * Only the user's own, person-written messages can be deleted: the tool reads
 * the message first and refuses anything else before any preview. One that
 * is already deleted is reported as such, with nothing asked or sent. Gated
 * by `confirmWrite` like every write: the preview shows the message that will
 * go. The text cannot be brought back by this MCP, so it is declared
 * destructive; deleting twice leaves it deleted, so it is idempotent.
 *
 * After the write the message is read back: `verified` says whether it now
 * reads as deleted (Teams keeps it with `deletetime` set; a 404 counts too).
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import {
  CONFIRM_FLOW_SENTENCE,
  CONFIRM_INJECTION_RULE,
  confirmTokenParam,
  confirmWrite,
  toolAnnotations,
} from '@chrischall/mcp-utils';
import type { TeamsDeps } from '../deps.js';
import { htmlToText, isDeleted, normaliseMessage, SELF_CHAT } from '../model.js';
import { ownMessage, writeAndReadBack } from './own-message.js';
import { wrapApiError } from './errors.js';
import { resolveTarget } from './send.js';
import { untrustedField, untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';
import { workMeta } from './hints.js';

const TOOL = 'teams_delete_message';

export function registerDeleteTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    TOOL,
    {
      title: 'Delete a Teams message',
      description:
        'Delete one of your own messages in any chat or channel. conversation is an id or name as for ' +
        'teams_list_messages; message_id comes from teams_list_messages and must be a message you sent ' +
        "(someone else's or a system event is refused). One already deleted is reported as status " +
        '"already_deleted" with nothing asked or sent. Everyone in the conversation sees that the message was ' +
        'deleted, and this MCP cannot bring it back. The result says whether reading the message back shows it ' +
        'deleted (verified); status "unknown" means Teams did not confirm it and the read-back did not show it ' +
        `deleted — check before retrying. conversation "${SELF_CHAT}" (or "notes" / "self") is the user's own ` +
        'notes self-chat — use it to try this tool. ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows the conversation and the message that will be deleted. ' +
        CONFIRM_INJECTION_RULE +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: false, destructive: true, idempotent: true, openWorld: true }),
      _meta: workMeta('change', { aliases: ['delete', 'delete message', 'remove message', 'unsend', 'take back'] }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe(
            `The conversation id (or "${SELF_CHAT}" — also by the name "notes" or "self" — for the self-chat), ` +
              "or a case-insensitive substring of the chat title, channel name, team name or an untitled chat's participant names.",
          ),
        message_id: z.string().min(1).describe('The message to delete — one you sent — as teams_list_messages lists it.'),
        confirmToken: confirmTokenParam,
      },
    },
    async ({ conversation, message_id, confirmToken }, ctx) => {
      try {
        const me = await chatsvc.meMri();
        const target = await resolveTarget(chatsvc, conversation, me, 'delete a message in');
        const m = await ownMessage(chatsvc, target.id, message_id, me, 'delete');
        if (isDeleted(m)) {
          return untrustedResult({ status: 'already_deleted', conversation: target, message_id });
        }

        const path = `/users/ME/conversations/${encodeURIComponent(target.id)}/messages/${encodeURIComponent(message_id)}`;
        const gate = await confirmWrite(ctx, {
          tool: TOOL,
          action: 'chat.delete_message',
          message: 'Review and confirm before this Teams message of yours is deleted:',
          // The title is someone else's text: JSON-quoted so it cannot break out of its quotes.
          summary: `Delete your message ${message_id} in ${target.kind} ${JSON.stringify(target.title)}`,
          account: me,
          target: `${target.id}/${message_id}`,
          request: { method: 'DELETE', path },
          preview: {
            conversation: untrustedField(target),
            message_id,
            time: normaliseMessage(m, me).time,
            // The user's own text, but a quote-reply's quote is someone else's.
            text: untrustedField(htmlToText(m.content ?? '')),
          },
          confirmToken,
        });
        if (gate) return gate;

        const { status, ...settled } = await writeAndReadBack(
          () => chatsvc.deleteMessage(target.id, message_id),
          () => chatsvc.getMessage(target.id, message_id),
          (back) => back === undefined || isDeleted(back),
          'delete',
          'it deleted',
        );
        return untrustedResult({ status: status === 'done' ? 'deleted' : 'unknown', conversation: target, message_id, ...settled });
      } catch (err) {
        throw wrapApiError(err, 'delete the message');
      }
    },
  );
}
