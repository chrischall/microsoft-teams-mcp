/**
 * `teams_edit_message` — replace the text of one of the signed-in user's own
 * messages, in plain text or Markdown, through the chat-service API
 * (`PUT …/messages/{id}`).
 *
 * Only the user's own, undeleted, person-written messages can be edited: the
 * tool reads the message first and refuses anything else before any preview.
 * A quote-reply keeps its quote (the quote block and `qtdMsgs`), so an edit
 * changes only what the user wrote. Gated by `confirmWrite` like every write:
 * the preview shows the current text, the new text and the exact HTML that
 * will replace it. Everyone in the conversation sees the edit, and the old
 * text cannot be brought back by this MCP, so it is declared destructive; the
 * same edit twice leaves the same message, so it is idempotent.
 *
 * After the write the message is read back: `verified` says whether it now
 * reads as the edit.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import {
  CONFIRM_FLOW_SENTENCE,
  CONFIRM_INJECTION_RULE,
  confirmTokenParam,
  confirmWrite,
  McpToolError,
  toolAnnotations,
  WriteOutcomeUnknownError,
} from '@chrischall/mcp-utils';
import type { ChatServiceClient, RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import { mentionsProperty, renderMessageHtml } from '../compose.js';
import { htmlToText, isUserMessage, SELF_CHAT, senderMri } from '../model.js';
import { wrapApiError } from './errors.js';
import { formatParam, mentionArg, resolveMentions, resolveTarget } from './send.js';
import { untrustedField, untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';
import { workMeta } from './hints.js';

const TOOL = 'teams_edit_message';

/** A quote-reply's quote block, as Teams (and `compose.quoteHtml`) put it at the start of the body. */
const QUOTE_BLOCK = /^\s*(<blockquote\b[^>]*schema\.skype\.com\/Reply[^>]*>[\s\S]*?<\/blockquote\s*>)/i;

/** The quote block a quote-reply starts with, kept through an edit; `''` for none. */
export function quotePrefix(content: string): string {
  return QUOTE_BLOCK.exec(content)?.[1] ?? '';
}

/** Whether `m` reads as the edit: its text is the text of `content`. */
const readsAs = (m: RawMessage | undefined, content: string): boolean =>
  m !== undefined && htmlToText(m.content ?? '') === htmlToText(content);

/** The message to edit, or the reason it cannot be edited. */
async function ownMessage(chatsvc: ChatServiceClient, conversationId: string, id: string, me: string): Promise<RawMessage> {
  const m = await chatsvc.getMessage(conversationId, id);
  if (m === undefined) {
    throw new McpToolError(`There is no message ${id} in this conversation to edit.`, {
      hint: 'Take message ids from teams_list_messages.',
    });
  }
  if (senderMri(m.from)?.toLowerCase() !== me.toLowerCase()) {
    throw new McpToolError(`Message ${id} was not sent by you: you can edit only your own messages.`, {
      hint: 'Pick one of your own messages (fromMe: true in teams_list_messages), or send a new message instead.',
    });
  }
  const del = m.properties?.deletetime;
  if (del !== undefined && del !== null && del !== '' && del !== 0) {
    throw new McpToolError(`Message ${id} has been deleted, so it cannot be edited.`, {
      hint: 'Send a new message instead (teams_send_message).',
    });
  }
  if (!isUserMessage(m)) {
    throw new McpToolError(`Message ${id} is a system event (${m.messagetype}), not a message anyone wrote.`, {
      hint: 'Only messages you wrote can be edited.',
    });
  }
  return m;
}

export function registerEditTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    TOOL,
    {
      title: 'Edit a Teams message',
      description:
        'Replace the text of one of your own messages in any chat or channel. conversation is an id or name as ' +
        'for teams_list_messages; message_id comes from teams_list_messages and must be a message you sent ' +
        "(someone else's, a deleted one or a system event is refused). text replaces the whole message: plain " +
        'text by default, or Markdown with format "markdown" (as teams_send_message reads it). mentions ' +
        '@-mention people as in teams_send_message. A quote-reply keeps its quote. Text that renders exactly as ' +
        'the message already does changes nothing (status "unchanged"). Everyone in the conversation sees the ' +
        'message marked edited. The result says whether reading the message back shows the edit (verified); ' +
        'status "unknown" means Teams did not confirm it and the read-back did not show it — check the message ' +
        `before retrying. conversation "${SELF_CHAT}" (or "notes" / "self") is the user's own notes self-chat — ` +
        'use it to try this tool. ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows the conversation, the current text, the new text and exactly what will replace it. ' +
        CONFIRM_INJECTION_RULE +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: false, destructive: true, idempotent: true, openWorld: true }),
      _meta: workMeta('change', { aliases: ['edit', 'edit message', 'fix my message', 'correct', 'reword', 'update message'] }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe(
            `The conversation id (or "${SELF_CHAT}" — also by the name "notes" or "self" — for the self-chat), ` +
              "or a case-insensitive substring of the chat title, channel name, team name or an untitled chat's participant names.",
          ),
        message_id: z.string().min(1).describe('The message to edit — one you sent — as teams_list_messages lists it.'),
        text: z
          .string()
          .min(1)
          .describe('The whole new message: plain text (newlines kept, :key: shortcodes become emoji), or Markdown with format "markdown".'),
        format: formatParam,
        mentions: z.array(mentionArg).max(20).optional().describe('People to @-mention in text.'),
        confirmToken: confirmTokenParam,
      },
    },
    async ({ conversation, message_id, text, format = 'text', mentions, confirmToken }, ctx) => {
      try {
        const me = await chatsvc.meMri();
        const target = await resolveTarget(chatsvc, conversation, me, 'edit a message in');
        const original = await ownMessage(chatsvc, target.id, message_id, me);
        const before = original.content ?? '';

        const people = await resolveMentions(chatsvc, target.id, mentions ?? []);
        const rendered = renderMessageHtml(
          text,
          people.map((p) => p.target),
          format,
        );
        const content = quotePrefix(before) + rendered.html;
        if (content === before) {
          return untrustedResult({ status: 'unchanged', conversation: target, message_id });
        }

        const qtdMsgs = original.properties?.qtdMsgs;
        const body = {
          content,
          messagetype: 'RichText/Html',
          contenttype: 'text',
          amsreferences: [] as string[],
          properties: {
            mentions: rendered.mentions.length > 0 ? mentionsProperty(rendered.mentions) : '[]',
            ...(typeof qtdMsgs === 'string' && quotePrefix(before) !== '' ? { qtdMsgs } : {}),
          },
        };
        const path = `/users/ME/conversations/${encodeURIComponent(target.id)}/messages/${encodeURIComponent(message_id)}`;

        const gate = await confirmWrite(ctx, {
          tool: TOOL,
          action: 'chat.edit_message',
          message: 'Review and confirm this edit before it replaces your Teams message:',
          // The title is someone else's text: JSON-quoted so it cannot break out of its quotes.
          summary: `Edit your message ${message_id} in ${target.kind} ${JSON.stringify(target.title)}`,
          account: me,
          target: `${target.id}/${message_id}`,
          request: { method: 'PUT', path, body },
          preview: {
            conversation: untrustedField(target),
            message_id,
            current_text: htmlToText(before),
            text,
            format,
            ...(people.length > 0
              ? {
                  mentions: untrustedField(
                    people.map((p, i) => ({ name_in_text: mentions![i].name_in_text, person: p.person, mri: p.target.mri })),
                  ),
                }
              : {}),
          },
          confirmToken,
        });
        if (gate) return gate;

        let unconfirmed: WriteOutcomeUnknownError | undefined;
        try {
          await chatsvc.editMessage(target.id, message_id, body);
        } catch (e) {
          if (!(e instanceof WriteOutcomeUnknownError)) throw e;
          unconfirmed = e;
        }

        let verified = false;
        let readBackFailed = false;
        try {
          verified = readsAs(await chatsvc.getMessage(target.id, message_id), content);
        } catch {
          readBackFailed = true;
        }

        if (unconfirmed !== undefined && !verified) {
          return untrustedResult({
            status: 'unknown',
            conversation: target,
            message_id,
            verified,
            warning:
              `Teams did not confirm this edit, and reading the message back ${readBackFailed ? 'failed' : 'did not show it'} — ` +
              'check the message (teams_list_messages) before retrying.',
          });
        }
        const warning =
          unconfirmed !== undefined
            ? 'Teams did not confirm this edit, but reading the message back shows it.'
            : readBackFailed
              ? 'Teams accepted the edit, but reading the message back failed — check the message (teams_list_messages).'
              : !verified
                ? 'Teams accepted the edit, but reading the message back did not show the new text — check the message (teams_list_messages).'
                : undefined;
        return untrustedResult({
          status: 'edited',
          conversation: target,
          message_id,
          verified,
          ...(warning !== undefined ? { warning } : {}),
        });
      } catch (err) {
        throw wrapApiError(err, 'edit the message');
      }
    },
  );
}
