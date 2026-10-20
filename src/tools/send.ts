/**
 * `teams_send_message` — post a message to any chat or channel through the
 * chat-service API: plain text (escaped; newlines kept), `:key:` emoji,
 * @mentions, a quote of an earlier message, or a threaded channel reply.
 *
 * Gated by `confirmWrite` exactly as office-outlook-mcp's writes are: the
 * first call returns a preview and a token, and only the repeat call carrying
 * it posts. A post is visible to everyone in the conversation and cannot be
 * recalled by this MCP, so it is declared destructive and non-idempotent.
 *
 * After posting, the newest messages are read back and the post is looked for
 * by its `clientmessageid` — a 201 alone is not proof it landed where meant.
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
  unknownOutcomeReason,
  writeOrUnknown,
  type WriteOutcomeUnknownError,
} from '@chrischall/mcp-utils';
import type { ChatServiceClient } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import {
  buildSendBody,
  clipText,
  mentionsProperty,
  newClientMessageId,
  qtdMsgsProperty,
  quoteHtml,
  renderMessageHtml,
  type MentionTarget,
  type QuotedMessage,
} from '../compose.js';
import { htmlToText, isUserMessage, normaliseMessage, SELF_CHAT, senderMri } from '../model.js';
import { resolveListed } from './conversation.js';
import { wrapApiError } from './errors.js';
import {
  THIRD_PARTY_NAMES_NOTE,
  untrustedField,
  untrustedResult,
  UNTRUSTED_DESCRIPTION_SUFFIX,
} from './untrusted.js';

const TOOL = 'teams_send_message';

/**
 * The user's own notes chat — the one `48:*` system conversation that takes
 * messages, and the safe place to try this tool out. Every other `48:*` id
 * (calendar, notifications, …) is refused.
 */
export { SELF_CHAT };

/** Recent messages scanned for mention candidates and for the read-back. */
const RECENT_MESSAGES = 200;
const READBACK_MESSAGES = 50;

/** Candidates a mention error lists. */
const MAX_CANDIDATES = 10;

/** Previewed text of a quoted message. */
const PREVIEW_CHARS = 200;

/** An MRI such as `8:orgid:<oid>` (a person) or `28:<id>` (a bot). */
const MRI = /^\d+:\S+$/;

export interface Target {
  id: string;
  title: string;
  kind: string;
  /** An untitled group chat's participants (its title names the first three), or an untitled 1:1's other person. */
  participants?: string[];
}

const mentionArg = z.object({
  name_in_text: z
    .string()
    .min(1)
    .describe('How the person is named in text — its first "@name" (else "name") becomes the mention.'),
  person: z
    .string()
    .min(1)
    .describe('Who to mention: an MRI (8:orgid:…), or a display name matched among recent senders in this conversation.'),
});

/**
 * The conversation a write goes to: `48:notes` (also by the names `notes` /
 * `self`), else an id or name as `teams_read_conversation` resolves it. `verb` words the refusals ("post to",
 * "react in").
 */
export async function resolveTarget(
  chatsvc: ChatServiceClient,
  conversation: string,
  me: string,
  verb = 'post to',
): Promise<Target> {
  if (conversation === SELF_CHAT) return { id: SELF_CHAT, title: 'Notes (self-chat)', kind: 'self' };
  if (conversation.startsWith('48:')) {
    throw new McpToolError(`Refusing to ${verb} the Teams system conversation ${conversation}.`, {
      hint: `Only ${SELF_CHAT} (the self-chat) takes messages among 48: ids; pass a chat or channel instead.`,
    });
  }
  const raw = await chatsvc.listConversations();
  const s = await resolveListed(chatsvc, raw, me, conversation, verb);
  return { id: s.id, title: s.title, kind: s.kind, ...(s.participants !== undefined ? { participants: s.participants } : {}) };
}

/** Who has written in the conversation lately, by MRI → display name. */
async function recentSenders(chatsvc: ChatServiceClient, conversationId: string): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for (const m of await chatsvc.getMessages(conversationId, { max: RECENT_MESSAGES })) {
    const mri = senderMri(m.from);
    const name = m.imdisplayname?.trim();
    if (isUserMessage(m) && mri !== undefined && name && !out.has(mri)) out.set(mri, name);
  }
  return out;
}

/**
 * Each mention's MRI: an MRI is taken as-is; a name must match exactly one
 * recent sender (exactly, else as a substring, ignoring case). Ambiguous or
 * unknown is an error naming the candidates — never a guess.
 */
async function resolveMentions(
  chatsvc: ChatServiceClient,
  conversationId: string,
  mentions: readonly { name_in_text: string; person: string }[],
): Promise<{ target: MentionTarget; person: string }[]> {
  if (mentions.length === 0) return [];
  const senders = mentions.every((m) => MRI.test(m.person.trim()))
    ? new Map<string, string>()
    : await recentSenders(chatsvc, conversationId);
  const known = [...senders].map(([mri, name]) => ({ mri, name }));

  return mentions.map(({ name_in_text, person }) => {
    const p = person.trim();
    if (MRI.test(p)) return { target: { text: name_in_text, mri: p }, person: p };
    const q = p.toLowerCase();
    const exact = known.filter((k) => k.name.toLowerCase() === q);
    const found = exact.length > 0 ? exact : known.filter((k) => k.name.toLowerCase().includes(q));
    if (found.length === 1) return { target: { text: name_in_text, mri: found[0].mri }, person: found[0].name };
    // Candidate names are other people's display names: JSON-quoted and
    // flagged as data, as a read tool's envelope would.
    if (found.length > 1) {
      const named = found
        .slice(0, MAX_CANDIDATES)
        .map((k) => `${JSON.stringify(k.name)} (${k.mri})`)
        .join('; ');
      throw new McpToolError(
        `"${person}" matches ${found.length} people in this conversation: ${named}. ${THIRD_PARTY_NAMES_NOTE}`,
        { hint: 'Pass the MRI of the one to mention as person.' },
      );
    }
    const some = known
      .slice(0, MAX_CANDIDATES)
      .map((k) => JSON.stringify(k.name))
      .join(', ');
    throw new McpToolError(`No recent sender in this conversation is called "${person}".`, {
      hint:
        (some ? `Recent senders: ${some}. ${THIRD_PARTY_NAMES_NOTE} ` : '') +
        'Pass an exact display name, or the person\'s MRI (8:orgid:…) as person.',
    });
  });
}

/** The message a quote points at, read for its sender and text. */
async function quoted(chatsvc: ChatServiceClient, conversationId: string, id: string): Promise<QuotedMessage> {
  const m = await chatsvc.getMessage(conversationId, id);
  if (m === undefined) {
    throw new McpToolError(`There is no message ${id} in this conversation to quote.`, {
      hint: 'Take message ids from teams_read_conversation.',
    });
  }
  const mri = senderMri(m.from) ?? '';
  return { id: m.id, senderMri: mri, senderName: m.imdisplayname?.trim() || mri, text: htmlToText(m.content ?? '') };
}

/**
 * The channel root post a threaded reply goes under, read to prove it is one:
 * it must exist in this channel and not itself be a reply (Teams threads are
 * one level deep, so a reply cannot itself be replied under).
 */
async function threadRoot(chatsvc: ChatServiceClient, channelId: string, id: string, me: string): Promise<void> {
  const m = await chatsvc.getMessage(channelId, id);
  if (m === undefined) {
    throw new McpToolError(`There is no post ${id} in this channel to reply under.`, {
      hint: 'Take the root post\'s id from teams_read_conversation.',
    });
  }
  const replyTo = normaliseMessage(m, me).replyTo;
  if (replyTo !== undefined) {
    throw new McpToolError(`Message ${id} is itself a reply in this channel, not a root post.`, {
      hint: `Reply under its root post instead: reply_to "${replyTo}".`,
    });
  }
}

/**
 * A send that may or may not have landed becomes a `status: "unknown"`
 * result, never an error: an error invites a fresh preview and a second post.
 * Any other failure is a real failure and still throws (`writeOrUnknown`).
 */
function unconfirmedPost(conversation: Target) {
  return (e: WriteOutcomeUnknownError) =>
    untrustedResult({
      status: 'unknown',
      conversation,
      warning:
        `Teams did not confirm this post (${unknownOutcomeReason(e, { sent: 'the post' })}), ` +
        'so it may already be in the conversation. Check the conversation before retrying ' +
        '(teams_read_conversation); do not resend blindly.',
    });
}

export function registerSendTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    TOOL,
    {
      title: 'Send a Teams message',
      description:
        'Post a message as the signed-in user to any chat or channel, by id (from teams_list_chats / ' +
        'teams_list_teams_and_channels) or by name. text is plain text: it is escaped, newlines are kept, ' +
        'and :key: shortcodes (e.g. :smile:, :1f389_partypopper:) become Teams emoji. reply_to in a ' +
        'channel posts into that root post\'s thread; in a chat it quote-replies that message. quote ' +
        'quote-replies a message in any conversation. mentions @-mention people: name_in_text must appear ' +
        'in text, and person is an MRI or a display name matched among recent senders (ambiguous names ' +
        `are refused with the candidates). conversation "${SELF_CHAT}" (or "notes" / "self") is the user's own notes self-chat — ` +
        'use it to try this tool without anyone else seeing; other 48: system conversations are refused. ' +
        'Everyone in the conversation sees the post. The result says whether the post was found again by ' +
        'reading the conversation back (verified); status "unknown" means Teams did not confirm it — check ' +
        'the conversation before retrying. ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows the conversation and exactly what would be posted. ' +
        CONFIRM_INJECTION_RULE +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: false, destructive: true, idempotent: false, openWorld: true }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe(
            `The conversation id (or "${SELF_CHAT}" — also by the name "notes" or "self" — for the self-chat), ` +
              "or a case-insensitive substring of the chat title, channel name, team name or an untitled chat's participant names.",
          ),
        text: z.string().min(1).describe('The message, as plain text. Newlines are kept; :key: shortcodes become emoji.'),
        reply_to: z
          .string()
          .min(1)
          .optional()
          .describe('A message id. Channel: the root post to reply under (threaded). Chat: the message to quote-reply.'),
        quote: z.string().min(1).optional().describe('A message id to quote-reply, in any conversation.'),
        mentions: z.array(mentionArg).max(20).optional().describe('People to @-mention in text.'),
        confirmToken: confirmTokenParam,
      },
    },
    async ({ conversation, text, reply_to, quote, mentions, confirmToken }, ctx) => {
      try {
        const [me, meName] = await Promise.all([chatsvc.meMri(), chatsvc.meDisplayName()]);
        const target = await resolveTarget(chatsvc, conversation, me);
        const inChannel = target.kind === 'channel';

        // In a chat there is no thread to reply under, so reply_to means "quote this".
        let quoteId = quote;
        if (!inChannel && reply_to !== undefined) {
          if (quote !== undefined && quote !== reply_to) {
            throw new McpToolError('reply_to and quote name different messages, and a chat message can quote only one.', {
              hint: 'In a chat, pass just one of them.',
            });
          }
          quoteId = reply_to;
        }
        // A threaded channel reply posts to "<channelId>;messageid=<rootId>" (unverified live; see the plan).
        const postTo = inChannel && reply_to !== undefined ? `${target.id};messageid=${reply_to}` : target.id;

        if (inChannel && reply_to !== undefined) await threadRoot(chatsvc, target.id, reply_to, me);
        const q = quoteId !== undefined ? await quoted(chatsvc, target.id, quoteId) : undefined;
        const people = await resolveMentions(chatsvc, target.id, mentions ?? []);
        const rendered = renderMessageHtml(
          text,
          people.map((p) => p.target),
        );
        const content = (q ? quoteHtml(q) : '') + rendered.html;
        const properties = {
          ...(rendered.mentions.length > 0 ? { mentions: mentionsProperty(rendered.mentions) } : {}),
          ...(q ? { qtdMsgs: qtdMsgsProperty(q) } : {}),
        };

        // What is bound and previewed: the body minus the parts minted at send
        // time (clientmessageid, timestamps), which differ on every call and
        // would otherwise refuse every phase-2 call as changed.
        const draft = buildSendBody({ conversationId: postTo, meMri: me, meName, content, clientmessageid: '', now: new Date(0), ...properties });
        const { clientmessageid: _id, composetime: _c, originalarrivaltime: _o, ...bound } = draft;
        const path = `/users/ME/conversations/${encodeURIComponent(postTo)}/messages`;

        const gate = await confirmWrite(ctx, {
          tool: TOOL,
          action: 'chat.send_message',
          message: 'Review and confirm this Teams message before it is posted:',
          // The title is someone else's text: JSON-quoted so it cannot break out of its quotes.
          summary: `Post a message to ${target.kind} ${JSON.stringify(target.title)}`,
          // The signed-in user, so an approval minted under one account never posts as another.
          account: me,
          target: postTo,
          request: { method: 'POST', path, body: bound },
          // The conversation title, the quoted message and the resolved names
          // are written by other people, so each is fenced as a read tool's
          // result is; `text` is the caller's own.
          preview: {
            conversation: untrustedField(target),
            text,
            ...(inChannel && reply_to !== undefined ? { thread_root: reply_to } : {}),
            ...(people.length > 0
              ? {
                  mentions: untrustedField(
                    people.map((p, i) => ({
                      name_in_text: mentions![i].name_in_text,
                      person: p.person,
                      mri: p.target.mri,
                    })),
                  ),
                }
              : {}),
            ...(q
              ? { quoting: untrustedField({ id: q.id, sender: q.senderName, preview: clipText(q.text, PREVIEW_CHARS) }) }
              : {}),
            note: 'clientmessageid and the timestamps are filled in when the message is posted.',
          },
          confirmToken,
        });
        if (gate) return gate;

        const clientmessageid = newClientMessageId();
        const body = buildSendBody({
          conversationId: postTo,
          meMri: me,
          meName,
          content,
          clientmessageid,
          now: new Date(),
          ...properties,
        });
        const sent = await writeOrUnknown(() => chatsvc.sendMessage(postTo, body), unconfirmedPost(target));
        if (!sent.ok) return sent.result;

        // Read back: a channel lists its thread replies, so read the channel itself.
        let found: string | undefined;
        let readBackFailed = false;
        try {
          const recent = await chatsvc.getMessages(target.id, { max: READBACK_MESSAGES });
          found = recent.find((m) => m.clientmessageid === clientmessageid)?.id;
        } catch {
          readBackFailed = true;
        }
        const verified = found !== undefined;
        const message_id = sent.value.id ?? found;
        return untrustedResult({
          status: 'sent',
          conversation: target,
          ...(message_id !== undefined ? { message_id } : {}),
          verified,
          ...(verified
            ? {}
            : {
                warning: readBackFailed
                  ? 'Teams accepted the post, but reading the conversation back failed — check it before posting again.'
                  : 'Teams accepted the post, but it was not among the newest messages on reading back — check the conversation before posting again.',
              }),
        });
      } catch (err) {
        throw wrapApiError(err, 'send the message');
      }
    },
  );
}
