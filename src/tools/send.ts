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
  MESSAGE_FORMATS,
  mentionsProperty,
  newClientMessageId,
  qtdMsgsProperty,
  quoteHtml,
  renderMessageHtml,
  type MentionTarget,
  type QuotedMessage,
} from '../compose.js';
import { htmlToText, localIsoTime, normaliseMessage, SELF_CHAT, senderMri } from '../model.js';
import { resolveListed } from './conversation.js';
import { wrapApiError } from './errors.js';
import { MRI, recentSenders } from './people.js';
import {
  THIRD_PARTY_NAMES_NOTE,
  untrustedField,
  untrustedResult,
  UNTRUSTED_DESCRIPTION_SUFFIX,
} from './untrusted.js';
import { workMeta } from './hints.js';
import { resolveTimeZone } from './timezone.js';

const TOOL = 'teams_send_message';

/**
 * The user's own notes chat — the one `48:*` system conversation that takes
 * messages, and the safe place to try this tool out. Every other `48:*` id
 * (calendar, notifications, …) is refused.
 */
export { SELF_CHAT };

/** Recent messages scanned for the read-back. */
const READBACK_MESSAGES = 50;

/** Candidates a mention error lists. */
const MAX_CANDIDATES = 10;

/** Previewed text of a quoted message. */
const PREVIEW_CHARS = 200;

/** Quoted text carried in the preview's one-line `quoting`. */
const QUOTING_LINE_CHARS = 80;

/** How `text` is read: shared by `teams_send_message` and `teams_edit_message`. */
export const formatParam = z
  .enum(MESSAGE_FORMATS)
  .optional()
  .describe(
    'How text is written (default "text"). text: plain text, escaped. markdown: **bold**, *italic*, ~~strike~~, ' +
      '`code`, fenced code blocks, - and 1. lists (nested by indenting), [text](https://…) links; nothing else is markup.',
  );

export interface Target {
  id: string;
  title: string;
  kind: string;
  /** An untitled group chat's participants (its title names the first three), or an untitled 1:1's other person. */
  participants?: string[];
}

export const mentionArg = z.object({
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
 * `self`), else an id or name as `teams_list_messages` resolves it. `verb` words the refusals ("post to",
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

/**
 * Each mention's MRI: an MRI is taken as-is; a name must match exactly one
 * recent sender (exactly, else as a substring, ignoring case). Ambiguous or
 * unknown is an error naming the candidates — never a guess.
 */
export async function resolveMentions(
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
async function quoted(
  chatsvc: ChatServiceClient,
  conversationId: string,
  id: string,
): Promise<QuotedMessage & { time?: string }> {
  const m = await chatsvc.getMessage(conversationId, id);
  if (m === undefined) {
    throw new McpToolError(`There is no message ${id} in this conversation to quote.`, {
      hint: 'Take message ids from teams_list_messages.',
    });
  }
  const mri = senderMri(m.from) ?? '';
  const time = m.originalarrivaltime ?? m.composetime;
  return {
    id: m.id,
    senderMri: mri,
    senderName: m.imdisplayname?.trim() || mri,
    text: htmlToText(m.content ?? ''),
    ...(time ? { time } : {}),
  };
}

/**
 * When a quoted message was sent, to the minute: in TEAMS_TIME_ZONE when one is
 * set ("2026-11-12 12:31 America/New_York"), else UTC. A bad TEAMS_TIME_ZONE
 * falls back to UTC here — it must not stop a preview.
 */
function quotedWhen(iso: string | undefined): string | undefined {
  if (iso === undefined || Number.isNaN(Date.parse(iso))) return undefined;
  let zone: string | undefined;
  try {
    zone = resolveTimeZone(undefined);
  } catch {
    zone = undefined;
  }
  const local = zone !== undefined ? localIsoTime(iso, zone) : undefined;
  if (zone !== undefined && local !== undefined) return `${local.slice(0, 16).replace('T', ' ')} ${zone}`;
  return `${new Date(iso).toISOString().slice(0, 16).replace('T', ' ')} UTC`;
}

/**
 * The preview's plain-words `quoting` line: `none`, or the quoted sender, when
 * and the start of what they wrote. Sender and text are other people's, so each
 * is JSON-quoted (on one line, its own quotes escaped) as the title is in the summary.
 */
function quotingLine(q: (QuotedMessage & { time?: string }) | undefined): string {
  if (q === undefined) return 'none';
  const when = quotedWhen(q.time);
  return `${JSON.stringify(q.senderName)}${when ? `, ${when}` : ''}: ${JSON.stringify(clipText(q.text, QUOTING_LINE_CHARS, '…'))}`;
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
      hint: 'Take the root post\'s id from teams_list_messages.',
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
        '(teams_list_messages); do not resend blindly.',
    });
}

export function registerSendTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    TOOL,
    {
      title: 'Send a Teams message',
      description:
        'Post a message as the signed-in user to any chat or channel, by id (from teams_list_conversations / ' +
        'teams_list_channels) or by name. text is plain text by default: it is escaped, newlines are kept, ' +
        'and :key: shortcodes become emoji: Teams\' animated ones for like, heart, laugh, surprised, sad, angry, smile, ' +
        'wink, cool, hearteyes, fire and party; the plain character for common names (:rocket:, :tada:, :thumbsup:, ' +
        ':eyes:, :white_check_mark:, …) and unicode-style keys (:1f680_rocket:). format "markdown" ' +
        'reads text as Markdown instead: **bold**, *italic*, ~~strike~~, `inline code`, ``` code blocks ```, ' +
        'bulleted (- item) and numbered (1. item) lists, [link text](https://…), blank-line paragraphs and ' +
        'line breaks; raw HTML, headings, quotes and images are escaped and shown as written. By default ' +
        'nothing is quoted. reply_to in a channel posts into that root post\'s thread (a thread reply, no ' +
        'quote); in a chat, reply_to and quote are the same thing: both quote-reply that message, so the ' +
        'post opens with a quote of it. quote quote-replies a message in any conversation. Pass reply_to ' +
        'or quote only when the user asked for a reply or a quote; when the request is ambiguous (they ' +
        'mention a message\'s timestamp, or say "reply" in a chat), ask them whether they want it quoted ' +
        'rather than guessing. mentions @-mention people: name_in_text must appear ' +
        'in text, and person is an MRI or a display name matched among recent senders (ambiguous names ' +
        `are refused with the candidates). conversation "${SELF_CHAT}" (or "notes" / "self") is the user's own notes self-chat — ` +
        'use it to try this tool without anyone else seeing; other 48: system conversations are refused. ' +
        'Everyone in the conversation sees the post. The result says whether the post was found again by ' +
        'reading the conversation back (verified); status "unknown" means Teams did not confirm it — check ' +
        'the conversation before retrying. ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows the conversation and exactly what would be posted, says in plain words what ' +
        'is quoted (quoting: the sender, time and text, or "none"), and warns when reply_to is set. ' +
        CONFIRM_INJECTION_RULE +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: false, destructive: true, idempotent: false, openWorld: true }),
      _meta: workMeta('send', { aliases: ['message', 'send a message', 'reply', 'tell', 'post', 'chat'] }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe(
            `The conversation id (or "${SELF_CHAT}" — also by the name "notes" or "self" — for the self-chat), ` +
              "or a case-insensitive substring of the chat title, channel name, team name or an untitled chat's participant names.",
          ),
        text: z
          .string()
          .min(1)
          .describe('The message: plain text (newlines kept, :key: shortcodes become emoji), or Markdown with format "markdown".'),
        format: formatParam,
        reply_to: z
          .string()
          .min(1)
          .optional()
          .describe(
            'A message id. Channel: the root post to reply under (threaded, no quote). Chat: the message to ' +
              'quote-reply — the same as quote. Only when the user asked for a reply or quote.',
          ),
        quote: z.string().min(1).optional().describe('A message id to quote-reply, in any conversation. Only when the user asked for a quote.'),
        mentions: z.array(mentionArg).max(20).optional().describe('People to @-mention in text.'),
        confirmToken: confirmTokenParam,
      },
    },
    async ({ conversation, text, format = 'text', reply_to, quote, mentions, confirmToken }, ctx) => {
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
          format,
        );
        const content = (q ? quoteHtml(q) : '') + rendered.html;
        const quoting = quotingLine(q);
        // reply_to is easy to pass by reflex; it changes what is posted, so the preview says so.
        const replyWarning =
          reply_to === undefined
            ? undefined
            : inChannel
              ? `reply_to is set: this posts into the thread under root post ${reply_to}, not as a new post ` +
                'in the channel. Drop reply_to unless the user asked to reply in that thread.'
              : `reply_to is set: in a chat it quote-replies message ${reply_to}, so the post opens with a ` +
                'quote of it (see quoting). Drop reply_to unless the user asked for a quote.';
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
          summary:
            `Post a message to ${target.kind} ${JSON.stringify(target.title)}` +
            (q ? `, quoting ${JSON.stringify(q.senderName)}` : ', no quote') +
            (inChannel && reply_to !== undefined ? `, in the thread under ${reply_to}` : ''),
          // The signed-in user, so an approval minted under one account never posts as another.
          account: me,
          target: postTo,
          request: { method: 'POST', path, body: bound },
          // The conversation title, the quoted message and the resolved names
          // are written by other people, so each is fenced as a read tool's
          // result is; `text` is the caller's own.
          preview: {
            conversation: untrustedField(target),
            quoting,
            ...(replyWarning ? { warning: replyWarning } : {}),
            text,
            format,
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
              ? { quoted_message: untrustedField({ id: q.id, sender: q.senderName, preview: clipText(q.text, PREVIEW_CHARS) }) }
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
