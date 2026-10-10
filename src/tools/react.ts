/**
 * `teams_react` — add or withdraw the signed-in user's reaction on a message
 * in any chat or channel, through the chat-service API.
 *
 * Gated by `confirmWrite` exactly as `teams_send_message` is: a reaction is
 * visible to everyone in the conversation. It replaces nothing and setting the
 * same reaction twice is still one reaction, so it is declared non-destructive
 * and idempotent.
 *
 * Only the message's `emotions` property is written. Teams also records the
 * user's `recentReactions` preference when reacting in its UI; that is a UI
 * nicety and this tool never touches the user's preferences.
 *
 * After the write the message is read back and the reaction looked for (or,
 * on removal, looked for and not found) among the user's own.
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
} from '@chrischall/mcp-utils';
import type { ChatServiceClient } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import { htmlToText, normaliseMessage, senderMri } from '../model.js';
import { wrapApiError } from './errors.js';
import { resolveTarget, SELF_CHAT } from './send.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

const TOOL = 'teams_react';

/**
 * The ✅ reaction's key. UNVERIFIED: inferred from the unicode-style key rule
 * (see `KEY_RULE`), not yet read off a real ✅ reaction — confirm it live and
 * fix it here if Teams uses another.
 */
export const CHECKMARK_KEY = '2705_whiteheavycheckmark';

/** Friendly names and emoji → reaction key. Emoji are matched without U+FE0F. */
const ALIASES: Readonly<Record<string, string>> = {
  thumbsup: 'like',
  '👍': 'like',
  like: 'like',
  heart: 'heart',
  '❤': 'heart',
  laugh: 'laugh',
  '😆': 'laugh',
  surprised: 'surprised',
  '😮': 'surprised',
  sad: 'sad',
  '😢': 'sad',
  angry: 'angry',
  '😡': 'angry',
  checkmark: CHECKMARK_KEY,
  '✅': CHECKMARK_KEY,
};

/** What a raw key looks like: a legacy name (`fire`) or a unicode-style key (`1f62c_grimacingface`). */
const RAW_KEY = /^[a-z0-9_]+$/;

const KEY_RULE =
  'Unicode-style reaction keys are "<lowercase hex code point>_<the Unicode character name, lowercased, ' +
  'spaces and punctuation removed>" — e.g. 💯 1f4af_hundredpointssymbol, 👀 1f440_eyes, 😬 1f62c_grimacingface. ' +
  'Teams also keeps short names for common ones: like, heart, laugh, surprised, sad, angry, yes, no, party, ' +
  'fire, smile, cool, wink, hearteyes, manshrug, fistbump, handsinair.';

/** Previewed text of the message reacted to. */
const PREVIEW_CHARS = 200;

/**
 * The reaction key for what the caller passed: an alias (`thumbsup`, 👍, …),
 * else a raw key as-is (case, spaces and surrounding `:colons:` ignored). An
 * emoji with no alias is refused with the key rule — no Unicode-name database
 * is bundled, so its key cannot be derived here.
 */
export function reactionKey(input: string): string {
  const s = input.trim().replace(/️/g, '').replace(/^:(.+):$/, '$1').toLowerCase();
  const alias = ALIASES[s];
  if (alias !== undefined) return alias;
  if (RAW_KEY.test(s)) return s;
  if (/\p{Extended_Pictographic}/u.test(s)) {
    const points = [...s].map((c) => c.codePointAt(0)!.toString(16));
    throw new McpToolError(`There is no known reaction key for ${input.trim()}; pass its key instead.`, {
      hint:
        (points.length === 1
          ? `Its code point is U+${points[0].toUpperCase()}, so its key is "${points[0]}_<its Unicode name>". `
          : `Its code points are ${points.map((p) => `U+${p.toUpperCase()}`).join(' ')}. `) + KEY_RULE,
    });
  }
  throw new McpToolError(`"${input}" is not a reaction: pass an emoji, a name such as thumbsup, or a reaction key.`, {
    hint: KEY_RULE,
  });
}

/** Whether the user's own reaction `key` is on the message, read fresh. */
async function hasMyReaction(
  chatsvc: ChatServiceClient,
  conversationId: string,
  messageId: string,
  key: string,
  me: string,
): Promise<boolean> {
  const m = await chatsvc.getMessage(conversationId, messageId);
  if (m === undefined) return false;
  return normaliseMessage(m, me).reactions?.some((r) => r.key === key && r.byMe) ?? false;
}

export function registerReactTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    TOOL,
    {
      title: 'React to a Teams message',
      description:
        'Add (or, with remove: true, withdraw) the signed-in user\'s reaction on a message in any chat or ' +
        'channel. conversation is an id or name as for teams_read_conversation; message_id comes from ' +
        'teams_read_conversation. reaction is an emoji or name — thumbsup/👍/like, heart/❤️, laugh/😆, ' +
        `surprised/😮, sad/😢, angry/😡, checkmark/✅ — or any raw Teams reaction key (e.g. fire, ` +
        '1f389_partypopper); an emoji without a known key is refused with the rule for spelling its key. ' +
        `conversation "${SELF_CHAT}" is the user's own notes self-chat — use it to try this tool. Everyone ` +
        'in the conversation sees the reaction. The result says whether reading the message back showed the ' +
        'change (verified). ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows the conversation, the message and the reaction. ' +
        CONFIRM_INJECTION_RULE +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: false, destructive: false, idempotent: true, openWorld: true }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe(`The conversation id (or "${SELF_CHAT}"), or a case-insensitive substring of the chat title, channel name or team name.`),
        message_id: z.string().min(1).describe('The message to react to, as teams_read_conversation lists it.'),
        reaction: z.string().min(1).describe('An emoji (👍), a name (thumbsup) or a Teams reaction key (like, 1f389_partypopper).'),
        remove: z.boolean().optional().describe('Withdraw this reaction instead of adding it (default false).'),
        confirmToken: confirmTokenParam,
      },
    },
    async ({ conversation, message_id, reaction, remove, confirmToken }, ctx) => {
      try {
        const key = reactionKey(reaction);
        const me = await chatsvc.meMri();
        const target = await resolveTarget(chatsvc, conversation, me, 'react in');

        const m = await chatsvc.getMessage(target.id, message_id);
        if (m === undefined) {
          throw new McpToolError(`There is no message ${message_id} in this conversation to react to.`, {
            hint: 'Take message ids from teams_read_conversation.',
          });
        }
        const already = normaliseMessage(m, me).reactions?.some((r) => r.key === key && r.byMe) ?? false;

        const method = remove ? 'DELETE' : 'PUT';
        const path =
          `/users/ME/conversations/${encodeURIComponent(target.id)}` +
          `/messages/${encodeURIComponent(message_id)}/properties?name=emotions`;
        const gate = await confirmWrite(ctx, {
          tool: TOOL,
          action: remove ? 'chat.remove_reaction' : 'chat.react',
          message: `Review and confirm this Teams reaction before it is ${remove ? 'removed' : 'added'}:`,
          summary: `${remove ? 'Remove' : 'Add'} reaction "${key}" on a message in ${target.kind} "${target.title}"`,
          // The signed-in user, so an approval minted under one account never reacts as another.
          account: me,
          target: `${target.id}/${message_id}`,
          // The reaction time is minted when the reaction is set, so it is not bound.
          request: { method, path, body: { emotions: { key } } },
          preview: {
            conversation: target,
            message: {
              id: m.id,
              sender: m.imdisplayname?.trim() || senderMri(m.from) || 'unknown',
              preview: htmlToText(m.content ?? '').slice(0, PREVIEW_CHARS),
            },
            reaction: key,
            already_reacted: already,
          },
          confirmToken,
        });
        if (gate) return gate;

        if (remove) await chatsvc.removeReaction(target.id, message_id, key);
        else await chatsvc.setReaction(target.id, message_id, key);

        let verified = false;
        let readBackFailed = false;
        try {
          const present = await hasMyReaction(chatsvc, target.id, message_id, key, me);
          verified = remove ? !present : present;
        } catch {
          readBackFailed = true;
        }
        return untrustedResult({
          status: remove ? 'removed' : 'reacted',
          conversation: target,
          message_id,
          reaction: key,
          verified,
          ...(verified
            ? {}
            : {
                warning: readBackFailed
                  ? 'Teams accepted the change, but reading the message back failed — check the message (teams_read_conversation).'
                  : `Teams accepted the change, but reading the message back did not show the reaction ${remove ? 'gone' : 'present'} — check the message (teams_read_conversation); the key may not be one Teams knows.`,
              }),
        });
      } catch (err) {
        throw wrapApiError(err, 'react to the message');
      }
    },
  );
}
