/**
 * `teams_read_conversation` — read any chat or channel, by id or name, from
 * the Teams chat-service API. Nothing has to be open in the browser, and the
 * read marker is never written, so the user's unread badges are untouched.
 *
 * Paging runs newest first and stops at the first boundary that applies: the
 * read horizon (`unread_only`), the `since` cutoff, or one message past
 * `limit`. The result is turned oldest first; a channel's posts are grouped
 * under their root post.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { McpToolError, toolAnnotations } from '@chrischall/mcp-utils';
import type { RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import {
  conversationSummary,
  isUnread,
  isUserMessage,
  matchConversations,
  normaliseMessage,
  readHorizon,
  type ConversationSummary,
  type NormalisedMessage,
} from '../model.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';
import { DEFAULT_LIMIT, limitArg, sinceArg } from './limit.js';

const MAX_MESSAGES = 500;

/**
 * Raw messages read before giving up (20 pages of 200). The real bound is the
 * `stopAt` boundary; this only caps a conversation of nothing but system events.
 */
const MAX_RAW_MESSAGES = 4000;

/** Candidates an ambiguous-name error lists. */
const MAX_CANDIDATES = 10;

export type ThreadRoot = NormalisedMessage | { id: string; missing: true };
export interface Thread {
  root: ThreadRoot;
  replies: NormalisedMessage[];
}

/** Arrival time in epoch ms: a message's id IS its arrival ms; the ISO fields are the fallback. */
function arrivalMs(m: RawMessage): number {
  const n = Number(m.id);
  if (Number.isFinite(n) && n > 0) return n;
  return Date.parse(m.originalarrivaltime ?? m.composetime ?? '');
}

/**
 * Group oldest-first channel messages under their root post, in order of first
 * appearance. A reply whose root is older than the window still appears, under
 * a `{ id, missing: true }` placeholder root.
 */
export function groupThreads(messages: readonly NormalisedMessage[]): Thread[] {
  const byRoot = new Map<string, Thread>();
  for (const m of messages) {
    const rootId = m.replyTo ?? m.id;
    let thread = byRoot.get(rootId);
    if (thread === undefined) {
      thread = { root: { id: rootId, missing: true }, replies: [] };
      byRoot.set(rootId, thread);
    }
    if (m.replyTo === undefined) thread.root = m;
    else thread.replies.push(m);
  }
  return [...byRoot.values()];
}

function describeCandidate(s: ConversationSummary): string {
  const where = s.team !== undefined ? ` in ${s.team}` : '';
  return `"${s.title}"${where} (${s.kind}, id ${s.id})`;
}

export function registerConversationTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_read_conversation',
    {
      title: 'Read a Teams chat or channel',
      description:
        'Read the messages of any chat or channel the signed-in user is in, by id (from ' +
        'teams_list_chats / teams_list_teams_and_channels) or by name, without opening it in ' +
        'Teams. Messages come oldest first, each with id, time (ISO 8601), sender, fromMe, ' +
        'text, subject, edited/deleted flags and attachments; a channel is returned as threads, ' +
        'each root post with its replies (a root older than the window is marked missing). ' +
        'unread_only returns just the messages the user has not read; since stops at a time; ' +
        'limit keeps the newest. This does NOT mark anything read — the user\'s Teams unread ' +
        'badges are unaffected. Attachments are listed per message (images by ref, files by ' +
        'name) and fetched with teams_get_attachment.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe('The conversation id, or a case-insensitive substring of the chat title, channel name or team name.'),
        unread_only: z
          .boolean()
          .optional()
          .describe('Only return messages the user has not read yet (default false). Never marks them read.'),
        since: sinceArg,
        limit: limitArg(MAX_MESSAGES, 'messages'),
      },
    },
    async ({ conversation, unread_only, since, limit }) => {
      try {
        const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
        const summaries = raw.map((c) => conversationSummary(c, me, raw));
        const matches = matchConversations(summaries, conversation);
        if (matches.length === 0) {
          throw new McpToolError(`No chat or channel matches "${conversation}".`, {
            hint: 'List them with teams_list_chats or teams_list_teams_and_channels and pass an id.',
          });
        }
        if (matches.length > 1) {
          const named = matches.slice(0, MAX_CANDIDATES).map(describeCandidate).join('; ');
          const more = matches.length > MAX_CANDIDATES ? ` (first ${MAX_CANDIDATES} shown)` : '';
          throw new McpToolError(
            `"${conversation}" matches ${matches.length} conversations${more}: ${named}.`,
            { hint: 'Pass the id of the one to read as conversation.' },
          );
        }
        const summary = matches[0];
        const conv = raw.find((c) => c.id === summary.id)!;
        const horizon = readHorizon(conv);
        const cutoff = since !== undefined ? Date.parse(since) : undefined;
        const n = limit ?? DEFAULT_LIMIT;
        const wanted = (m: RawMessage): boolean => isUserMessage(m) && (!unread_only || isUnread(m, horizon, me));

        // `stopAt` sees each message once, newest first, so it can count: the
        // first wanted message past `limit` is the boundary, and its existence
        // is what says the result was truncated.
        let kept = 0;
        let truncated = false;
        const stopAt = (m: RawMessage): boolean => {
          const at = arrivalMs(m);
          if (unread_only && Number(m.id) <= horizon) return true;
          if (cutoff !== undefined && Number.isFinite(at) && at < cutoff) return true;
          if (!wanted(m)) return false;
          if (kept >= n) return (truncated = true);
          kept++;
          return false;
        };
        const fetched = await chatsvc.getMessages(summary.id, { stopAt, max: MAX_RAW_MESSAGES });

        const messages = fetched
          .filter(wanted)
          .reverse()
          .map((m) => normaliseMessage(m, me));
        const unread_count = fetched.filter((m) => wanted(m) && isUnread(m, horizon, me)).length;
        const body = summary.kind === 'channel' ? { threads: groupThreads(messages) } : { messages };

        return untrustedResult({
          conversation: summary,
          unread_count,
          ...body,
          ...(truncated
            ? {
                truncated: {
                  returned: messages.length,
                  hint:
                    `Only the newest ${messages.length} messages were returned; pass a larger limit ` +
                    `(max ${MAX_MESSAGES}) or a since time to read further back.`,
                },
              }
            : {}),
        });
      } catch (err) {
        throw wrapApiError(err, 'read the conversation');
      }
    },
  );
}
