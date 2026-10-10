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
import type { ChatServiceClient, RawConversation, RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import {
  isUnread,
  isUserMessage,
  matchConversations,
  normaliseMessage,
  oldestFirst,
  readHorizon,
  SELF_CHAT_ALIASES,
  summariseAll,
  type ConversationSummary,
  type NormalisedMessage,
} from '../model.js';
import { confirmedSummary } from './confirm.js';
import { enrichTitles } from './titles.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';
import { DEFAULT_LIMIT, limitArg, makeLimitedStopAt, sinceArg } from './limit.js';

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

/**
 * Settle a listing's unread hint with what the read actually found: an unread
 * message confirms it, and a read that saw everything past the horizon without
 * one refutes it. A window-cut read with none proves nothing, so the hint stays.
 */
export function confirmUnread(summary: ConversationSummary, unreadCount: number, exact: boolean): ConversationSummary {
  if (unreadCount > 0) return confirmedSummary(summary, true);
  return exact ? confirmedSummary(summary, false) : summary;
}

function describeCandidate(s: ConversationSummary): string {
  const where = s.team !== undefined ? ` in ${s.team}` : '';
  return `"${s.title}"${where} (${s.kind}, id ${s.id})`;
}

/**
 * The one conversation a `conversation` argument names: an exact id, else a
 * unique title/team substring. None, or several, is an `McpToolError` — the
 * latter naming up to 10 candidates so the caller can pass an id. `verb` fills
 * the hint ("Pass the id of the one to read").
 */
export function resolveConversation(
  summaries: readonly ConversationSummary[],
  conversation: string,
  verb: string,
): ConversationSummary {
  const matches = matchConversations(summaries, conversation);
  if (matches.length === 0) {
    throw new McpToolError(`No chat or channel matches "${conversation}".`, {
      hint: 'List them with teams_list_chats or teams_list_teams_and_channels and pass an id.',
    });
  }
  if (matches.length > 1) {
    const named = matches.slice(0, MAX_CANDIDATES).map(describeCandidate).join('; ');
    const more = matches.length > MAX_CANDIDATES ? ` (first ${MAX_CANDIDATES} shown)` : '';
    throw new McpToolError(`"${conversation}" matches ${matches.length} conversations${more}: ${named}.`, {
      hint: `Pass the id of the one to ${verb} as conversation.`,
    });
  }
  return matches[0];
}

/**
 * {@link resolveConversation} against the whole listing, with untitled group
 * chats titled by their participants and 1:1s by the other person
 * (`titles.ts`) so a name can pick one of several chats that share a last
 * sender, or a 1:1 the user wrote in last. An exact id titles just that one,
 * and a self-chat alias (`notes`, `self`) titles nothing.
 */
export async function resolveListed(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  conversation: string,
  verb: string,
): Promise<ConversationSummary> {
  const summaries = summariseAll(raw, me);
  if (SELF_CHAT_ALIASES.has(conversation.trim().toLowerCase())) return resolveConversation(summaries, conversation, verb);
  const exact = summaries.find((s) => s.id === conversation);
  return resolveConversation(await enrichTitles(chatsvc, raw, me, exact ? [exact] : summaries), conversation, verb);
}

export function registerConversationTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_read_conversation',
    {
      title: 'Read a Teams chat or channel',
      description:
        'Read the messages of any chat or channel the signed-in user is in, by id (from ' +
        'teams_list_chats / teams_list_teams_and_channels) or by name, without opening it in ' +
        'Teams. "48:notes" (or the name "notes" / "self") is the user\'s own notes self-chat. Messages come oldest first, each with id, time (ISO 8601), sender, fromMe, ' +
        'text, subject, edited/deleted flags and attachments; a channel is returned as threads, ' +
        'each root post with its replies (a root older than the window is marked missing). ' +
        'unread_only returns just the messages the user has not read; since stops at a time; ' +
        'limit keeps the newest (truncated says when more exist; unread_count_scope "window" says ' +
        'unread_count only covers what limit/since let through). This does NOT mark anything ' +
        'read — the user\'s Teams unread badges are unaffected. Attachments are listed per message (images by ref, files by ' +
        'name) and fetched with teams_get_attachment.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe(
            'The conversation id ("48:notes", "notes" or "self" for the self-chat), or a case-insensitive ' +
              "substring of the chat title, channel name, team name or an untitled chat's participant names.",
          ),
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
        const summary = await resolveListed(chatsvc, raw, me, conversation, 'read');
        // The self-chat resolves even when the listing does not carry it.
        const conv = raw.find((c) => c.id === summary.id) ?? { id: summary.id };
        const horizon = readHorizon(conv);
        const cutoff = since !== undefined ? Date.parse(since) : undefined;
        const wanted = (m: RawMessage): boolean => isUserMessage(m) && (!unread_only || isUnread(m, horizon, me));

        const limited = makeLimitedStopAt(unread_only ? horizon : undefined, wanted, limit ?? DEFAULT_LIMIT);
        // A `since` cutoff above the read horizon leaves unread messages unseen.
        let sinceCut = false;
        let capped = false;
        const stopAt = (m: RawMessage): boolean => {
          // The read horizon (checked inside `limited`) wins over `since`, as it is the nearer boundary.
          const atHorizon = unread_only && Number(m.id) <= horizon;
          const at = arrivalMs(m);
          if (!atHorizon && cutoff !== undefined && Number.isFinite(at) && at < cutoff) {
            if (Number(m.id) > horizon) sinceCut = true;
            return true;
          }
          return limited.stopAt(m);
        };
        const fetched = await chatsvc.getMessages(summary.id, {
          stopAt,
          max: MAX_RAW_MESSAGES,
          onCapped: () => (capped = true),
        });

        const messages = oldestFirst(fetched.filter(wanted)).map((m) => normaliseMessage(m, me));
        const unread_count = fetched.filter((m) => wanted(m) && isUnread(m, horizon, me)).length;
        const body = summary.kind === 'channel' ? { threads: groupThreads(messages) } : { messages };
        const truncated = limited.truncated || capped;
        // Whether every message past the read horizon was seen, so unread_count is the whole count.
        const exact = !truncated && !sinceCut;

        return untrustedResult({
          conversation: confirmUnread(summary, unread_count, exact),
          unread_count,
          ...(exact ? {} : { unread_count_scope: 'window' as const }),
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
