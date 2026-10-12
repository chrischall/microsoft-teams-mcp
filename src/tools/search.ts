/**
 * `teams_search_messages` — find messages across the user's chats and
 * channels by text, from the Teams chat-service API.
 *
 * The chat service has no search endpoint, so this is a bounded CLIENT-SIDE
 * SCAN: one listing picks the conversations with activity inside the window
 * (most recent first, `max_conversations`, default {@link DEFAULT_SCANNED_CONVERSATIONS}); each is read
 * newest first down to the `since` cutoff (default the last
 * {@link DEFAULT_SINCE_DAYS} days), a few at a time, and every real message
 * whose plain text matches is a hit. What the bounds left out is reported in
 * `scan`, never dropped silently.
 *
 * Verified live 2026-10-12 against a real tenant: a 30-day `deploy` search
 * read 100 conversations (~5,400 messages) in ~15s and found 115 hits; `from`,
 * `since` and a `conversation` narrowed to the self-chat behaved as tested.
 * That listing had ~500 conversations with SOME activity in 30 days — mostly
 * system events in channels and meeting chats — hence the default cap and the
 * `scan` report.
 *
 * Why not Teams' own search: the web client searches through Microsoft
 * Substrate — seen live 2026-10-12, the search box issues
 * `POST https://substrate.office.com/searchservice/api/v2/query` (200). Its
 * bearer is a different token from the chat service's (Substrate audience,
 * presumably — not yet observed). Capturing it needs `substrate.office.com`
 * added to the bridge pairing's trust scope: a re-pair the user approves in the
 * ContextMint Bridge popup (the probe's capture stopped at "pairing not yet
 * approved"). The query is also issued from a worker, so its request and
 * response bodies could not be observed from the page either. Until both are
 * verified live the scan is the path: slower and window-bound, but it needs
 * nothing beyond the chat-service token every other tool already uses, and
 * nothing it reads marks anything read.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { mapWithConcurrency, toolAnnotations } from '@chrischall/mcp-utils';
import type { RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import {
  conversationSummary,
  isUserMessage,
  newestActivityMs,
  normaliseMessage,
  type ConversationSummary,
  type NormalisedMessage,
} from '../model.js';
import { arrivalMs, resolveListed } from './conversation.js';
import { CONFIRM_CONCURRENCY, isUnreachable } from './confirm.js';
import { enrichTitles } from './titles.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

const DEFAULT_LIMIT = 25;
const MAX_LIMIT = 200;

/** The window searched when `since` is not given. */
export const DEFAULT_SINCE_DAYS = 30;
const DAY_MS = 86_400_000;

/**
 * Conversations read per call, most recently active first, unless
 * `max_conversations` says otherwise. Live 2026-10-12, 100 reads took ~15s;
 * the listing had ~500 conversations with some activity in 30 days, most of it
 * system events (call recordings, member changes) in channels and meeting chats.
 */
export const DEFAULT_SCANNED_CONVERSATIONS = 100;
export const MAX_SCANNED_CONVERSATIONS = 500;

/**
 * Raw messages read per conversation before giving up on it (5 pages of 200).
 * The real bound is the `since` cutoff; this caps one very busy conversation.
 */
const MAX_RAW_PER_CONVERSATION = 1000;

/**
 * The query's terms, lowercased: whitespace-separated words, a `"quoted
 * phrase"` kept whole.
 */
function queryTerms(query: string): string[] {
  const terms: string[] = [];
  for (const [, phrase, word] of query.toLowerCase().matchAll(/"([^"]*)"|(\S+)/g)) {
    const t = (phrase ?? word ?? '').replace(/"/g, '').trim();
    if (t !== '') terms.push(t);
  }
  return terms;
}

/**
 * Whether `text` matches `query`: every term (word or `"quoted phrase"`)
 * appears in it, in any order, ignoring case. A query with no terms matches
 * nothing.
 */
export function matchesQuery(text: string, query: string): boolean {
  const terms = queryTerms(query);
  if (terms.length === 0) return false;
  const haystack = text.toLowerCase();
  return terms.every((t) => haystack.includes(t));
}

/** What a hit carries of its message: no attachments or reactions, which a search does not need. */
type HitMessage = Omit<NormalisedMessage, 'attachments' | 'reactions' | 'deleted'>;

interface Hit {
  conversation: Pick<ConversationSummary, 'id' | 'title' | 'kind' | 'team'>;
  message: HitMessage;
}

function hitMessage(m: NormalisedMessage): HitMessage {
  const { attachments: _a, reactions: _r, deleted: _d, ...rest } = m;
  return rest;
}

function hitConversation(s: ConversationSummary): Hit['conversation'] {
  return { id: s.id, title: s.title, kind: s.kind, ...(s.team !== undefined ? { team: s.team } : {}) };
}

export function registerSearchTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_search_messages',
    {
      title: 'Search Teams messages',
      description:
        'Find messages by text across the chats and channels the signed-in user is in. Every word of ' +
        'query must appear in the message text (any order, ignoring case; "quoted words" match as a ' +
        'phrase). Hits come newest first, each with its conversation (id, title, kind, team) and the ' +
        'message (id, time ISO 8601, sender, fromMe, text, subject, replyTo for a channel reply); ' +
        'read a hit in context with teams_read_conversation. conversation narrows to one chat or ' +
        'channel (id or name); from keeps messages whose sender display name contains it; since ' +
        `(ISO 8601) sets how far back to look (default the last ${DEFAULT_SINCE_DAYS} days); limit ` +
        `caps the hits (default ${DEFAULT_LIMIT}, max ${MAX_LIMIT}; total_hits and truncated say ` +
        'when there were more). Teams has no search endpoint this server can use, so this scans: ' +
        `the ${DEFAULT_SCANNED_CONVERSATIONS} (max_conversations, up to ${MAX_SCANNED_CONVERSATIONS}) most ` +
        'recently active conversations with activity since the cutoff are read back to it (channels hidden in Teams are skipped unless named), and ' +
        'scan reports what was read and what the bounds left out — a busy window can take a while, ' +
        'so narrow with conversation or since when you can. Deleted messages and system events ' +
        'never match. This does NOT mark anything read.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        query: z
          .string()
          .refine((s) => queryTerms(s).length > 0, { message: 'query must contain a word to search for' })
          .describe(
            'Words to find in the message text, ignoring case; every word must appear. Wrap words in ' +
              'double quotes to match them as one phrase.',
          ),
        conversation: z
          .string()
          .min(1)
          .optional()
          .describe(
            'Search only this chat or channel: its id ("48:notes", "notes" or "self" for the self-chat), ' +
              "or a case-insensitive substring of its title, channel name, team name or an untitled chat's " +
              'participant names (it must name exactly one).',
          ),
        from: z
          .string()
          .min(1)
          .optional()
          .describe("Only messages whose sender's display name contains this, ignoring case."),
        since: z
          .string()
          .refine((s) => !Number.isNaN(Date.parse(s)), { message: 'since must be an ISO 8601 date-time' })
          .optional()
          .describe(
            `Only messages at or after this ISO 8601 date-time (default ${DEFAULT_SINCE_DAYS} days ago). ` +
              'An older since searches further back and takes longer.',
          ),
        limit: z
          .number()
          .int()
          .min(1)
          .max(MAX_LIMIT)
          .optional()
          .describe(`Maximum hits to return, newest kept (default ${DEFAULT_LIMIT}, max ${MAX_LIMIT}).`),
        max_conversations: z
          .number()
          .int()
          .min(1)
          .max(MAX_SCANNED_CONVERSATIONS)
          .optional()
          .describe(
            `Most conversations to read, most recently active first (default ${DEFAULT_SCANNED_CONVERSATIONS}, ` +
              `max ${MAX_SCANNED_CONVERSATIONS}). More finds hits in quieter conversations but takes longer ` +
              '(roughly 15s per 100).',
          ),
      },
    },
    async ({ query, conversation, from, since, limit, max_conversations }) => {
      try {
        const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
        const cutoff = since !== undefined ? Date.parse(since) : Date.now() - DEFAULT_SINCE_DAYS * DAY_MS;

        // The conversations to read: the one named, else every one with activity
        // inside the window, most recent first, hidden channels left out.
        let targets: ConversationSummary[];
        let outsideWindow = 0;
        if (conversation !== undefined) {
          targets = [await resolveListed(chatsvc, raw, me, conversation, 'search')];
        } else {
          const inWindow: { summary: ConversationSummary; newest: number }[] = [];
          // Only what the listing carries: the self-chat is searched when Teams lists it, or when named.
          for (const c of raw) {
            const summary = conversationSummary(c, me, raw);
            if (summary.hidden === true) continue;
            const newest = newestActivityMs(c);
            // A conversation the listing gives no time for is read anyway, last.
            if (newest !== undefined && newest < cutoff) outsideWindow++;
            else inWindow.push({ summary, newest: newest ?? 0 });
          }
          inWindow.sort((a, b) => b.newest - a.newest);
          targets = inWindow.map((c) => c.summary);
        }
        const maxScanned = max_conversations ?? DEFAULT_SCANNED_CONVERSATIONS;
        const scanned = targets.slice(0, maxScanned);
        const notScanned = targets.length - scanned.length;

        const fromLc = from?.trim().toLowerCase();
        type Read =
          | { kind: 'read'; hits: RawMessage[]; seen: number; capped: boolean }
          | { kind: 'unreachable' }
          | { kind: 'failed'; error: unknown };
        const reads = await mapWithConcurrency(scanned, CONFIRM_CONCURRENCY, async (summary): Promise<Read> => {
          let capped = false;
          try {
            const fetched = await chatsvc.getMessages(summary.id, {
              stopAt: (m) => {
                const at = arrivalMs(m);
                return Number.isFinite(at) && at < cutoff;
              },
              max: MAX_RAW_PER_CONVERSATION,
              onCapped: () => (capped = true),
            });
            const hits = fetched.filter((m) => {
              if (!isUserMessage(m)) return false;
              const n = normaliseMessage(m, me);
              if (n.deleted) return false;
              if (fromLc !== undefined && !n.sender.toLowerCase().includes(fromLc)) return false;
              return matchesQuery(n.subject !== undefined ? `${n.subject}\n${n.text}` : n.text, query);
            });
            return { kind: 'read', hits, seen: fetched.length, capped };
          } catch (error) {
            return isUnreachable(error) ? { kind: 'unreachable' } : { kind: 'failed', error };
          }
        });

        const found: { summary: ConversationSummary; message: NormalisedMessage; at: number }[] = [];
        const failed: { summary: ConversationSummary; error: unknown }[] = [];
        const partial: ConversationSummary[] = [];
        let unreachable = 0;
        let messagesScanned = 0;
        let anyRead = false;
        reads.forEach((r, i) => {
          const summary = scanned[i];
          if (r.kind === 'unreachable') unreachable++;
          else if (r.kind === 'failed') failed.push({ summary, error: r.error });
          else {
            anyRead = true;
            messagesScanned += r.seen;
            if (r.capped) partial.push(summary);
            for (const m of r.hits) found.push({ summary, message: normaliseMessage(m, me), at: arrivalMs(m) });
          }
        });
        // Every read failing is a failure of the call (a revoked token, a rate limit), not of one conversation.
        if (failed.length > 0 && !anyRead) throw failed[0].error;

        found.sort((a, b) => (Number.isFinite(b.at) ? b.at : 0) - (Number.isFinite(a.at) ? a.at : 0));
        const cap = limit ?? DEFAULT_LIMIT;
        const kept = found.slice(0, cap);

        // Only conversations that are returned are titled (an untitled chat's title costs a read).
        const named = [...new Map([...kept.map((h) => h.summary), ...partial, ...failed.map((f) => f.summary)].map((s) => [s.id, s])).values()];
        const titled = new Map((await enrichTitles(chatsvc, raw, me, named)).map((s) => [s.id, s]));
        const final = (s: ConversationSummary): ConversationSummary => titled.get(s.id) ?? s;

        const hits: Hit[] = kept.map((h) => ({ conversation: hitConversation(final(h.summary)), message: hitMessage(h.message) }));
        const sinceIso = new Date(cutoff).toISOString();
        const leftOut = notScanned > 0 || partial.length > 0;

        return untrustedResult({
          query,
          total_hits: found.length,
          hits,
          ...(found.length > kept.length
            ? {
                truncated: {
                  total: found.length,
                  returned: kept.length,
                  hint: `Only the newest ${kept.length} of ${found.length} hits were returned; pass a larger limit (max ${MAX_LIMIT}) or narrow with conversation, from or since.`,
                },
              }
            : {}),
          scan: {
            method: 'client_scan',
            since: sinceIso,
            conversations_scanned: scanned.length,
            conversations_outside_window: outsideWindow,
            conversations_not_scanned: notScanned,
            messages_scanned: messagesScanned,
            skipped_unreachable: unreachable,
            ...(partial.length > 0
              ? { partially_scanned: partial.map((s) => ({ id: s.id, title: final(s).title })) }
              : {}),
            ...(leftOut
              ? {
                  hint:
                    (notScanned > 0
                      ? `${notScanned} less recently active conversation${notScanned === 1 ? ' was' : 's were'} not scanned (max_conversations ${maxScanned}; up to ${MAX_SCANNED_CONVERSATIONS}). `
                      : '') +
                    (partial.length > 0
                      ? `${partial.length} conversation${partial.length === 1 ? ' was' : 's were'} only read back ${MAX_RAW_PER_CONVERSATION} messages, short of ${sinceIso}. `
                      : '') +
                    'Raise max_conversations, or narrow with conversation or a later since, to search everything in range.',
                }
              : {}),
          },
          ...(failed.length > 0
            ? {
                errors: failed.map((f) => ({
                  conversation: hitConversation(final(f.summary)),
                  error: wrapApiError(f.error, 'search this conversation').message,
                })),
              }
            : {}),
        });
      } catch (err) {
        throw wrapApiError(err, 'search messages');
      }
    },
  );
}
