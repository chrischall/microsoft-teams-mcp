/**
 * `teams_search_messages` — find messages across the user's chats and
 * channels by text.
 *
 * Two paths, chosen by `mode` ("auto" by default):
 *
 * 1. SUBSTRATE (primary). Teams' own search box searches through Microsoft
 *    Substrate — `POST https://substrate.office.com/searchservice/api/v2/query`.
 *    `substrate.ts` is the client; its token is a third one, captured off the
 *    Teams tab's own search request (`substrate-auth.ts`) and cached in
 *    `substrate-token.json`. Verified live 2026-10-12 (see `substrate.ts` for
 *    the request and response shapes): `conversation` and `since` are sent
 *    SERVER-SIDE (`ClientThreadId:"…"`, `sent>=…`), and also re-checked here;
 *    `from` is applied client-side, because Substrate's `from:` matches word
 *    prefixes rather than "display name contains". Pages of
 *    {@link SUBSTRATE_PAGE_SIZE} are read until `limit` is filled, Substrate
 *    runs out, a whole page falls before `since`, or
 *    {@link MAX_SUBSTRATE_PAGES} pages were read. Hit text is Substrate's
 *    plain-text preview (`excerpt: true`), and hits carry no `fromMe` (no
 *    sender MRI). "auto" still falls back to the scan on ANY Substrate
 *    failure, reporting `fallback_reason`.
 *
 * 2. SCAN (fallback, or `mode: "scan"`). The chat service has no search
 *    endpoint, so this is a bounded CLIENT-SIDE SCAN: one listing picks the
 *    conversations with activity inside the window (most recent first,
 *    `max_conversations`, default {@link DEFAULT_SCANNED_CONVERSATIONS}); each is
 *    read newest first down to the `since` cutoff (default the last
 *    {@link DEFAULT_SINCE_DAYS} days), a few at a time, and every real message
 *    whose plain text matches is a hit. What the bounds left out is reported in
 *    `scan`, never dropped silently.
 *
 *    Verified live 2026-10-12 against a real tenant: a 30-day `deploy` search
 *    read 100 conversations (~5,400 messages) in ~15s and found 115 hits;
 *    `from`, `since` and a `conversation` narrowed to the self-chat behaved as
 *    tested. That listing had ~500 conversations with SOME activity in 30 days —
 *    mostly system events in channels and meeting chats — hence the default cap
 *    and the `scan` report. It needs nothing beyond the chat-service token every
 *    other tool already uses.
 *
 * Neither path marks anything read.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { mapWithConcurrency, McpToolError, toolAnnotations } from '@chrischall/mcp-utils';
import type { ChatServiceClient, RawConversation, RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import {
  conversationKind,
  conversationSummary,
  isUserMessage,
  newestActivityMs,
  normaliseMessage,
  type ConversationSummary,
  type NormalisedMessage,
} from '../model.js';
import type { SubstrateMessage, SubstrateSearch } from '../substrate.js';
import { arrivalMs, resolveListed } from './conversation.js';
import { CONFIRM_CONCURRENCY, isUnreachable } from './confirm.js';
import { conversationName, enrichTitles, participantsOf, type ConversationName } from './titles.js';
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

/** Substrate results asked for per request. */
export const SUBSTRATE_PAGE_SIZE = 50;

/** Substrate pages one call reads at most (so 500 results before filtering). */
export const MAX_SUBSTRATE_PAGES = 10;

const MODES = ['auto', 'substrate', 'scan'] as const;
type Mode = (typeof MODES)[number];

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

/**
 * A Substrate hit's message: the scan's fields, less what Substrate may not
 * say (`id`, `time`) or never says (`edited`, `fromMe` — a hit carries no
 * MRI), plus `excerpt: true`, since its text is always Substrate's plain-text
 * preview (at most ~255 characters), never the whole body.
 */
type SubstrateHitMessage = Omit<HitMessage, 'id' | 'time' | 'edited' | 'time_local' | 'fromMe'> & {
  id?: string;
  time?: string;
  excerpt: true;
};

interface Hit {
  conversation: ConversationName;
  message: HitMessage | SubstrateHitMessage;
}

function hitMessage(m: NormalisedMessage): HitMessage {
  const { attachments: _a, reactions: _r, deleted: _d, ...rest } = m;
  return rest;
}

/** A hit's conversation: its id and its name — title, kind, team, and an untitled chat's participants. */
function hitConversation(s: ConversationSummary): Hit['conversation'] {
  return conversationName(s);
}

/** What both paths are asked for. */
interface SearchArgs {
  query: string;
  /** The one conversation named, resolved against the listing. */
  target?: ConversationSummary;
  fromLc?: string;
  cutoff: number;
  cap: number;
}

export function registerSearchTools(
  server: McpServer,
  { chatsvc, substrate }: Pick<TeamsDeps, 'chatsvc'> & Partial<Pick<TeamsDeps, 'substrate'>>,
): void {
  server.registerTool(
    'teams_search_messages',
    {
      title: 'Search Teams messages',
      description:
        'Find messages by text across the chats and channels the signed-in user is in. Hits come ' +
        'newest first, each with its conversation (id, title, kind, team) and the message (id, time ' +
        'ISO 8601, sender, fromMe (scan only), text, subject, replyTo for a channel reply); read a hit in context ' +
        'with teams_list_messages. conversation narrows to one chat or channel (id or name); from ' +
        'keeps messages whose sender display name contains it; since (ISO 8601) sets how far back to ' +
        `look (default the last ${DEFAULT_SINCE_DAYS} days); limit caps the hits (default ` +
        `${DEFAULT_LIMIT}, max ${MAX_LIMIT}; total_hits and truncated say when there were more). ` +
        'source says which path answered. mode "auto" (default) uses Teams\' own search (Microsoft ' +
        'Substrate: matched the way the Teams search box matches — words ANDed with stemming, "quoted ' +
        'phrases", OR — and fast; conversation and since apply to it too; its text is a plain-text preview ' +
        'of up to ~255 characters, marked excerpt: true, and its hits carry no fromMe) and falls back to a ' +
        'scan, with fallback_reason, when Substrate is unavailable or fails — its token is captured from a Teams tab the first time, so a first ' +
        'search may need you to run any search in the Teams search box. mode "substrate" uses only ' +
        'Substrate (an error rather than a fallback); mode "scan" only scans. The SCAN requires every ' +
        'word of query to appear in the message text (any order, ignoring case; "quoted words" match ' +
        `as a phrase) and reads the ${DEFAULT_SCANNED_CONVERSATIONS} (max_conversations, up to ` +
        `${MAX_SCANNED_CONVERSATIONS}) most recently active conversations with activity since the ` +
        'cutoff back to it (channels hidden in Teams are skipped unless named); scan reports what was ' +
        'read and what the bounds left out — a busy window can take a while, so narrow with ' +
        'conversation or since when you can. Deleted messages and system events never match. This ' +
        'does NOT mark anything read.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        query: z
          .string()
          .refine((s) => queryTerms(s).length > 0, { message: 'query must contain a word to search for' })
          .describe(
            'Words to find in the message text, ignoring case. Substrate (the default) matches the way the ' +
              'Teams search box does (words ANDed with stemming, "quoted phrases", OR). The scan requires ' +
              'every word to appear verbatim, any order; words in double quotes match as one phrase.',
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
              'An older since makes a scan search further back and take longer.',
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
            `Scan only: most conversations to read, most recently active first (default ${DEFAULT_SCANNED_CONVERSATIONS}, ` +
              `max ${MAX_SCANNED_CONVERSATIONS}). More finds hits in quieter conversations but takes longer ` +
              '(roughly 15s per 100).',
          ),
        mode: z
          .enum(MODES)
          .optional()
          .describe(
            '"auto" (default): Teams\' own search (Substrate), falling back to a scan when it is ' +
              'unavailable or fails. "substrate": Substrate only, failing instead of falling back. ' +
              '"scan": the client-side scan only.',
          ),
      },
    },
    async ({ query, conversation, from, since, limit, max_conversations, mode }) => {
      try {
        const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
        const cutoff = since !== undefined ? Date.parse(since) : Date.now() - DEFAULT_SINCE_DAYS * DAY_MS;
        const target = conversation !== undefined ? await resolveListed(chatsvc, raw, me, conversation, 'search') : undefined;
        const args: SearchArgs = {
          query,
          ...(target !== undefined ? { target } : {}),
          ...(from !== undefined ? { fromLc: from.trim().toLowerCase() } : {}),
          cutoff,
          cap: limit ?? DEFAULT_LIMIT,
        };

        const m: Mode = mode ?? 'auto';
        let fallbackReason: string | undefined;
        if (m !== 'scan') {
          const unavailable =
            substrate === undefined
              ? 'Substrate search is not available in this server.'
              : !substrate.isConfigured()
                ? 'Substrate search is disabled: TEAMS_DISABLE_FETCHPROXY is set, and the browser bridge is its only token source.'
                : undefined;
          if (unavailable !== undefined || substrate === undefined) {
            const reason = unavailable ?? 'Substrate search is not available in this server.';
            if (m === 'substrate') {
              throw new McpToolError(reason, { kind: 'no_credential', hint: 'Use mode "auto" or "scan" to search by scanning instead.' });
            }
            fallbackReason = reason;
          } else {
            try {
              return untrustedResult({
                query,
                source: 'substrate',
                ...(await substrateSearch(substrate, chatsvc, raw, me, args, m === 'substrate')),
              });
            } catch (err) {
              if (m === 'substrate') throw err;
              fallbackReason = wrapApiError(err, 'search with Substrate').message;
            }
          }
        }

        return untrustedResult({
          query,
          source: 'scan',
          ...(fallbackReason !== undefined ? { fallback_reason: fallbackReason } : {}),
          ...(await scanSearch(chatsvc, raw, me, args, max_conversations)),
        });
      } catch (err) {
        throw wrapApiError(err, 'search messages');
      }
    },
  );
}

/** A conversation's summary from the listing, or a bare one named by its id when the listing lacks it. */
function summaryFor(id: string, listed: Map<string, RawConversation>, raw: readonly RawConversation[], me: string): ConversationSummary {
  const c = listed.get(id);
  if (c !== undefined) return conversationSummary(c, me, raw);
  return { id, kind: conversationKind({ id } as RawConversation), title: id, unread: false, unread_confirmed: false };
}

/** The Substrate path: page through Teams' own search, filter, and map hits to the scan's shape. */
async function substrateSearch(
  substrate: SubstrateSearch,
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  { query, target, fromLc, cutoff, cap }: SearchArgs,
  force: boolean,
): Promise<Record<string, unknown>> {
  const found: { hit: SubstrateMessage; at: number }[] = [];
  let pages = 0;
  let resultsRead = 0;
  let filteredOut = 0;
  let unmapped = 0;
  let total: number | undefined;
  let more = false;

  for (let offset = 0; pages < MAX_SUBSTRATE_PAGES; offset += SUBSTRATE_PAGE_SIZE) {
    const page = await substrate.searchMessages(
      {
        query,
        from: offset,
        size: SUBSTRATE_PAGE_SIZE,
        // Both restrictions are verified server-side; the checks below stay as a backstop.
        ...(target !== undefined ? { conversationId: target.id } : {}),
        sentAfter: cutoff,
      },
      force ? { force } : undefined,
    );
    pages++;
    resultsRead += page.results.length;
    unmapped += page.unmapped;
    total ??= page.total;
    more = page.moreAvailable;

    let anyInWindow = false;
    for (const r of page.results) {
      const at = r.time !== undefined ? Date.parse(r.time) : NaN;
      if (Number.isFinite(at) && at < cutoff) {
        filteredOut++;
        continue;
      }
      // An undated hit counts as inside the window: Substrate returned it, and dropping it would hide it.
      anyInWindow = true;
      if (target !== undefined && r.conversationId !== target.id) filteredOut++;
      else if (fromLc !== undefined && !(r.sender ?? '').toLowerCase().includes(fromLc)) filteredOut++;
      else found.push({ hit: r, at });
    }
    // Newest first, so a page wholly before the cutoff means everything after it is too.
    if (!more || page.results.length === 0 || found.length > cap || !anyInWindow) break;
  }

  found.sort((a, b) => (Number.isFinite(b.at) ? b.at : 0) - (Number.isFinite(a.at) ? a.at : 0));
  const kept = found.slice(0, cap);

  const listed = new Map(raw.map((c) => [c.id, c]));
  const summaries = [...new Map(kept.map(({ hit }) => [hit.conversationId, summaryFor(hit.conversationId, listed, raw, me)])).values()];
  // Only listed conversations are titled (an unlisted one has no listing entry to title from).
  const titled = new Map(
    (await enrichTitles(chatsvc, raw, me, summaries.filter((s) => listed.has(s.id)))).map((s) => [s.id, s]),
  );
  const bySummary = new Map(summaries.map((s) => [s.id, titled.get(s.id) ?? s]));

  const hits: Hit[] = kept.map(({ hit }) => ({
    conversation: hitConversation(bySummary.get(hit.conversationId) ?? summaryFor(hit.conversationId, listed, raw, me)),
    message: {
      ...(hit.messageId !== undefined ? { id: hit.messageId } : {}),
      ...(hit.time !== undefined ? { time: hit.time } : {}),
      sender: hit.sender ?? '',
      // No fromMe: a Substrate hit names its sender by display name and SMTP address only.
      text: hit.text,
      excerpt: true as const,
      ...(hit.subject !== undefined ? { subject: hit.subject } : {}),
      ...(hit.replyTo !== undefined ? { replyTo: hit.replyTo } : {}),
    },
  }));

  // Exact when Substrate ran out; otherwise its own (approximate) count, which
  // only means anything when nothing was filtered out of what it returned.
  const totalHits = more && filteredOut === 0 && total !== undefined ? Math.max(total, found.length) : found.length;
  const sinceIso = new Date(cutoff).toISOString();
  const pageCapped = pages >= MAX_SUBSTRATE_PAGES && more && found.length <= cap;

  return {
    total_hits: totalHits,
    hits,
    ...(totalHits > kept.length
      ? {
          truncated: {
            total: totalHits,
            returned: kept.length,
            hint: `Only the newest ${kept.length} hits were returned; pass a larger limit (max ${MAX_LIMIT}) or narrow with conversation, from or since.`,
          },
        }
      : {}),
    substrate: {
      since: sinceIso,
      pages_read: pages,
      results_read: resultsRead,
      results_filtered_out: filteredOut,
      ...(unmapped > 0 ? { results_unmapped: unmapped } : {}),
      more_available: more,
      ...(pageCapped
        ? {
            hint:
              `Read the first ${MAX_SUBSTRATE_PAGES * SUBSTRATE_PAGE_SIZE} Substrate results and stopped; ` +
              'narrow with conversation, from or a later since, or a more specific query.',
          }
        : {}),
    },
  };
}

/** The scan path: read the most recently active conversations back to the cutoff and match locally. */
async function scanSearch(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  { query, target, fromLc, cutoff, cap }: SearchArgs,
  maxConversations: number | undefined,
): Promise<Record<string, unknown>> {
  // The conversations to read: the one named, else every one with activity
  // inside the window, most recent first, hidden channels left out.
  let targets: ConversationSummary[];
  let outsideWindow = 0;
  if (target !== undefined) {
    targets = [target];
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
  const maxScanned = maxConversations ?? DEFAULT_SCANNED_CONVERSATIONS;
  const scanned = targets.slice(0, maxScanned);
  const notScanned = targets.length - scanned.length;

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
  const kept = found.slice(0, cap);

  // Only conversations that are returned are titled (an untitled chat's title costs a read).
  const named = [...new Map([...kept.map((h) => h.summary), ...partial, ...failed.map((f) => f.summary)].map((s) => [s.id, s])).values()];
  const titled = new Map((await enrichTitles(chatsvc, raw, me, named)).map((s) => [s.id, s]));
  const final = (s: ConversationSummary): ConversationSummary => titled.get(s.id) ?? s;

  const hits: Hit[] = kept.map((h) => ({ conversation: hitConversation(final(h.summary)), message: hitMessage(h.message) }));
  const sinceIso = new Date(cutoff).toISOString();
  const leftOut = notScanned > 0 || partial.length > 0;

  return {
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
      ...(partial.length > 0 ? { partially_scanned: partial.map((s) => ({ id: s.id, title: final(s).title, ...participantsOf(final(s)) })) } : {}),
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
  };
}
