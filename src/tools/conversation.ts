/**
 * `teams_read_conversation` — read any chat or channel, by id or name, from
 * the Teams chat-service API. Nothing has to be open in the browser, and the
 * read marker is never written, so the user's unread badges are untouched.
 *
 * Paging runs newest first and stops at the first boundary that applies: the
 * read horizon (`unread_only`), the `after_message_id` cursor, the `since`
 * cutoff, or one message past `limit`. The result is turned oldest first; a channel's posts are grouped
 * under their root post, and `format` (`format.ts`) trims what is shown. An
 * array of conversations is resolved against one listing and read a few at a
 * time, each failure reported beside the reads that worked, and an ambiguous
 * name returned in its place as its candidates.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { mapWithConcurrency, McpToolError, toolAnnotations } from '@chrischall/mcp-utils';
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
} from '../model.js';
import { confirmedSummary } from './confirm.js';
import { enrichTitles } from './titles.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';
import { DEFAULT_LIMIT, limitArg, makeLimitedStopAt, sinceArg } from './limit.js';
import { resolveTimeZone, TIME_ZONE_DESCRIPTION, tzArg } from './timezone.js';
import { afterMessageIdArg, CURSOR_NO_CACHE_DESCRIPTION, newestId } from './cursor.js';
import { FORMAT_DESCRIPTION, formatArg, formatBody, formatConversation, type ReadFormat } from './format.js';
import { fetchOutsideRoots, MAX_ROOT_FETCHES, rootBudget, type RootBudget } from './roots.js';

const MAX_MESSAGES = 500;

/**
 * Raw messages read before giving up (20 pages of 200). The real bound is the
 * `stopAt` boundary; this only caps a conversation of nothing but system events.
 */
const MAX_RAW_MESSAGES = 4000;

/** Candidates an ambiguous name lists, as an error or as data. */
const MAX_CANDIDATES = 10;

/** Arrival time in epoch ms: a message's id IS its arrival ms; the ISO fields are the fallback. */
export function arrivalMs(m: RawMessage): number {
  const n = Number(m.id);
  if (Number.isFinite(n) && n > 0) return n;
  return Date.parse(m.originalarrivaltime ?? m.composetime ?? '');
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
  const how = s.match_reason !== undefined ? `, matched by ${s.match_reason}` : '';
  return `"${s.title}"${where} (${s.kind}, id ${s.id}${how})`;
}

/**
 * A name that matched several conversations. It is an `McpToolError` — the
 * write tools let it through, as a write must never guess — that also carries
 * every match, so `teams_read_conversation` can return the candidates as data.
 */
export class AmbiguousConversationError extends McpToolError {
  constructor(
    readonly query: string,
    readonly matches: readonly ConversationSummary[],
    verb: string,
  ) {
    const named = matches.slice(0, MAX_CANDIDATES).map(describeCandidate).join('; ');
    const more = matches.length > MAX_CANDIDATES ? ` (first ${MAX_CANDIDATES} shown)` : '';
    super(`"${query}" matches ${matches.length} conversations${more}: ${named}.`, {
      hint: `Pass the id of the one to ${verb} as conversation.`,
    });
    this.name = 'AmbiguousConversationError';
  }
}

/**
 * An ambiguous name as a read's result: the first 10 candidates, each with
 * what tells them apart and how it matched, and how many matched in all.
 */
export function ambiguousResult(err: AmbiguousConversationError): Record<string, unknown> {
  return {
    ambiguous: true,
    query: err.query,
    match_count: err.matches.length,
    candidates: err.matches.slice(0, MAX_CANDIDATES).map((s) => ({
      id: s.id,
      title: s.title,
      kind: s.kind,
      ...(s.team !== undefined ? { team: s.team } : {}),
      ...(s.lastActivity !== undefined ? { lastActivity: s.lastActivity } : {}),
      ...(s.match_reason !== undefined ? { match_reason: s.match_reason } : {}),
    })),
    hint: 'pass one id as conversation',
  };
}

/**
 * The one conversation a `conversation` argument names: an exact id, else a
 * unique title/team/participant substring — never what was said in a chat,
 * since this names the target of sends and reactions too. The match carries
 * its `match_reason`. None is an `McpToolError`; several is an
 * {@link AmbiguousConversationError} naming up to 10 candidates, each with how
 * it matched, so the caller can pass an id. `verb` fills the hint ("Pass the
 * id of the one to read").
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
  if (matches.length > 1) throw new AmbiguousConversationError(conversation, matches, verb);
  return matches[0];
}


/**
 * A resolver of `conversation` arguments against one listing: an exact id is
 * {@link resolveConversation}d with just that one titled, a self-chat alias
 * (`notes`, `self`) with nothing titled, and a name against the whole listing
 * with untitled group chats titled by their participants and 1:1s by the
 * other person (`titles.ts`) — so a name can pick one of several chats that
 * share a last sender, or a 1:1 the user wrote in last. The whole-listing
 * titling runs once, however many names the resolver is asked for.
 */
function listedResolver(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  verb: string,
): (conversation: string) => Promise<ConversationSummary> {
  const summaries = summariseAll(raw, me);
  let titled: Promise<ConversationSummary[]> | undefined;
  return async (conversation) => {
    if (SELF_CHAT_ALIASES.has(conversation.trim().toLowerCase())) return resolveConversation(summaries, conversation, verb);
    const exact = summaries.find((s) => s.id === conversation);
    if (exact) return resolveConversation(await enrichTitles(chatsvc, raw, me, [exact]), conversation, verb);
    titled ??= enrichTitles(chatsvc, raw, me, summaries);
    return resolveConversation(await titled, conversation, verb);
  };
}

/** {@link resolveConversation} against the whole listing, titled as {@link listedResolver} says. */
export async function resolveListed(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  conversation: string,
  verb: string,
): Promise<ConversationSummary> {
  return listedResolver(chatsvc, raw, me, verb)(conversation);
}

/** Conversations one call may read. */
export const MAX_CONVERSATIONS = 20;

/** Conversations read at once in a multi-conversation read. */
export const READ_CONCURRENCY = 4;

interface ReadOptions {
  unreadOnly: boolean;
  since?: string;
  afterMessageId?: string;
  limit: number;
  timeZone?: string;
  format: ReadFormat;
  /** The call's root fetches, shared by every conversation it reads (`roots.ts`). */
  roots: RootBudget;
}

/**
 * Read one resolved conversation: what a single-conversation call returns,
 * less `fetched_at`, which the call stamps once.
 */
async function readOne(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  summary: ConversationSummary,
  { unreadOnly, since, afterMessageId, limit, timeZone, format, roots }: ReadOptions,
): Promise<Record<string, unknown>> {
  // The self-chat resolves even when the listing does not carry it.
  const conv = raw.find((c) => c.id === summary.id) ?? { id: summary.id };
  const horizon = readHorizon(conv);
  const cutoff = since !== undefined ? Date.parse(since) : undefined;
  const cursor = afterMessageId !== undefined ? Number(afterMessageId) : undefined;
  const wanted = (m: RawMessage): boolean => isUserMessage(m) && (!unreadOnly || isUnread(m, horizon, me));

  const limited = makeLimitedStopAt(unreadOnly ? horizon : undefined, wanted, limit);
  // A cursor or `since` cutoff above the read horizon leaves unread messages unseen.
  let windowCut = false;
  let capped = false;
  const stopAt = (m: RawMessage): boolean => {
    // The read horizon (checked inside `limited`) wins over the cursor and `since`, as the nearer boundary.
    const atHorizon = unreadOnly && Number(m.id) <= horizon;
    const at = arrivalMs(m);
    const atCursor = cursor !== undefined && Number(m.id) <= cursor;
    const atSince = cutoff !== undefined && Number.isFinite(at) && at < cutoff;
    if (!atHorizon && (atCursor || atSince)) {
      if (Number(m.id) > horizon) windowCut = true;
      return true;
    }
    return limited.stopAt(m);
  };
  const fetched = await chatsvc.getMessages(summary.id, {
    stopAt,
    max: MAX_RAW_MESSAGES,
    onCapped: () => (capped = true),
  });

  const returned = fetched.filter(wanted);
  const next_cursor = newestId(returned) ?? afterMessageId;
  const messages = oldestFirst(returned).map((m) => normaliseMessage(m, me, { timeZone }));
  const unread_count = fetched.filter((m) => wanted(m) && isUnread(m, horizon, me)).length;
  const outsideRoots =
    summary.kind === 'channel'
      ? await fetchOutsideRoots(chatsvc, summary.id, messages, me, { budget: roots, timeZone })
      : undefined;
  const truncated = limited.truncated || capped;
  // Whether every message past the read horizon was seen, so unread_count is the whole count.
  const exact = !truncated && !windowCut;

  return {
    conversation: formatConversation(confirmUnread(summary, unread_count, exact), format),
    unread_count,
    ...(exact ? {} : { unread_count_scope: 'window' as const }),
    ...formatBody(summary.kind, messages, format, outsideRoots),
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
    ...(next_cursor !== undefined ? { next_cursor } : {}),
  };
}

/** One `conversation` of a multi-conversation read that could not be resolved or read. */
interface ReadError {
  /** The argument as passed. */
  conversation: string;
  /** The conversation it resolved to, when it got that far. */
  id?: string;
  error: string;
  hint?: string;
}

function readError(conversation: string, err: unknown, id?: string): ReadError {
  const e = wrapApiError(err, 'read this conversation');
  return {
    conversation,
    ...(id !== undefined ? { id } : {}),
    error: e.message,
    ...(e.hint !== undefined ? { hint: e.hint } : {}),
  };
}

/**
 * Read several conversations: each argument resolved against one listing, the
 * distinct conversations read {@link READ_CONCURRENCY} at a time, in the order
 * asked. An ambiguous name takes its place among them as the same
 * {@link ambiguousResult} a single read returns; one that cannot be resolved
 * or read goes into `errors`; the rest still come back.
 */
async function readMany(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  conversations: readonly string[],
  opts: ReadOptions,
): Promise<{ conversations: Record<string, unknown>[]; errors: ReadError[] }> {
  const resolve = listedResolver(chatsvc, raw, me, 'read');
  const errors: ReadError[] = [];
  /** One slot per distinct conversation asked: a resolved one to read, or an ambiguous name's candidates. */
  const slots: ({ asked: string; summary: ConversationSummary } | { ambiguous: Record<string, unknown> })[] = [];
  const seen = new Set<string>();
  for (const asked of new Set(conversations)) {
    try {
      const summary = await resolve(asked);
      // A conversation named twice (by id and by name, say) is read once.
      if (seen.has(summary.id)) continue;
      seen.add(summary.id);
      slots.push({ asked, summary });
    } catch (err) {
      if (err instanceof AmbiguousConversationError) slots.push({ ambiguous: ambiguousResult(err) });
      else errors.push(readError(asked, err));
    }
  }
  const reads = await mapWithConcurrency(slots, READ_CONCURRENCY, async (slot) => {
    if ('ambiguous' in slot) return { ok: true as const, result: slot.ambiguous };
    try {
      return { ok: true as const, result: await readOne(chatsvc, raw, me, slot.summary, opts) };
    } catch (err) {
      return { ok: false as const, error: readError(slot.asked, err, slot.summary.id) };
    }
  });
  const results: Record<string, unknown>[] = [];
  for (const r of reads) {
    if (r.ok) results.push(r.result);
    else errors.push(r.error);
  }
  // Resolution failures were found first; list every error in the order asked.
  const order = new Map(conversations.map((c, i) => [c, i] as const));
  errors.sort((a, b) => (order.get(a.conversation) ?? 0) - (order.get(b.conversation) ?? 0));
  return { conversations: results, errors };
}

export function registerConversationTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_read_conversation',
    {
      title: 'Read a Teams chat or channel',
      description:
        'Read the messages of any chat or channel the signed-in user is in, by id (from ' +
        'teams_list_chats / teams_list_teams_and_channels) or by name, without opening it in ' +
        'Teams. "48:notes" (or the name "notes" / "self") is the user\'s own notes self-chat. Messages come oldest first, each with id, time (UTC, ISO 8601), sender, fromMe, ' +
        'text (never truncated; table cells separated by " | "), subject, edited/deleted flags and attachments; a channel is returned as threads, ' +
        'each root post with its replies. A root older than the window is fetched and marked outside_window: ' +
        'true (it is context, not counted or cursored); up to ' +
        `${MAX_ROOT_FETCHES} roots are fetched per call, and any beyond that or that cannot be fetched is ` +
        '{ id, missing: true }. ' +
        'unread_only returns just the messages the user has not read; since stops at a time; ' +
        'limit keeps the newest (truncated says when more exist; unread_count_scope "window" says ' +
        'unread_count only covers what limit/since/after_message_id let through). after_message_id returns only ' +
        'messages newer than that id; next_cursor is the newest returned message id (or after_message_id ' +
        'when nothing is new) — pass it back as after_message_id to read only what arrived since. ' +
        'after_message_id takes message-id cursors only: to read what teams_get_changes counted, pass that ' +
        "row's after_message_id, never teams_get_changes' next_cursor (a time token for teams_get_changes). " +
        CURSOR_NO_CACHE_DESCRIPTION +
        FORMAT_DESCRIPTION +
        `conversation may also be an array of up to ${MAX_CONVERSATIONS} ids/names: the result is then ` +
        '{ conversations: [...], errors, fetched_at }, one single-read result per conversation (the other ' +
        'arguments apply to each); an ambiguous name takes its place among them as the same { ambiguous: true, ' +
        '... } object a single read returns, and one that cannot be resolved or read is listed in errors ' +
        'while the rest still come back. ' +
        'This does NOT mark anything ' +
        'read — the user\'s Teams unread badges are unaffected. Attachments are listed per message (images by ref, files by ' +
        'name) and fetched with teams_get_attachment. A name that matches several conversations is ' +
        'not an error: the result is {ambiguous: true, query, match_count, candidates (up to 10, each ' +
        'with id, title, kind, team, lastActivity and match_reason), hint} and no messages — read again ' +
        'with one candidate\'s id. ' +
        TIME_ZONE_DESCRIPTION +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        conversation: z
          .union([z.string().min(1), z.array(z.string().min(1)).min(1).max(MAX_CONVERSATIONS)])
          .describe(
            'The conversation id ("48:notes", "notes" or "self" for the self-chat), or a case-insensitive ' +
              "substring of the chat title, channel name, team name or an untitled chat's participant names; " +
              `or an array of up to ${MAX_CONVERSATIONS} of them to read several at once.`,
          ),
        unread_only: z
          .boolean()
          .optional()
          .describe('Only return messages the user has not read yet (default false). Never marks them read.'),
        since: sinceArg,
        after_message_id: afterMessageIdArg,
        limit: limitArg(MAX_MESSAGES, 'messages'),
        tz: tzArg,
        format: formatArg,
      },
    },
    async ({ conversation, unread_only, since, after_message_id, limit, tz, format }) => {
      try {
        const opts: ReadOptions = {
          unreadOnly: unread_only ?? false,
          since,
          afterMessageId: after_message_id,
          limit: limit ?? DEFAULT_LIMIT,
          timeZone: resolveTimeZone(tz),
          format: format ?? 'full',
          roots: rootBudget(),
        };
        const fetched_at = new Date().toISOString();
        const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
        if (Array.isArray(conversation)) {
          const many = await readMany(chatsvc, raw, me, conversation, opts);
          return untrustedResult({
            conversations: many.conversations,
            ...(many.errors.length > 0 ? { errors: many.errors } : {}),
            fetched_at,
          });
        }
        const summary = await resolveListed(chatsvc, raw, me, conversation, 'read');
        return untrustedResult({ ...(await readOne(chatsvc, raw, me, summary, opts)), fetched_at });
      } catch (err) {
        // A read can safely hand back the candidates; the write tools keep refusing.
        if (err instanceof AmbiguousConversationError) return untrustedResult(ambiguousResult(err));
        throw wrapApiError(err, 'read the conversation');
      }
    },
  );
}
