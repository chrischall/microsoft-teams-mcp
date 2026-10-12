/**
 * `teams_get_changes` — which conversations got new messages since a cursor,
 * across every chat and channel, in one call.
 *
 * The cursor is a time: an opaque base64url `{"t": <epoch ms>}` this tool hands
 * back as `next_cursor` (a plain ISO 8601 time works too; none means the last
 * 24 hours). One listing nominates the conversations that changed after it
 * (`lastChangeMs` — a channel reply shows only in `lastUpdatedMessageId`), and
 * each is read newest change first down to it, a few at a time. A channel
 * hidden in Teams is the exception: its listing entry stops moving (live,
 * 2026-10-12: two posts from that day under a `lastUpdatedMessageId` three
 * days old), so with `include_hidden` every hidden channel is read. What comes
 * back is counts, not messages: per conversation its summary, how many real
 * messages arrived (system events and deleted messages are not counted) and,
 * for a channel, the threads they landed in — including a thread whose root is
 * older than the cursor. Read the messages themselves with
 * `teams_read_conversation`.
 *
 * Reads stop on a message's last CHANGE (`version`), not its arrival: live
 * (2026-10-12), pages come newest change first, so an old message edited or
 * reacted to after the cursor is served ahead of new ones, and stopping at the
 * first message older than the cursor lost 104 of 294 new messages over 24
 * hours; stopping on the change lost none.
 *
 * `next_cursor` is at-least-once: it is the newest arrival seen, but never
 * past the newest change the listing knew of (a message that lands mid-call in
 * a conversation the listing called quiet is then not skipped), and it stays
 * at the input time when any read failed. Nothing writes the read marker.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { mapWithConcurrency, McpToolError, toolAnnotations } from '@chrischall/mcp-utils';
import type { RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import {
  conversationSummary,
  isUserMessage,
  lastChangeMs,
  normaliseMessage,
  type ConversationKind,
  type ConversationSummary,
} from '../model.js';
import { isUnreachable } from './confirm.js';
import { includeHiddenArg } from './listing.js';
import { enrichTitles } from './titles.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';
import { CURSOR_NO_CACHE_DESCRIPTION } from './cursor.js';
import { DEFAULT_LIMIT, limitArg } from './limit.js';
import { KIND_FILTERS, type KindFilter } from './unread.js';

const MAX_CONVERSATIONS = 200;

/** How far back a call with no cursor looks. */
export const DEFAULT_WINDOW_MS = 86_400_000;

/**
 * Raw messages read per conversation before giving up; past it the
 * conversation is marked truncated and its count is a lower bound.
 */
const MAX_RAW_MESSAGES = 1000;

/** Messages per page: a conversation that changed usually changed by a few. */
const PAGE_SIZE = 50;

/** Conversations read at once. */
const READ_CONCURRENCY = 4;

/** `cursor` as handed out: base64url JSON `{ t }`, `t` epoch ms. */
export function encodeChangesCursor(t: number): string {
  return Buffer.from(JSON.stringify({ t }), 'utf8').toString('base64url');
}

function badCursor(cursor: string): McpToolError {
  return new McpToolError(`"${cursor}" is not a changes cursor.`, {
    hint: "Pass a previous result's next_cursor unchanged, or an ISO 8601 time such as 2026-10-12T09:00:00Z.",
  });
}

/** The epoch ms a `cursor` names: a token from {@link encodeChangesCursor}, or an ISO 8601 time. */
export function decodeChangesCursor(cursor: string): number {
  if (/^[A-Za-z0-9_-]+$/.test(cursor)) {
    let parsed: unknown;
    try {
      parsed = JSON.parse(Buffer.from(cursor, 'base64url').toString('utf8'));
    } catch {
      parsed = undefined;
    }
    if (parsed !== undefined) {
      const t = (parsed as { t?: unknown } | null)?.t;
      if (typeof t === 'number' && Number.isFinite(t) && t >= 0) return t;
      throw badCursor(cursor);
    }
  }
  // Only a string that looks like a date: Date.parse is lenient with words.
  const t = /^\d{4}-\d{2}-\d{2}/.test(cursor) ? Date.parse(cursor) : NaN;
  if (Number.isNaN(t)) throw badCursor(cursor);
  return t;
}

/** A message's last change (epoch ms): its `version`, else its arrival (the id). */
function changedMs(m: RawMessage): number {
  const v = Number(m.version);
  const id = Number(m.id);
  return Math.max(Number.isFinite(v) ? v : 0, Number.isFinite(id) ? id : 0);
}

/** One channel thread the new messages landed in. */
interface ChangedThread {
  root_id: string;
  /** The root post itself arrived after the cursor. */
  new_root: boolean;
  /** Replies that arrived after the cursor. */
  new_reply_count: number;
}

interface ChangeRead {
  /** Real, undeleted messages that arrived after the cursor. */
  count: number;
  fromOthers: number;
  /** Arrival of the newest of them (epoch ms); 0 when there are none. */
  newest: number;
  newestTime?: string;
  /** The newest arrival of ANY message seen past the cursor, system events included. */
  maxSeen: number;
  threads?: ChangedThread[];
  truncated: boolean;
}

/** Read one conversation down to `t` and count what arrived after it. */
async function readChanges(
  chatsvc: TeamsDeps['chatsvc'],
  summary: ConversationSummary,
  t: number,
  me: string,
): Promise<ChangeRead> {
  let capped = false;
  const fetched = await chatsvc.getMessages(summary.id, {
    stopAt: (m) => changedMs(m) <= t,
    max: MAX_RAW_MESSAGES,
    pageSize: PAGE_SIZE,
    onCapped: () => (capped = true),
  });
  const arrived = fetched.filter((m) => Number(m.id) > t);
  const maxSeen = arrived.reduce((n, m) => Math.max(n, Number(m.id)), 0);
  const fresh = arrived
    .filter(isUserMessage)
    .map((m) => normaliseMessage(m, me))
    .filter((m) => m.deleted !== true)
    .sort((a, b) => Number(b.id) - Number(a.id));

  let threads: ChangedThread[] | undefined;
  if (summary.kind === 'channel') {
    // Newest first, so threads come in order of their newest new message.
    const byRoot = new Map<string, ChangedThread>();
    for (const m of fresh) {
      const rootId = m.replyTo ?? m.id;
      let thread = byRoot.get(rootId);
      if (thread === undefined) byRoot.set(rootId, (thread = { root_id: rootId, new_root: false, new_reply_count: 0 }));
      if (m.replyTo === undefined) thread.new_root = true;
      else thread.new_reply_count++;
    }
    threads = [...byRoot.values()];
  }

  return {
    count: fresh.length,
    fromOthers: fresh.filter((m) => !m.fromMe).length,
    newest: fresh.length > 0 ? Number(fresh[0].id) : 0,
    ...(fresh.length > 0 ? { newestTime: fresh[0].time } : {}),
    maxSeen,
    ...(threads !== undefined ? { threads } : {}),
    truncated: capped,
  };
}

export function registerChangesTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_get_changes',
    {
      title: 'Get Teams conversations with new messages',
      description:
        'List the chats and channels that got new messages since a cursor, most recent first — counts, not ' +
        'the messages: each with its summary (id, kind, title, team), new_message_count (real messages ' +
        'only — call recordings and other system events, and deleted messages, are not counted), ' +
        'new_from_others_count, newest_time and, for a channel, threads: one { root_id, new_root, ' +
        'new_reply_count } per thread the new messages landed in, including a thread whose root post is ' +
        'older than the cursor. Read the messages with teams_read_conversation (since = the since this ' +
        'returns). Without cursor it looks back 24 hours. next_cursor is an opaque token — pass it back as ' +
        'cursor to get only what arrived since this call (cursor also takes a plain ISO 8601 time). It is ' +
        'at-least-once: a conversation can come back again but none is skipped; it stays at the input ' +
        'time when a read failed. Use the same kinds and include_hidden with a cursor as the call that ' +
        'returned it. Channels hidden in Teams are left out unless include_hidden (then every hidden channel ' +
        'is read, since Teams stops updating their listing); one that is gone or no ' +
        'longer readable (404/403) is skipped and counted in skipped_unreachable; one whose read failed ' +
        'is listed in errors. limit caps the conversations returned; the rest are named, with their ' +
        'new_message_count, in more_conversations, and total_new_messages counts them all. A conversation ' +
        `with more than ${MAX_RAW_MESSAGES} new messages is marked truncated (its count is a lower bound). ` +
        CURSOR_NO_CACHE_DESCRIPTION +
        'This does NOT mark anything read — the user\'s Teams unread badges are unaffected. ' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        cursor: z
          .string()
          .min(1)
          .optional()
          .describe(
            "A previous result's next_cursor, or an ISO 8601 time (e.g. 2026-10-12T09:00:00Z): only messages " +
              'that arrived after it count. Default: 24 hours ago.',
          ),
        kinds: z
          .array(z.enum(Object.keys(KIND_FILTERS) as [KindFilter, ...KindFilter[]]))
          .min(1)
          .optional()
          .describe(
            'Only these kinds of conversation (default all): chat = every group, meeting and 1:1 ' +
              'chat; channel; meeting = meeting chats only; oneOnOne = 1:1 chats only.',
          ),
        include_hidden: includeHiddenArg.describe(
          'Also include channels the user has hidden in Teams (default false: Teams shows them nowhere).',
        ),
        limit: limitArg(MAX_CONVERSATIONS, 'conversations', 'most recent activity first kept'),
      },
    },
    async ({ cursor, kinds, include_hidden, limit }) => {
      try {
        const t = cursor !== undefined ? decodeChangesCursor(cursor) : Date.now() - DEFAULT_WINDOW_MS;
        const fetched_at = new Date().toISOString();
        const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
        const allowed = new Set<ConversationKind>((kinds ?? ['chat', 'channel']).flatMap((k) => KIND_FILTERS[k]));
        const lastChange = new Map(raw.map((c) => [c.id, lastChangeMs(c)] as const));
        // Summaries are built against the FULL listing so a channel can find its team's name. A
        // hidden channel's listing entry stops moving (live: days stale while posts arrived), so
        // with include_hidden every hidden channel is read; otherwise the listing nominates.
        const candidates = raw
          .map((c) => conversationSummary(c, me, raw))
          .filter((s) => allowed.has(s.kind))
          .filter((s) => (s.hidden === true ? include_hidden === true : (lastChange.get(s.id) ?? 0) > t));

        type Result =
          | { kind: 'read'; read: ChangeRead }
          | { kind: 'unreachable' }
          | { kind: 'failed'; error: unknown };
        const results = await mapWithConcurrency(candidates, READ_CONCURRENCY, async (summary): Promise<Result> => {
          try {
            return { kind: 'read', read: await readChanges(chatsvc, summary, t, me) };
          } catch (error) {
            return isUnreachable(error) ? { kind: 'unreachable' } : { kind: 'failed', error };
          }
        });

        let skipped_unreachable = 0;
        let maxSeen = t;
        let anyRead = false;
        const changed: { summary: ConversationSummary; read: ChangeRead }[] = [];
        const failed: { summary: ConversationSummary; error: unknown }[] = [];
        results.forEach((r, i) => {
          if (r.kind === 'unreachable') skipped_unreachable++;
          else if (r.kind === 'failed') failed.push({ summary: candidates[i], error: r.error });
          else {
            anyRead = true;
            maxSeen = Math.max(maxSeen, r.read.maxSeen);
            if (r.read.count > 0) changed.push({ summary: candidates[i], read: r.read });
          }
        });
        // Every read failed: a failure of the call (a revoked token, a rate limit), not of one conversation.
        if (failed.length > 0 && !anyRead) throw failed[0].error;
        changed.sort((a, b) => b.read.newest - a.read.newest);

        // Never past the newest change the listing saw: a quiet conversation's mid-call message stays ahead.
        const listed = Math.max(t, ...[...lastChange.values()].map((n) => n ?? 0));
        const next = failed.length > 0 ? t : Math.min(maxSeen, listed);

        const cap = limit ?? DEFAULT_LIMIT;
        const shown = changed.slice(0, cap);
        const rest = changed.slice(cap);
        const titled = await enrichTitles(chatsvc, raw, me, [
          ...shown.map((c) => c.summary),
          ...rest.map((c) => c.summary),
          ...failed.map((f) => f.summary),
        ]);
        const conversations = shown.map(({ read }, i) => ({
          conversation: titled[i],
          new_message_count: read.count,
          new_from_others_count: read.fromOthers,
          ...(read.newestTime !== undefined ? { newest_time: read.newestTime } : {}),
          ...(read.threads !== undefined ? { threads: read.threads } : {}),
          ...(read.truncated ? { truncated: true as const } : {}),
        }));
        const more = rest.map(({ read }, i) => ({
          id: titled[shown.length + i].id,
          title: titled[shown.length + i].title,
          new_message_count: read.count,
        }));
        const errors = failed.map((f, i) => ({
          conversation: titled[changed.length + i],
          error: wrapApiError(f.error, 'read this conversation').message,
        }));

        return untrustedResult({
          since: new Date(t).toISOString(),
          total_new_messages: changed.reduce((n, c) => n + c.read.count, 0),
          conversations,
          ...(more.length > 0 ? { more_conversations: more } : {}),
          skipped_unreachable,
          ...(errors.length > 0 ? { errors } : {}),
          next_cursor: next === t && cursor !== undefined ? cursor : encodeChangesCursor(next),
          next_cursor_time: new Date(next).toISOString(),
          fetched_at,
        });
      } catch (err) {
        throw wrapApiError(err, 'get changed conversations');
      }
    },
  );
}
