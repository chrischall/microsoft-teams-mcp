/**
 * Output-size modes for the read tools (`teams_list_messages`,
 * `teams_list_unread_messages`): how much of each conversation a result carries.
 *
 * - `full` (default) — every message field and the whole conversation summary,
 *   exactly as before.
 * - `compact` — each message only `{ id, time, time_local?, sender, text }`
 *   (plus `subject` when it has one): no reactions, attachments, edited/deleted
 *   flags, `fromMe` or `replyTo` (a channel's threads already say what replies
 *   to what), and the conversation only `{ id, kind, title, team? }`.
 * - `summary` — no messages at all: per conversation `message_count`, the
 *   newest message's time and, for a channel, one line per thread (root id,
 *   subject, reply count, newest time).
 *
 * A channel thread whose root is older than the window shows the root fetched
 * by id (`roots.ts`), marked `outside_window` (`root_outside_window` in
 * summary), or — when it could not be fetched — a `{ id, missing: true }`
 * placeholder (`root_missing`).
 *
 * Every format keeps the conversation's id, and everything else a read reports
 * (unread_count, next_cursor, truncated, …) is the same in all three: the
 * format only changes how the messages are shown, never which were read.
 */
import { z } from 'zod';
import type { ConversationKind, ConversationSummary, NormalisedMessage } from '../model.js';

export const READ_FORMATS = ['full', 'compact', 'summary'] as const;
export type ReadFormat = (typeof READ_FORMATS)[number];

export const formatArg = z
  .enum(READ_FORMATS)
  .optional()
  .describe(
    'How much to return (default full). full: every message field. compact: each message only id, time, ' +
      'time_local, sender, text (and subject when set) — no reactions, attachments or edited/deleted flags. ' +
      'summary: no messages, just per conversation its counts, newest message time and, for a channel, ' +
      'each thread\'s root id, subject, reply count and newest time.',
  );

/** Shared description text for the `format` argument. */
export const FORMAT_DESCRIPTION =
  'format "compact" trims each message to id, time, sender and text, and the conversation to id, kind ' +
  'and title; format "summary" returns no messages, only message_count, newest_time and, for a channel, ' +
  'threads of { root_id, subject, reply_count, newest_time } — every format keeps the conversation id. ';

/** A root the read's window did not hold and that could not be fetched. */
export interface MissingRoot {
  id: string;
  missing: true;
}

/** A root fetched by id from outside the read's window: context, not part of the read. */
export type OutsideRoot = NormalisedMessage & { outside_window: true };

export type ThreadRoot = NormalisedMessage | OutsideRoot | MissingRoot;
export interface Thread {
  root: ThreadRoot;
  replies: NormalisedMessage[];
}

/** Roots fetched from outside a read's window, by id (`roots.ts`). */
export type OutsideRoots = ReadonlyMap<string, NormalisedMessage>;

const NO_ROOTS: OutsideRoots = new Map();

/**
 * Group oldest-first channel messages under their root post, in order of first
 * appearance. A reply whose root is older than the window still appears, under
 * that root as fetched (`roots`, marked `outside_window`), else a
 * `{ id, missing: true }` placeholder.
 */
export function groupThreads(messages: readonly NormalisedMessage[], roots: OutsideRoots = NO_ROOTS): Thread[] {
  const byRoot = new Map<string, Thread>();
  for (const m of messages) {
    const rootId = m.replyTo ?? m.id;
    let thread = byRoot.get(rootId);
    if (thread === undefined) {
      const outside = roots.get(rootId);
      const root: ThreadRoot = outside !== undefined ? { ...outside, outside_window: true } : { id: rootId, missing: true };
      thread = { root, replies: [] };
      byRoot.set(rootId, thread);
    }
    if (m.replyTo === undefined) thread.root = m;
    else thread.replies.push(m);
  }
  return [...byRoot.values()];
}

export interface CompactMessage {
  id: string;
  time: string;
  time_local?: string;
  sender: string;
  text: string;
  subject?: string;
}

function compactMessage(m: NormalisedMessage): CompactMessage {
  return {
    id: m.id,
    time: m.time,
    ...(m.time_local !== undefined ? { time_local: m.time_local } : {}),
    sender: m.sender,
    text: m.text,
    ...(m.subject !== undefined ? { subject: m.subject } : {}),
  };
}

/**
 * One message of a flat, cross-conversation list (`teams_list_unread_messages` with
 * `from`): whole in full; compact, but keeping `replyTo`, since no thread
 * grouping says what a channel reply answers. Summary has no messages.
 */
export function formatFlatMessage(m: NormalisedMessage, format: Exclude<ReadFormat, 'summary'>) {
  if (format === 'full') return m;
  return { ...compactMessage(m), ...(m.replyTo !== undefined ? { replyTo: m.replyTo } : {}) };
}

const isMissing = (root: ThreadRoot): root is MissingRoot => 'missing' in root;
const isOutside = (root: ThreadRoot): root is OutsideRoot => 'outside_window' in root;

/** The newest of oldest-first `messages`, as `newest_time` (and `newest_time_local`); nothing for none. */
function newestTime(messages: readonly NormalisedMessage[]): { newest_time?: string; newest_time_local?: string } {
  const newest = messages.at(-1);
  if (newest === undefined) return {};
  return {
    newest_time: newest.time,
    ...(newest.time_local !== undefined ? { newest_time_local: newest.time_local } : {}),
  };
}

function summariseThread({ root, replies }: Thread) {
  const subject = isMissing(root) ? undefined : root.subject;
  // A root from outside the window is not one of the read's messages, so it does not time the thread.
  const inWindow = !isMissing(root) && !isOutside(root);
  return {
    root_id: root.id,
    ...(isMissing(root) ? { root_missing: true as const } : {}),
    ...(isOutside(root) ? { root_outside_window: true as const } : {}),
    ...(subject !== undefined ? { subject } : {}),
    reply_count: replies.length,
    // Replies are oldest first and newer than their root, so the last reply is the newest.
    ...newestTime(inWindow ? [root, ...replies] : replies),
  };
}

function compactRoot(root: ThreadRoot): MissingRoot | (CompactMessage & { outside_window?: true }) {
  if (isMissing(root)) return root;
  return { ...compactMessage(root), ...(isOutside(root) ? { outside_window: true as const } : {}) };
}

/** The conversation summary a result carries: whole in full, else just what names it. */
export function formatConversation(
  summary: ConversationSummary,
  format: ReadFormat,
): ConversationSummary | Pick<ConversationSummary, 'id' | 'kind' | 'title' | 'team'> {
  if (format === 'full') return summary;
  const { id, kind, title, team } = summary;
  return { id, kind, title, ...(team !== undefined ? { team } : {}) };
}

/**
 * A conversation's oldest-first messages as a result shows them: `messages`
 * (or a channel's `threads`, with `roots` fetched from outside the window) in
 * full or compact, or their summary.
 */
export function formatBody(
  kind: ConversationKind,
  messages: readonly NormalisedMessage[],
  format: ReadFormat,
  roots: OutsideRoots = NO_ROOTS,
) {
  const channel = kind === 'channel';
  if (format === 'summary') {
    return {
      message_count: messages.length,
      ...newestTime(messages),
      ...(channel ? { threads: groupThreads(messages, roots).map(summariseThread) } : {}),
    };
  }
  if (format === 'compact') {
    if (!channel) return { messages: messages.map(compactMessage) };
    return {
      threads: groupThreads(messages, roots).map(({ root, replies }) => ({
        root: compactRoot(root),
        replies: replies.map(compactMessage),
      })),
    };
  }
  return channel ? { threads: groupThreads(messages, roots) } : { messages: [...messages] };
}
