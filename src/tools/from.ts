/**
 * Oldest-first paging from a time: `from` in, `next_from` out.
 *
 * A caller ingesting messages passes a UTC time as `from` and gets what
 * arrived at or after it, OLDEST first, at most `limit`; `next_from` is where
 * the page ended, and passing it back as `from` returns the next page. It is
 * always returned — on the last page too, and on an empty one (the position
 * unchanged) — so the same loop that backfills history keeps polling for what
 * is new. The server keeps no state: the position is the token.
 *
 * A position is a time plus the keys already returned AT that time (`seen`),
 * so messages sharing a millisecond — across conversations, say — are neither
 * skipped nor repeated when a page ends among them. The token is base64url
 * JSON `{ t, s? }`; a plain ISO 8601 time is a position with nothing seen.
 *
 * The chat service pages newest first, so a page from `from` is found by
 * reading everything that changed since `from` and taking its oldest: a call
 * costs what has happened since its `from`, not its `limit`. Reads therefore
 * stop on a message's last change, not its arrival (`changedMs`, as
 * `teams_get_changes` learned), and run deeper than other reads
 * ({@link FROM_MAX_RAW}); a conversation busier than that since `from` cannot
 * be paged oldest first and says so.
 */
import { z } from 'zod';
import { McpToolError } from '@chrischall/mcp-utils';
import type { ChatServiceClient, RawMessage } from '../chatsvc.js';
import { changedMs } from '../model.js';

/** Raw messages a `from` read may page through per conversation (100 pages of 200). */
export const FROM_MAX_RAW = 20_000;
export const FROM_MAX_PAGES = 100;

export interface FromPosition {
  /** Epoch ms: everything at or after it, less `seen`, is still to come. */
  t: number;
  /** Keys returned at exactly `t`. */
  seen: readonly string[];
}

export const fromArg = z
  .string()
  .min(1)
  .optional()
  .describe(
    'Page oldest first from here: a UTC time (ISO 8601, e.g. 2026-10-01T00:00:00Z; one with no zone is read ' +
      "as UTC) or a previous result's next_from, unchanged. Returns messages that arrived at or after it, " +
      'oldest first, up to limit, with next_from (pass it back as from for the next page — also once has_more ' +
      'is false, to pick up what arrives later) and has_more.',
  );

/** Shared description text for `from` paging. */
export const FROM_DESCRIPTION =
  'from pages oldest first for ingestion and backfill: messages that arrived at or after a UTC time (or a ' +
  "previous result's next_from), oldest first, at most limit; next_from (always returned, with next_from_time) " +
  'resumes after the last message returned and has_more says whether more are already waiting. Pass an old ' +
  'from and keep passing next_from to backfill; once has_more is false the same next_from picks up what ' +
  'arrives later. Delivery is at-least-once: dedupe on message id. Teams pages newest first, so a from call ' +
  `reads everything since from (up to ${FROM_MAX_RAW} raw messages per conversation, past which it reports ` +
  'an error rather than skip any): the further back from is, the slower the call. ';

function badFrom(from: string): McpToolError {
  return new McpToolError(`"${from}" is not a from position.`, {
    hint: "Pass a UTC time such as 2026-10-01T00:00:00Z, or a previous result's next_from unchanged.",
  });
}

/** The position a `from` names: a {@link encodeFrom} token, or an ISO 8601 time (UTC when it names no zone). */
export function decodeFrom(from: string): FromPosition {
  if (/^[A-Za-z0-9_-]+$/.test(from)) {
    let parsed: unknown;
    try {
      parsed = JSON.parse(Buffer.from(from, 'base64url').toString('utf8'));
    } catch {
      parsed = undefined;
    }
    if (parsed !== undefined) {
      const { t, s } = (parsed ?? {}) as { t?: unknown; s?: unknown };
      const seen = s === undefined ? [] : s;
      if (
        typeof t === 'number' &&
        Number.isFinite(t) &&
        t >= 0 &&
        Array.isArray(seen) &&
        seen.every((k) => typeof k === 'string')
      ) {
        return { t, seen: seen as string[] };
      }
      throw badFrom(from);
    }
  }
  // Only a string that looks like a date: Date.parse is lenient with words. A date-time with no zone
  // would parse as LOCAL time, so it is pinned to UTC.
  if (!/^\d{4}-\d{2}-\d{2}/.test(from)) throw badFrom(from);
  const zoned = /T\d{2}:\d{2}(:\d{2}(\.\d+)?)?$/.test(from) ? `${from}Z` : from;
  const t = Date.parse(zoned);
  if (Number.isNaN(t)) throw badFrom(from);
  return { t, seen: [] };
}

export function encodeFrom({ t, seen }: FromPosition): string {
  return Buffer.from(JSON.stringify(seen.length > 0 ? { t, s: seen } : { t }), 'utf8').toString('base64url');
}

/** One pageable item: its arrival (epoch ms) and a key unique among everything paged together. */
export interface Keyed<T> {
  ms: number;
  key: string;
  item: T;
}

export interface FromPage<T> {
  /** Oldest first. */
  page: T[];
  has_more: boolean;
  /** Items past the position not in `page`. */
  remaining: number;
  next: FromPosition;
}

/**
 * The page at `pos`: the items not yet returned, oldest first (ties by key),
 * the first `limit` of them, and the position after the last. `ceiling`
 * (epoch ms) caps how far the position moves: an item newer than it is
 * returned but will come again, because something that arrived as late as it
 * may not have been read yet.
 */
export function takePage<T>(items: readonly Keyed<T>[], pos: FromPosition, limit: number, ceiling?: number): FromPage<T> {
  const seen = new Set(pos.seen);
  const pending = items
    .filter((i) => i.ms > pos.t || (i.ms === pos.t && !seen.has(i.key)))
    .sort((a, b) => a.ms - b.ms || (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
  const page = pending.slice(0, limit);
  const last = page.at(-1);
  let next: FromPosition = pos;
  if (last !== undefined) {
    const t = Math.max(pos.t, ceiling !== undefined ? Math.min(last.ms, ceiling) : last.ms);
    const atT = page.filter((i) => i.ms === t).map((i) => i.key);
    next = { t, seen: t === pos.t ? [...pos.seen, ...atT] : atT };
  }
  return { page: page.map((i) => i.item), has_more: pending.length > page.length, remaining: pending.length - page.length, next };
}

/** What every `from` result carries about where it is and how to go on. */
export function fromReport(pos: FromPosition, next: FromPosition, hasMore: boolean) {
  return {
    from_time: new Date(pos.t).toISOString(),
    has_more: hasMore,
    next_from: encodeFrom(next),
    next_from_time: new Date(next.t).toISOString(),
  };
}

/**
 * Every message of a conversation that changed at or after `boundary` (epoch
 * ms), newest change first — so every message that ARRIVED at or after it,
 * plus older ones edited or reacted to since. A read that runs out of depth
 * before reaching `boundary` throws: its oldest messages are unknown, and a
 * page taken from what was read would skip them.
 */
export async function readFrom(chatsvc: ChatServiceClient, id: string, boundary: number): Promise<RawMessage[]> {
  let capped = false;
  const fetched = await chatsvc.getMessages(id, {
    stopAt: (m) => changedMs(m) < boundary,
    max: FROM_MAX_RAW,
    maxPages: FROM_MAX_PAGES,
    onCapped: () => (capped = true),
  });
  if (capped) {
    throw new McpToolError(
      `More than ${FROM_MAX_RAW} messages changed in this conversation since ${new Date(boundary).toISOString()}, ` +
        'so its oldest cannot be reached: Teams pages newest first.',
      { hint: 'Pass a later from.' },
    );
  }
  return fetched;
}
