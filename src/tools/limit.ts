/**
 * Client-side trimming for every list/read tool. A DOM read or a conversation
 * listing returns hundreds of rows at once; dumping all of them into one
 * result fills the context window with other people's messages the question
 * did not need. Each tool takes an optional `limit` (default
 * {@link DEFAULT_LIMIT}) and keeps the rows that lead its ordering — the
 * NEWEST for anything time-ordered (the END of an oldest-first thread, the
 * START of a newest-first feed), the first for an alphabetical listing.
 */
import { z } from 'zod';
import type { RawMessage } from '../chatsvc.js';

export const DEFAULT_LIMIT = 50;

/**
 * The optional `limit` argument, capped at what the source can return. `kept`
 * says which rows survive, for a listing that is not newest-first.
 */
export function limitArg(max: number, what: string, kept = 'newest first kept') {
  return z
    .number()
    .int()
    .min(1)
    .max(max)
    .optional()
    .describe(`Maximum ${what} to return, ${kept} (default ${DEFAULT_LIMIT}, max ${max}).`);
}

/** Where the newest rows sit in what the DOM read returned. */
export type NewestAt = 'start' | 'end';

export type Trimmed<T> = {
  rows: T[];
  /** Present only when rows were left out. */
  truncated?: { total: number; returned: number; hint: string };
};

/** `which` names the kept rows in the hint ("newest", or "first" for a non-time ordering). */
export function trimRows<T>(
  rows: T[],
  limit: number | undefined,
  newestAt: NewestAt,
  max: number,
  which = 'newest',
): Trimmed<T> {
  const n = limit ?? DEFAULT_LIMIT;
  if (rows.length <= n) return { rows };
  const kept = newestAt === 'end' ? rows.slice(rows.length - n) : rows.slice(0, n);
  return {
    rows: kept,
    truncated: {
      total: rows.length,
      returned: kept.length,
      hint: `Only the ${which} ${kept.length} of ${rows.length} were returned; pass a larger limit (max ${max}) to read more.`,
    },
  };
}

/** The optional `since` argument (ISO 8601 date-time) for ISO-timed rows. */
export const sinceArg = z
  .string()
  .refine((s) => !Number.isNaN(Date.parse(s)), { message: 'since must be an ISO 8601 date-time' })
  .optional()
  .describe(
    'Only return messages at or after this ISO 8601 date-time (e.g. 2026-09-21T09:00:00Z). ' +
      'Messages whose time cannot be read are kept.',
  );

/** Drop rows whose ISO `time` is before `since`; rows with no readable time are kept. */
export function filterSince<T extends { time?: string }>(rows: T[], since: string | undefined): T[] {
  if (since === undefined) return rows;
  const cutoff = Date.parse(since);
  return rows.filter((row) => {
    const t = row.time === undefined ? Number.NaN : Date.parse(row.time);
    return Number.isNaN(t) || t >= cutoff;
  });
}

export interface LimitedStopAt {
  /** A `getMessages` `stopAt`: true at the read horizon, or at the first wanted message past `limit`. */
  stopAt: (m: RawMessage) => boolean;
  /** True once a wanted message past `limit` was seen — there is more than was kept. */
  readonly truncated: boolean;
}

/**
 * The `stopAt` behind `teams_read_conversation` and `teams_get_unread`. It
 * sees each message once, newest first, so it can count: the first wanted
 * message past `limit` is the boundary, and its existence is what says the
 * result was truncated. `horizon` (when given) ends the read at the first
 * message the user had already read — checked first, so a boundary there is
 * never a truncation.
 */
export function makeLimitedStopAt(
  horizon: number | undefined,
  wanted: (m: RawMessage) => boolean,
  limit: number,
): LimitedStopAt {
  let kept = 0;
  let truncated = false;
  return {
    stopAt: (m) => {
      if (horizon !== undefined && Number(m.id) <= horizon) return true;
      if (!wanted(m)) return false;
      if (kept >= limit) return (truncated = true);
      kept++;
      return false;
    },
    get truncated() {
      return truncated;
    },
  };
}
