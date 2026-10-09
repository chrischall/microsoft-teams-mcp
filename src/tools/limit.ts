/**
 * Client-side trimming for every list/read tool. The bridge returns up to the
 * selector's `maxItems` rows (200, or 300 for the teams-and-channels sidebar)
 * in one read; dumping all of them into one result fills the context window
 * with other people's messages the question did not need. Each tool takes an
 * optional `limit` (default {@link DEFAULT_LIMIT}) and keeps the NEWEST rows —
 * which sit at the END of a message thread (rendered oldest-first) and at the
 * START of a sidebar or the Activity feed (rendered newest-first).
 */
import { z } from 'zod';

export const DEFAULT_LIMIT = 50;

/** The optional `limit` argument, capped at what the selector can return. */
export function limitArg(max: number, what: string) {
  return z
    .number()
    .int()
    .min(1)
    .max(max)
    .optional()
    .describe(`Maximum ${what} to return, newest first kept (default ${DEFAULT_LIMIT}, max ${max}).`);
}

/** Where the newest rows sit in what the DOM read returned. */
export type NewestAt = 'start' | 'end';

export type Trimmed<T> = {
  rows: T[];
  /** Present only when rows were left out. */
  truncated?: { total: number; returned: number; hint: string };
};

export function trimRows<T>(rows: T[], limit: number | undefined, newestAt: NewestAt, max: number): Trimmed<T> {
  const n = limit ?? DEFAULT_LIMIT;
  if (rows.length <= n) return { rows };
  const kept = newestAt === 'end' ? rows.slice(rows.length - n) : rows.slice(0, n);
  return {
    rows: kept,
    truncated: {
      total: rows.length,
      returned: kept.length,
      hint: `Only the newest ${kept.length} of ${rows.length} were returned; pass a larger limit (max ${max}) to read more.`,
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
