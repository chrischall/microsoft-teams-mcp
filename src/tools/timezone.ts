/**
 * The optional `tz` argument of the message-reading tools. Every message's
 * `time` stays UTC (ISO 8601); a time zone adds `time_local` beside it, the
 * same instant as wall-clock time there with its offset. `TEAMS_TIME_ZONE`
 * supplies the default, so a user can set their zone once.
 */
import { z } from 'zod';
import { McpToolError, readEnvVar } from '@chrischall/mcp-utils';
import { isTimeZone } from '../model.js';

export const TIME_ZONE_ENV = 'TEAMS_TIME_ZONE';

/** The tool-description sentence every `tz`-taking tool carries. */
export const TIME_ZONE_DESCRIPTION =
  'tz (an IANA time zone, e.g. America/New_York; default TEAMS_TIME_ZONE) adds time_local — ' +
  'ISO 8601 with its offset — next to each message\'s UTC time.';

export const tzArg = z
  .string()
  .refine(isTimeZone, { message: 'tz must be an IANA time zone such as America/New_York' })
  .optional()
  .describe(
    'IANA time zone (e.g. America/New_York, Europe/London, UTC) to add each message\'s time_local in, ' +
      `ISO 8601 with its offset; time stays UTC. Defaults to ${TIME_ZONE_ENV} when set, else no time_local.`,
  );

/**
 * The time zone to render `time_local` in: `tz` when given (already validated
 * by {@link tzArg}), else `TEAMS_TIME_ZONE`, else none. A `TEAMS_TIME_ZONE`
 * that is not a time zone is an `McpToolError` naming it, rather than being
 * silently ignored.
 */
export function resolveTimeZone(tz: string | undefined, env: NodeJS.ProcessEnv = process.env): string | undefined {
  if (tz !== undefined) return tz;
  const fromEnv = readEnvVar(TIME_ZONE_ENV, { env });
  if (fromEnv === undefined) return undefined;
  if (!isTimeZone(fromEnv)) {
    throw new McpToolError(`${TIME_ZONE_ENV} "${fromEnv}" is not an IANA time zone.`, {
      hint: `Set ${TIME_ZONE_ENV} to a name such as America/New_York, or unset it.`,
    });
  }
  return fromEnv;
}
