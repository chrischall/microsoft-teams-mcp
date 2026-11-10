/**
 * `teams_get_meeting_recap` — what Teams' Recap tab shows for a meeting: its
 * recordings, transcripts, meeting notes, attendance reports and shared files,
 * and whether Teams made AI notes, read from the Meeting Collab Platform
 * Service (`recap.ts`) with its own captured token.
 *
 * A one-off meeting has one recap. A recurring one has one per occurrence:
 * `date` names it; by default the chat's newest messages are read for the days
 * a call or recording happened, and the newest of those (up to
 * {@link MAX_PROBE}) whose recap holds anything is returned, with the series'
 * own resources (files shared across the series) beside it.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { McpToolError, toolAnnotations } from '@chrischall/mcp-utils';
import type { RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import { isOccurrenceDate, occurrenceUid, parseMeetingInfo, summarizeRecap, type RecapSummary } from '../recap.js';
import { AmbiguousConversationError, ambiguousResult, resolveListed } from './conversation.js';
import { wrapApiError } from './errors.js';
import { workMeta } from './hints.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

/** Messages read for the days a recurring meeting met (newest first). */
export const MAX_SCAN = 1000;

/** Occurrences tried, newest first, for one with a recap. */
export const MAX_PROBE = 5;

/** Other occurrence days listed beside the one returned. */
const MAX_OTHERS = 20;

/** Messages that mark a meeting having happened: a call starting or ending, or a recording. */
const CALL_TYPES = new Set(['Event/Call', 'RichText/Media_CallRecording', 'RichText/Media_CallTranscript']);

/** The UTC days of the calls in `messages`, newest first, one each. */
export function callDays(messages: readonly RawMessage[]): string[] {
  const days = new Map<string, number>();
  for (const m of messages) {
    if (!CALL_TYPES.has(m.messagetype ?? '')) continue;
    const ms = Date.parse(m.composetime ?? m.originalarrivaltime ?? '') || Number(m.id);
    if (!Number.isFinite(ms) || ms <= 0) continue;
    const day = new Date(ms).toISOString().slice(0, 10);
    days.set(day, Math.max(days.get(day) ?? 0, ms));
  }
  return [...days.entries()].sort((a, b) => b[1] - a[1]).map(([d]) => d);
}

const isEmpty = (r: RecapSummary) => Object.keys(r).length === 0;

function hintFor(recap: RecapSummary, occurrence: string | undefined, recurring: boolean, searched: boolean): string | undefined {
  const parts: string[] = [];
  if (isEmpty(recap)) {
    parts.push(
      recurring && occurrence === undefined && !searched
        ? 'No call was found in this chat to pick an occurrence from: pass date (YYYY-MM-DD) for the occurrence you want.'
        : `Nothing was recorded or shared for ${occurrence !== undefined ? `the ${occurrence} occurrence` : 'this meeting'}: no recording, transcript, notes or files.`,
    );
  }
  if (recap.transcripts !== undefined) parts.push('teams_get_transcript reads a recording\'s transcript text from this same chat.');
  if (recap.ai_summary !== undefined) {
    parts.push("Teams made AI notes (Intelligent Recap) for this meeting; they are shown only in Teams, to licensed users, and are not read here.");
  }
  return parts.length > 0 ? parts.join(' ') : undefined;
}

export function registerRecapTools(server: McpServer, { chatsvc, recap }: Pick<TeamsDeps, 'chatsvc' | 'recap'>): void {
  server.registerTool(
    'teams_get_meeting_recap',
    {
      title: 'Get a Teams meeting recap',
      description:
        "Get what a Teams meeting's Recap tab shows, from its meeting chat (by id or by name, as teams_list_conversations " +
        'lists it): recordings (start, end, duration, play link), transcripts, the meeting notes (Loop) link, attendance ' +
        'reports, files shared in the meeting chat or on stage (e.g. a whiteboard), and whether Teams made AI notes ' +
        '(their text is not read). A recurring meeting has one recap per occurrence: pass date (YYYY-MM-DD) for one, else ' +
        `the newest day the chat shows a call on that has a recap is returned (up to ${MAX_PROBE} tried), with series ` +
        '(what belongs to the whole series) and other_occurrences (other call days, newest first). For the transcript text ' +
        'use teams_get_transcript. Uses its own token, captured when a Recap tab is opened in Teams. Does not mark anything read.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      _meta: workMeta('lookup', {
        aliases: ['recap', 'meeting recap', 'meeting notes', 'attendance report', 'attendance', 'recording link', 'intelligent recap', 'ai notes'],
      }),
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        conversation: z.string().min(1).describe("The meeting chat's id, or a case-insensitive substring of its title."),
        date: z
          .string()
          .refine(isOccurrenceDate, 'date must be a real day, YYYY-MM-DD')
          .optional()
          .describe('For a recurring meeting: the occurrence, YYYY-MM-DD (the UTC day it met). Default: the newest with a recap.'),
      },
    },
    async ({ conversation, date }) => {
      try {
        const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
        const summary = await resolveListed(chatsvc, raw, me, conversation, 'read the recap of');
        const meeting = parseMeetingInfo(await chatsvc.getThread(summary.id));
        if (meeting === undefined) {
          throw new McpToolError(`"${summary.title}" is not a meeting chat: it has no scheduled meeting to recap.`, {
            hint: "Pass a meeting's own chat (teams_list_conversations lists them with type \"meeting\").",
          });
        }
        const read = async (uid: string) => summarizeRecap(await recap.readCollabObject(meeting, uid, summary.id));
        const base = {
          conversation: summary.id,
          title: summary.title,
          meeting: {
            ...(meeting.subject !== undefined ? { subject: meeting.subject } : {}),
            recurring: meeting.recurring,
            ...(meeting.joinUrl !== undefined ? { join_url: meeting.joinUrl } : {}),
          },
        };

        if (!meeting.recurring) {
          const r = await read(meeting.iCalUid);
          const hint = hintFor(r, undefined, false, true);
          return untrustedResult({ ...base, recap: r, ...(hint !== undefined ? { hint } : {}) });
        }

        const days = date !== undefined ? [date] : callDays(await chatsvc.getMessages(summary.id, { max: MAX_SCAN }));
        const series = await read(meeting.iCalUid);
        let chosen: { day: string; recap: RecapSummary } | undefined;
        for (const day of days.slice(0, MAX_PROBE)) {
          const r = await read(occurrenceUid(meeting.iCalUid, day));
          chosen ??= { day, recap: r };
          if (!isEmpty(r)) {
            chosen = { day, recap: r };
            break;
          }
        }
        const r = chosen?.recap ?? {};
        const others = days.filter((d) => d !== chosen?.day).slice(0, MAX_OTHERS);
        const hint = hintFor(r, chosen?.day, true, days.length > 0);
        return untrustedResult({
          ...base,
          ...(chosen !== undefined ? { occurrence: chosen.day } : {}),
          recap: r,
          ...(!isEmpty(series) ? { series } : {}),
          ...(others.length > 0 ? { other_occurrences: others } : {}),
          ...(hint !== undefined ? { hint } : {}),
        });
      } catch (err) {
        if (err instanceof AmbiguousConversationError) return untrustedResult(ambiguousResult(err));
        throw wrapApiError(err, 'get the meeting recap');
      }
    },
  );
}
