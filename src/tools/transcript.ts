/**
 * `teams_get_transcript` — the transcript of a recorded meeting, read from the
 * meeting's chat.
 *
 * The transcript comes from Teams' own object store (`transcript.ts`), with the
 * chat-service token — not from the organizer's OneDrive copy — so it is read
 * even when the organizer turned transcript download off: if you can read the
 * transcript in Teams' Recap tab, this reads it too. No SharePoint token and
 * no extra bridge scope is involved.
 *
 * The chat is read newest first (up to {@link MAX_SCAN} messages) for its
 * recordings; the newest transcribed one is returned unless `recording_id`
 * names another. With `save` the WebVTT and a readable `.txt` are written
 * through mcp-utils' export sink — a folder on a local install, a signed link
 * to a zip on a hosted one — and only where they went comes back.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { McpToolError, currentCallSignal, toolAnnotations } from '@chrischall/mcp-utils';
import { openExportSink, sanitizeExportFileName } from '@chrischall/mcp-utils/export';
import type { RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import { listRecordings, parseRecording, parseVtt, renderTranscriptText, transcriptSpeakers, type Recording } from '../transcript.js';
import { AmbiguousConversationError, ambiguousResult, resolveListed } from './conversation.js';
import { wrapApiError } from './errors.js';
import { workMeta } from './hints.js';
import { EXPORT_OUTPUT_ENV, EXPORT_SERVER_NAME } from './export.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

/** Messages read looking for recordings (newest first). */
export const MAX_SCAN = 2000;

/** Most transcript characters returned inline; past this, `save` it. */
export const MAX_INLINE_CHARS = 200_000;

/** Other recordings listed beside the one returned. */
const MAX_OTHERS = 20;

type Format = 'text' | 'vtt';

/** A recording as the result shows it: `recording_id`, never the internal object id. */
function shown(r: Recording) {
  const { message_id, object_id: _objectId, has_transcript: _has, duration_seconds: _secs, ...rest } = r;
  return { recording_id: message_id, ...rest };
}

export function registerTranscriptTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_get_transcript',
    {
      title: 'Get a Teams meeting transcript',
      description:
        'Get the transcript of a recorded Teams meeting from its meeting chat (by id or by name, as ' +
        'teams_list_conversations lists it). Works even when the organizer disabled transcript download: it ' +
        "reads Teams' own copy, the one the Recap tab shows. Returns the newest transcribed recording unless " +
        'recording_id names another (other_recordings lists the rest, newest first): its time, duration and ' +
        'play_url, the speakers, and the transcript — format "text" (default: one line per speaker turn, ' +
        '"[hh:mm:ss] Name: words") or "vtt" (the WebVTT as Teams serves it). A transcript longer than ' +
        `${MAX_INLINE_CHARS.toLocaleString('en-US')} characters is cut short (truncated: true). save: true writes ` +
        `<meeting>-<time>.vtt and .txt instead (on a local install into a fresh folder under dest, else ${EXPORT_OUTPUT_ENV}, ` +
        `else ~/Downloads/${EXPORT_SERVER_NAME}; on a hosted install as a zip behind a signed link) and returns where. ` +
        "Teams' copy expires after some weeks: an expired one is an error, not another recording. " +
        'Does not mark anything read.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      _meta: workMeta('lookup', { aliases: ['transcript', 'meeting transcript', 'recording', 'recap', 'vtt', 'what was said'] }),
      // save writes a local file the user asked for; nothing in Teams changes.
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe("The meeting chat's id, or a case-insensitive substring of its title."),
        recording_id: z
          .string()
          .min(1)
          .optional()
          .describe('A recording_id from a previous result (recording or other_recordings). Default: the newest.'),
        format: z.enum(['text', 'vtt']).optional().describe('"text" (default, readable turns) or "vtt" (raw WebVTT).'),
        save: z
          .boolean()
          .optional()
          .describe('Write the transcript to files (.vtt and .txt) and return where, instead of the text. Default false.'),
        dest: z
          .string()
          .min(1)
          .optional()
          .describe(
            `With save, on a local install: the folder to create the output folder in (under ${EXPORT_OUTPUT_ENV} when ` +
              'set, else under ~/Downloads, ~/Documents or ~/Desktop).',
          ),
      },
    },
    async ({ conversation, recording_id, format, save, dest }) => {
      try {
        const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
        const summary = await resolveListed(chatsvc, raw, me, conversation, 'read the transcript of');
        const messages = await chatsvc.getMessages(summary.id, { max: MAX_SCAN });
        const recordings = listRecordings(messages);
        if (recordings.length === 0) {
          throw new McpToolError(`No meeting transcript was found in "${summary.title}".`, {
            hint:
              `Read its newest ${MAX_SCAN} messages. A transcript is posted to a meeting's own chat once a recorded ` +
              'or transcribed meeting ends; pass the meeting chat (teams_list_conversations lists them).',
          });
        }

        let chosen: Recording | undefined = recordings[0];
        if (recording_id !== undefined) {
          // Any message of a recording names it: a chunk posts several as its status moves.
          const objectId = recordingObjectOf(messages, recording_id);
          chosen = recordings.find((r) => r.message_id === recording_id || r.object_id === objectId);
          if (chosen === undefined) {
            throw new McpToolError(`No transcribed recording ${recording_id} in "${summary.title}".`, {
              hint: `Its recordings, newest first: ${recordings
                .slice(0, MAX_OTHERS)
                .map((r) => `${r.message_id} (${r.time})`)
                .join(', ')}.`,
            });
          }
        }

        const vtt = await chatsvc.fetchTranscript(chosen.object_id);
        const cues = parseVtt(vtt);
        const text = renderTranscriptText(cues);
        const others = recordings.filter((r) => r !== chosen).slice(0, MAX_OTHERS).map(shown);
        const base = {
          conversation: summary.id,
          title: summary.title,
          recording: shown(chosen),
          speakers: transcriptSpeakers(cues),
          cue_count: cues.length,
          ...(others.length > 0 ? { other_recordings: others } : {}),
        };

        if (save) {
          const sink = await openExportSink({
            name: EXPORT_SERVER_NAME,
            envVar: EXPORT_OUTPUT_ENV,
            baseName: 'teams-transcript',
            ...(dest !== undefined ? { dest } : {}),
            ...(currentCallSignal() !== undefined ? { signal: currentCallSignal() } : {}),
          });
          const stem = sanitizeExportFileName(`${chosen.title ?? summary.title}-${chosen.time.slice(0, 16).replace(':', '')}`);
          const files = [await sink.addFile(`${stem}.vtt`, vtt, 'text/vtt'), await sink.addFile(`${stem}.txt`, `${text}\n`, 'text/plain')];
          const loc = await sink.finish();
          return untrustedResult({
            ...base,
            mode: loc.mode,
            location: loc.location,
            ...(loc.urlExpiresAt !== undefined ? { url_expires_at: loc.urlExpiresAt } : {}),
            files: files.map((f) => ({ ...(f.path !== undefined ? { path: f.path } : { name: f.name }), bytes: f.bytes })),
          });
        }

        const fmt: Format = format ?? 'text';
        const full = fmt === 'vtt' ? vtt : text;
        const truncated = full.length > MAX_INLINE_CHARS;
        return untrustedResult({
          ...base,
          format: fmt,
          transcript: truncated ? full.slice(0, MAX_INLINE_CHARS) : full,
          ...(truncated
            ? {
                truncated: true,
                hint: `The transcript is ${full.length.toLocaleString('en-US')} characters; pass save: true to get all of it as files.`,
              }
            : {}),
        });
      } catch (err) {
        if (err instanceof AmbiguousConversationError) return untrustedResult(ambiguousResult(err));
        throw wrapApiError(err, 'get the transcript');
      }
    },
  );
}

/** The object id of the recording message `id`, if it is one (any of a chunk's messages names it). */
function recordingObjectOf(messages: readonly RawMessage[], id: string): string | undefined {
  const m = messages.find((x) => x.id === id);
  return m === undefined ? undefined : parseRecording(m)?.object_id;
}
