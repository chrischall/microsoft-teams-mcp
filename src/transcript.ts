/**
 * Meeting transcripts: finding them in a meeting chat, and reading the WebVTT
 * Teams serves.
 *
 * Verified live 2026-11-04 against a real tenant:
 *   - every transcribed meeting chunk posts a `RichText/Media_CallRecording`
 *     message (several, as its status moves on) whose content is a
 *     `<URIObject>` carrying the chunk's AMS object id
 *     (`<Id type="AMSDocumentID" value="0-wus-d11-…">`) and, when it was
 *     transcribed, an `<item type="amsTranscript" uri=".../views/transcript">`;
 *     the separate `RichText/Media_CallTranscript` message holds only ids;
 *   - `GET https://us-prod.asyncgw.teams.microsoft.com/v1/objects/<id>/views/transcript`
 *     with the chat-service bearer returns the whole transcript as `text/vtt`
 *     (`<v Speaker>text</v>` cues) — for objects of any region (`0-wus-…`,
 *     `0-cin-…`), and ALSO for a meeting whose organizer disabled transcript
 *     download, where the OneDrive copy is the one locked;
 *   - an older object (about four weeks) answers 404: the store's copy expired.
 *
 * The bearer is only ever sent to a URL this module BUILDS from a vetted
 * object id (see `isAmsObjectId`), never to a URL taken from message content.
 */
import type { RawMessage } from './chatsvc.js';
import { decodeEntities } from './model.js';

/** One meeting recording chunk, as its newest recording message describes it. */
export interface Recording {
  /** The id of the recording message (pass back as `recording_id`). */
  message_id: string;
  /** When the message was posted (UTC, ISO 8601). */
  time: string;
  title?: string;
  /** `h:mm:ss`. */
  duration?: string;
  duration_seconds?: number;
  /** The recording's processing status (`Success`, `ChunkFinished`, …). */
  status?: string;
  /** The AMS object id the video and transcript are served under. */
  object_id: string;
  has_transcript: boolean;
  /** The recording's OneDrive / Stream link, for a person to open. */
  play_url?: string;
}

/** One WebVTT cue. Times are as the file wrote them (`hh:mm:ss.mmm`). */
export interface Cue {
  start: string;
  end: string;
  speaker?: string;
  text: string;
}

const RECORDING_TYPE = 'RichText/Media_CallRecording';

/**
 * An AMS object id as Teams issues them: `0-<region>-<shard>-<32 hex>`. Only
 * an id of this shape is ever put in a URL, so a message cannot point the
 * bearer at another path or host.
 */
export function isAmsObjectId(id: string): boolean {
  return /^\d-[a-z]{2,8}-[a-z0-9]{1,8}-[0-9a-f]{32}$/.test(id);
}

/** The transcript view of an object on the Teams gateway (see the module note). */
export function transcriptUrl(objectId: string): string {
  if (!isAmsObjectId(objectId)) throw new Error(`not an AMS object id: ${JSON.stringify(objectId)}`);
  return `https://us-prod.asyncgw.teams.microsoft.com/v1/objects/${objectId}/views/transcript`;
}

function attr(tagRe: RegExp, xml: string): string | undefined {
  const m = tagRe.exec(xml);
  return m?.[1] !== undefined ? decodeEntities(m[1]) : undefined;
}

/** `0:22:43.96` → `{ text: '0:22:43', seconds: 1363 }`. */
function parseDuration(raw: string | undefined): { text: string; seconds: number } | undefined {
  const m = raw !== undefined ? /^(\d+):(\d{2}):(\d{2})/.exec(raw) : null;
  if (!m) return undefined;
  const [h, mm, ss] = [Number(m[1]), Number(m[2]), Number(m[3])];
  return { text: `${h}:${m[2]}:${m[3]}`, seconds: h * 3600 + mm * 60 + ss };
}

/** The recording a `RichText/Media_CallRecording` message describes, or undefined for any other message. */
export function parseRecording(m: RawMessage): Recording | undefined {
  if (m.messagetype !== RECORDING_TYPE || typeof m.content !== 'string' || m.id === undefined) return undefined;
  const xml = m.content;
  const objectId = attr(/<Id\s+type="AMSDocumentID"\s+value="([^"]*)"/i, xml);
  if (objectId === undefined || !isAmsObjectId(objectId)) return undefined;
  const title = attr(/<Title>([^<]*)<\/Title>/i, xml);
  const duration = parseDuration(attr(/<RecordingContent\b[^>]*\bduration="([^"]*)"/i, xml));
  const status = attr(/<RecordingStatus\b[^>]*\bstatus="([^"]*)"/i, xml);
  const playUrl =
    attr(/<item\s+type="onedriveForBusinessVideo"\s+uri="(https:[^"]*)"/i, xml) ?? attr(/<a\s+href="(https:[^"]*)"/i, xml);
  return {
    message_id: m.id,
    time: m.composetime ?? m.originalarrivaltime ?? new Date(Number(m.id)).toISOString(),
    ...(title !== undefined && title.trim() !== '' ? { title: title.trim() } : {}),
    ...(duration !== undefined ? { duration: duration.text, duration_seconds: duration.seconds } : {}),
    ...(status !== undefined ? { status } : {}),
    object_id: objectId,
    has_transcript: /<item\s+type="amsTranscript"/i.test(xml),
    ...(playUrl !== undefined ? { play_url: playUrl } : {}),
  };
}

/**
 * The transcribed recordings among `messages`, newest first: one per object,
 * described by its newest message (a chunk posts several as its status moves).
 */
export function listRecordings(messages: readonly RawMessage[]): Recording[] {
  const byObject = new Map<string, Recording>();
  for (const m of messages) {
    const r = parseRecording(m);
    if (r === undefined) continue;
    const seen = byObject.get(r.object_id);
    if (seen === undefined || Number(r.message_id) > Number(seen.message_id)) byObject.set(r.object_id, r);
  }
  return [...byObject.values()].filter((r) => r.has_transcript).sort((a, b) => Number(b.message_id) - Number(a.message_id));
}

const TIMING = /^(\d{1,2}:)?\d{2}:\d{2}\.\d{3}\s+-->\s+((?:\d{1,2}:)?\d{2}:\d{2}\.\d{3})/;

/** Parse a WebVTT document into cues; a `<v Speaker>` voice span names the speaker. */
export function parseVtt(vtt: string): Cue[] {
  const blocks = vtt.replace(/^﻿/, '').replace(/\r\n?/g, '\n').split(/\n{2,}/);
  const cues: Cue[] = [];
  for (const block of blocks) {
    const lines = block.split('\n').filter((l) => l.trim() !== '');
    const at = lines.findIndex((l) => TIMING.test(l.trim()));
    if (at < 0) continue; // the header, a NOTE, a STYLE or REGION block
    const timing = lines[at].trim();
    const start = timing.split(/\s+-->/)[0];
    const end = TIMING.exec(timing)![2];
    const body = lines.slice(at + 1).join('\n');
    const voice = /<v(?:\.[^\s>]*)?\s+([^>]+)>/.exec(body);
    const text = decodeEntities(body.replace(/<[^>]*>/g, ''))
      .replace(/\s+/g, ' ')
      .trim();
    if (text === '') continue;
    cues.push({ start, end, ...(voice ? { speaker: decodeEntities(voice[1].trim()) } : {}), text });
  }
  return cues;
}

/**
 * The transcript as readable text: one line per speaker turn — a speaker's
 * consecutive cues run together under the first cue's time (`[hh:mm:ss]`).
 */
export function renderTranscriptText(cues: readonly Cue[]): string {
  const lines: string[] = [];
  let turn: { at: string; speaker?: string; parts: string[] } | undefined;
  const flush = () => {
    if (turn) lines.push(`[${turn.at}] ${turn.speaker !== undefined ? `${turn.speaker}: ` : ''}${turn.parts.join(' ')}`);
  };
  for (const c of cues) {
    if (turn && c.speaker !== undefined && c.speaker === turn.speaker) {
      turn.parts.push(c.text);
      continue;
    }
    flush();
    turn = { at: clockTime(c.start), speaker: c.speaker, parts: [c.text] };
  }
  flush();
  return lines.join('\n');
}

/** A cue time (`[h]h:mm:ss.mmm` or `mm:ss.mmm`) as `hh:mm:ss`. */
function clockTime(t: string): string {
  const parts = t.replace(/\.\d+$/, '').split(':');
  while (parts.length < 3) parts.unshift('0');
  return parts.map((p) => p.padStart(2, '0')).join(':');
}

/** Each speaker once, in order of first appearance. */
export function transcriptSpeakers(cues: readonly Cue[]): string[] {
  return [...new Set(cues.flatMap((c) => (c.speaker !== undefined ? [c.speaker] : [])))];
}
