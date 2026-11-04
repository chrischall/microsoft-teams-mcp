import { mkdtempSync, readFileSync, realpathSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import type { ChatServiceClient, RawMessage } from '../../src/chatsvc.js';
import { registerTranscriptTools } from '../../src/tools/transcript.js';
import { fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const MEETING = '19:meeting_ZXhhbXBsZQ@thread.v2';
const OTHER_CHAT = '19:group0001@thread.v2';
const T0 = 1_793_797_437_680;
const MIN = 60_000;
const OBJ_NEW = '0-wus-d11-0123456789abcdef0123456789abcdef';
const OBJ_OLD = '0-cin-d4-fedcba9876543210fedcba9876543210';

function recording(at: number, obj: string, opts: { status?: string; transcript?: boolean } = {}): RawMessage {
  const gw = `https://us-prod.asyncgw.teams.microsoft.com/v1/objects/${obj}/views`;
  return {
    ...rawMsg(at, { conv: MEETING }),
    messagetype: 'RichText/Media_CallRecording',
    composetime: new Date(at).toISOString(),
    content:
      '<URIObject type="Video.2/CallRecording.1">' +
      `<RecordingStatus status="${opts.status ?? 'Success'}" code="200" />` +
      '<Title>Team Standup</Title>' +
      `<Identifiers><Id type="AMSDocumentID" value="${obj}" /></Identifiers>` +
      '<RecordingContent duration="0:12:18.02">' +
      (opts.transcript === false ? '' : `<item type="amsTranscript" uri="${gw}/transcript" />`) +
      '<item type="onedriveForBusinessVideo" uri="https://contoso-my.sharepoint.com/:v:/p/x/Play" />' +
      '</RecordingContent></URIObject>',
  };
}

const VTT_NEW =
  'WEBVTT\n\n00:00:01.000 --> 00:00:02.000\n<v Lee, Ana>Morning all.</v>\n\n' +
  '00:00:03.000 --> 00:00:04.000\n<v Lee, Ana>Two updates.</v>\n\n00:00:05.000 --> 00:00:06.000\n<v Kim, Bo>Thanks.</v>\n';

const MSGS = [
  recording(T0, OBJ_OLD),
  rawMsg(T0 + 1 * MIN, { conv: MEETING, text: 'see you' }),
  recording(T0 + 30 * MIN, OBJ_NEW, { status: 'Started' }),
  recording(T0 + 31 * MIN, OBJ_NEW),
];

function chatsvcWith(messages = MSGS, fetchTranscript = vi.fn().mockResolvedValue(VTT_NEW)): ChatServiceClient {
  return fakeChatsvc([rawConv(MEETING, { topic: 'Team Standup' }), rawConv(OTHER_CHAT, { topic: 'Lunch' })], {
    getMessages: messagesFrom(messages),
    fetchTranscript,
  });
}

async function call(chatsvc: ChatServiceClient, args: Record<string, unknown>) {
  const harness = await createTestHarness((server) => registerTranscriptTools(server, { chatsvc }));
  try {
    return await harness.callTool('teams_get_transcript', args);
  } finally {
    await harness.close();
  }
}

const errorText = (r: { content?: unknown[] }) => (r.content?.[0] as { text?: string })?.text ?? '';

let outDir: string;
beforeEach(() => {
  outDir = realpathSync(mkdtempSync(join(tmpdir(), 'teams-transcript-')));
  vi.stubEnv('TEAMS_OUTPUT_DIR', outDir);
  vi.stubEnv('MCP_BLOB_BASE_URL', '');
  vi.stubEnv('MCP_BLOB_SIGNING_KEY', '');
});
afterEach(() => {
  vi.unstubAllEnvs();
  rmSync(outDir, { recursive: true, force: true });
});

describe('teams_get_transcript', () => {
  it('returns the newest transcript of a meeting chat, by name, as readable text', async () => {
    const chatsvc = chatsvcWith();
    const result = await call(chatsvc, { conversation: 'standup' });
    const out = parseToolResult(result) as Record<string, any>;

    expect(chatsvc.fetchTranscript).toHaveBeenCalledWith(OBJ_NEW);
    expect(out.conversation).toBe(MEETING);
    expect(out.title).toBe('Team Standup');
    expect(out.recording).toMatchObject({
      recording_id: String(T0 + 31 * MIN),
      time: new Date(T0 + 31 * MIN).toISOString(),
      duration: '0:12:18',
      play_url: 'https://contoso-my.sharepoint.com/:v:/p/x/Play',
    });
    expect(out.recording.object_id).toBeUndefined();
    expect(out.format).toBe('text');
    expect(out.speakers).toEqual(['Lee, Ana', 'Kim, Bo']);
    expect(out.cue_count).toBe(3);
    expect(out.transcript).toBe('[00:00:01] Lee, Ana: Morning all. Two updates.\n[00:00:05] Kim, Bo: Thanks.');
    // The others, to pick from by recording_id.
    expect(out.other_recordings).toEqual([
      expect.objectContaining({ recording_id: String(T0), time: new Date(T0).toISOString(), duration: '0:12:18' }),
    ]);
  });

  it('returns the raw WebVTT with format "vtt"', async () => {
    const result = await call(chatsvcWith(), { conversation: MEETING, format: 'vtt' });
    const out = parseToolResult(result) as Record<string, any>;
    expect(out.format).toBe('vtt');
    expect(out.transcript).toBe(VTT_NEW);
  });

  it('reads the recording named by recording_id — any of its messages', async () => {
    const chatsvc = chatsvcWith();
    await call(chatsvc, { conversation: MEETING, recording_id: String(T0) });
    expect(chatsvc.fetchTranscript).toHaveBeenCalledWith(OBJ_OLD);
    await call(chatsvc, { conversation: MEETING, recording_id: String(T0 + 30 * MIN) });
    expect(chatsvc.fetchTranscript).toHaveBeenLastCalledWith(OBJ_NEW);
  });

  it('names the recordings it has when recording_id matches none', async () => {
    const result = await call(chatsvcWith(), { conversation: MEETING, recording_id: '123' });
    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/123/);
    expect(errorText(result)).toContain(String(T0 + 31 * MIN));
  });

  it('says plainly when a chat has no transcribed recording', async () => {
    const msgs = [rawMsg(T0, { conv: MEETING }), recording(T0 + MIN, OBJ_NEW, { transcript: false })];
    const result = await call(chatsvcWith(msgs), { conversation: MEETING });
    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/no meeting transcript/i);
  });

  it('passes on a transcript the store no longer holds, without trying another recording', async () => {
    const gone = vi.fn().mockRejectedValue(new McpToolError('This transcript is no longer available from Teams (404).', { status: 404 }));
    const chatsvc = chatsvcWith(MSGS, gone);
    const result = await call(chatsvc, { conversation: MEETING });
    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/no longer available/);
    expect(gone).toHaveBeenCalledTimes(1);
  });

  it('saves the WebVTT and the readable text to files with save, returning where — not the transcript', async () => {
    const result = await call(chatsvcWith(), { conversation: MEETING, save: true });
    const out = parseToolResult(result) as Record<string, any>;
    expect(out.transcript).toBeUndefined();
    expect(out.mode).toBe('disk');
    expect(out.location.startsWith(outDir)).toBe(true);
    const vtt = out.files.find((f: { path: string }) => f.path.endsWith('.vtt'));
    const txt = out.files.find((f: { path: string }) => f.path.endsWith('.txt'));
    expect(readFileSync(vtt.path, 'utf8')).toBe(VTT_NEW);
    expect(readFileSync(txt.path, 'utf8')).toContain('[00:00:05] Kim, Bo: Thanks.');
    expect(vtt.path).toMatch(/Team Standup/);
  });

  it('cuts a very long transcript short inline, and says to save it instead', async () => {
    const cue = (i: number) => `00:00:${String(i % 60).padStart(2, '0')}.000 --> 00:00:${String(i % 60).padStart(2, '0')}.500\n<v S${i}>${'word '.repeat(60)}</v>\n`;
    const long = 'WEBVTT\n\n' + Array.from({ length: 1200 }, (_, i) => cue(i)).join('\n');
    const result = await call(chatsvcWith(MSGS, vi.fn().mockResolvedValue(long)), { conversation: MEETING });
    const out = parseToolResult(result) as Record<string, any>;
    expect(out.truncated).toBe(true);
    expect(out.transcript.length).toBeLessThanOrEqual(200_000);
    expect(out.hint).toMatch(/save/);
  });
});
