import { describe, expect, it } from 'vitest';
import type { RawMessage } from '../src/chatsvc.js';
import {
  isAmsObjectId,
  listRecordings,
  parseRecording,
  parseVtt,
  renderTranscriptText,
  transcriptSpeakers,
} from '../src/transcript.js';

// The shape of a live `RichText/Media_CallRecording` message (verified
// 2026-11-04), with invented ids, names and hosts.
const AMS = '0-wus-d11-0123456789abcdef0123456789abcdef';
const AMS_2 = '0-cin-d4-fedcba9876543210fedcba9876543210';

function recordingContent(opts: { ams?: string; status?: string; transcript?: boolean; title?: string; duration?: string } = {}): string {
  const ams = opts.ams ?? AMS;
  const gw = `https://us-prod.asyncgw.teams.microsoft.com/v1/objects/${ams}/views`;
  return (
    '<URIObject format_version="1.1" type="Video.2/CallRecording.1" uri="" version="1.0">' +
    `<RecordingStatus status="${opts.status ?? 'Success'}" code="200" shortChunk="False"><amsErrorResult /></RecordingStatus>` +
    `<Title>${opts.title ?? 'Weekly Sync &amp; Review'}</Title>` +
    '<a href="https://contoso-my.sharepoint.com/:v:/p/someone/AbCdEf">Play</a>' +
    '<Identifiers><Id type="callId" value="11111111-2222-3333-4444-555555555555" />' +
    `<Id type="AMSDocumentID" value="${ams}" /></Identifiers>` +
    `<RecordingContent contentTypes="Recording+Transcript" timestamp="2026-11-04T13:03:57Z" duration="${opts.duration ?? '0:22:43.96'}">` +
    `<item type="amsVideo" uri="${gw}/video" />` +
    (opts.transcript === false ? '' : `<item type="amsTranscript" uri="${gw}/transcript" />`) +
    '<item type="onedriveForBusinessVideo" uri="https://contoso-my.sharepoint.com/:v:/p/someone/AbCdEf" />' +
    '</RecordingContent></URIObject>'
  );
}

function recordingMsg(id: number, opts: Parameters<typeof recordingContent>[0] = {}): RawMessage {
  return {
    id: String(id),
    messagetype: 'RichText/Media_CallRecording',
    composetime: new Date(id).toISOString(),
    content: recordingContent(opts),
  };
}

const T = 1_793_797_437_680;

describe('isAmsObjectId', () => {
  it('accepts the object ids Teams issues and nothing that could change the URL', () => {
    expect(isAmsObjectId(AMS)).toBe(true);
    expect(isAmsObjectId(AMS_2)).toBe(true);
    expect(isAmsObjectId('0-wus-d11-0123/../../evil')).toBe(false);
    expect(isAmsObjectId('https://evil.com/x')).toBe(false);
    expect(isAmsObjectId('')).toBe(false);
  });
});

describe('parseRecording', () => {
  it('reads the title, time, duration, object id and play link of a recording with a transcript', () => {
    expect(parseRecording(recordingMsg(T))).toEqual({
      message_id: String(T),
      time: new Date(T).toISOString(),
      title: 'Weekly Sync & Review',
      duration: '0:22:43',
      duration_seconds: 1363,
      status: 'Success',
      object_id: AMS,
      has_transcript: true,
      play_url: 'https://contoso-my.sharepoint.com/:v:/p/someone/AbCdEf',
    });
  });

  it('says so when the recording carries no transcript', () => {
    expect(parseRecording(recordingMsg(T, { transcript: false }))?.has_transcript).toBe(false);
  });

  it('ignores any other message, and a recording with no usable object id', () => {
    expect(parseRecording({ id: '1', messagetype: 'Text', content: recordingContent() })).toBeUndefined();
    expect(parseRecording(recordingMsg(T, { ams: '../evil' }))).toBeUndefined();
  });

  it('takes the object id from the Identifiers, never from a transcript URL on another host', () => {
    const msg = recordingMsg(T);
    msg.content = msg.content!.replace(
      `https://us-prod.asyncgw.teams.microsoft.com/v1/objects/${AMS}/views/transcript`,
      'https://evil.example/v1/objects/x/views/transcript',
    );
    // Still a transcript (the item is present), still the vetted object id.
    expect(parseRecording(msg)).toMatchObject({ object_id: AMS, has_transcript: true });
  });
});

describe('listRecordings', () => {
  it('lists recordings with a transcript newest first, one per object (its newest message)', () => {
    const msgs = [
      recordingMsg(T, { status: 'Started' }),
      recordingMsg(T + 1000),
      recordingMsg(T + 5000, { ams: AMS_2, title: 'Later' }),
      recordingMsg(T + 9000, { ams: '0-eus-d1-00000000000000000000000000000000', transcript: false }),
      { id: String(T + 2000), messagetype: 'Text', content: 'hi' },
    ];
    expect(listRecordings(msgs).map((r) => [r.message_id, r.object_id])).toEqual([
      [String(T + 5000), AMS_2],
      [String(T + 1000), AMS],
    ]);
  });
});

const VTT = `WEBVTT

NOTE produced by Teams

1
00:00:03.377 --> 00:00:10.817
<v Remus, Jim>for data onboarding &amp; customers.</v>

00:00:12.017 --> 00:00:12.337
<v Remus, Jim>Um...</v>

00:00:13.457 --> 00:01:33.857
<v Lee, Ana>Next point
spans two lines.</v>

01:02:03.000 --> 01:02:04.500
No speaker here.
`;

describe('parseVtt', () => {
  it('reads cues with their speaker, skipping the header, notes and cue ids', () => {
    expect(parseVtt(VTT)).toEqual([
      { start: '00:00:03.377', end: '00:00:10.817', speaker: 'Remus, Jim', text: 'for data onboarding & customers.' },
      { start: '00:00:12.017', end: '00:00:12.337', speaker: 'Remus, Jim', text: 'Um...' },
      { start: '00:00:13.457', end: '00:01:33.857', speaker: 'Lee, Ana', text: 'Next point spans two lines.' },
      { start: '01:02:03.000', end: '01:02:04.500', text: 'No speaker here.' },
    ]);
  });

  it('copes with CRLF line endings and a BOM', () => {
    expect(parseVtt('\uFEFF' + VTT.replace(/\n/g, '\r\n'))).toHaveLength(4);
  });

  it('is empty for an empty transcript', () => {
    expect(parseVtt('WEBVTT\n\n')).toEqual([]);
  });
});

describe('renderTranscriptText', () => {
  it('runs one speaker\'s consecutive cues together under the first cue\'s time', () => {
    expect(renderTranscriptText(parseVtt(VTT))).toBe(
      [
        '[00:00:03] Remus, Jim: for data onboarding & customers. Um...',
        '[00:00:13] Lee, Ana: Next point spans two lines.',
        '[01:02:03] No speaker here.',
      ].join('\n'),
    );
  });
});

describe('renderTranscriptText times', () => {
  it('writes a cue time with no hours as hh:mm:ss', () => {
    expect(renderTranscriptText(parseVtt('WEBVTT\n\n03:21.500 --> 03:22.000\n<v A>x</v>\n'))).toBe('[00:03:21] A: x');
  });
});

describe('transcriptSpeakers', () => {
  it('lists each speaker once, in order of first appearance', () => {
    expect(transcriptSpeakers(parseVtt(VTT))).toEqual(['Remus, Jim', 'Lee, Ana']);
  });
});
