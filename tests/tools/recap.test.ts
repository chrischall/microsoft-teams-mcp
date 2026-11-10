import { describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import type { ChatServiceClient, RawMessage, RawThread } from '../../src/chatsvc.js';
import type { RecapReader, RecapResource } from '../../src/recap.js';
import { registerRecapTools } from '../../src/tools/recap.js';
import { fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const MEETING = '19:meeting_ZXhhbXBsZQ@thread.v2';
const ONE_OFF = '19:meeting_b25lb2Zm@thread.v2';
const GROUP = '19:group0001@thread.v2';
const ORG = '00000000-0000-0000-0000-0000000000c3';
const TENANT = '00000000-0000-0000-0000-0000000000d4';
const SERIES = '040000008200E00074C5B7101A82E00800000000D0D6A79F81A2DC010000000000000000100000002E6B83DFACC35643A5AE3E04B7C84C36';
const ONE_OFF_UID = '040000008200E00074C5B7101A82E00800000000AAAAAAAAAAAAAAAA0000000000000000100000002E6B83DFACC35643A5AE3E04B7C84C36';

const occ = (date: string) => {
  const [y, m, d] = date.split('-').map(Number);
  return SERIES.slice(0, 32) + [y >> 8, y & 255, m, d].map((b) => b.toString(16).padStart(2, '0')).join('').toUpperCase() + SERIES.slice(40);
};

function thread(id: string, meeting: Record<string, unknown> | undefined): RawThread {
  return { id, properties: meeting ? { meeting: JSON.stringify(meeting) } : {} };
}

const SERIES_MEETING = {
  subject: 'Team Standup',
  iCalUid: SERIES,
  organizerId: ORG,
  tenantId: TENANT,
  meetingType: 'Recurring',
  startTime: '2026-02-25T15:00:00+00:00',
  meetingJoinUrl: 'https://teams.microsoft.com/l/meetup-join/x',
};

/** A call event or recording message on `iso`. */
function callEvent(iso: string, type = 'Event/Call'): RawMessage {
  const at = Date.parse(iso);
  return { ...rawMsg(at, { conv: MEETING, type }), composetime: iso };
}

const RECORDING: RecapResource = {
  id: 'r1',
  type: 'Recording',
  location: 'https://contoso-my.sharepoint.com/:v:/p/x/Play',
  metadata: { startTime: '639287934631201136', endTime: '639287938406201136' },
};
const TRANSCRIPT: RecapResource = { id: 't1', type: 'TranscriptV2', metadata: { startTime: '639287934631201136', duration: '00:06:17.5' } };
const FILE: RecapResource = { id: 'f1', type: 'MeetingChat', location: 'https://contoso-my.sharepoint.com/Plan.docx', metadata: { fileTitle: 'Plan.docx' } };
const AI: RecapResource = { id: 'a1', type: 'AISummary', metadata: {}, added: '2026-10-28T15:20:15.608Z' };

function recapWith(byUid: Record<string, RecapResource[]>): RecapReader & { readCollabObject: ReturnType<typeof vi.fn> } {
  return {
    isConfigured: () => true,
    readCollabObject: vi.fn(async (_m, uid: string) => byUid[uid] ?? []),
  };
}

function chatsvcWith(messages: RawMessage[] = []): ChatServiceClient {
  const threads: Record<string, RawThread> = {
    [MEETING]: thread(MEETING, SERIES_MEETING),
    [ONE_OFF]: thread(ONE_OFF, { ...SERIES_MEETING, subject: 'Kickoff', iCalUid: ONE_OFF_UID, meetingType: 'Scheduled' }),
    [GROUP]: thread(GROUP, undefined),
  };
  return fakeChatsvc(
    [rawConv(MEETING, { topic: 'Team Standup' }), rawConv(ONE_OFF, { topic: 'Kickoff' }), rawConv(GROUP, { topic: 'Lunch' })],
    { getMessages: messagesFrom(messages), getThread: vi.fn(async (id: string) => threads[id]) },
  );
}

async function call(chatsvc: ChatServiceClient, recap: RecapReader, args: Record<string, unknown>) {
  const harness = await createTestHarness((server) => registerRecapTools(server, { chatsvc, recap }));
  try {
    return await harness.callTool('teams_get_meeting_recap', args);
  } finally {
    await harness.close();
  }
}

const errorText = (r: { content?: unknown[] }) => (r.content?.[0] as { text?: string })?.text ?? '';
type Out = Record<string, any>;

describe('teams_get_meeting_recap', () => {
  it("reads a one-off meeting's recap by chat name, from its own meeting id", async () => {
    const recap = recapWith({ [ONE_OFF_UID]: [RECORDING, TRANSCRIPT, FILE] });
    const out = parseToolResult(await call(chatsvcWith(), recap, { conversation: 'kickoff' })) as Out;

    expect(recap.readCollabObject).toHaveBeenCalledTimes(1);
    expect(recap.readCollabObject).toHaveBeenCalledWith(expect.objectContaining({ organizerId: ORG, tenantId: TENANT }), ONE_OFF_UID, ONE_OFF);
    expect(out.conversation).toBe(ONE_OFF);
    expect(out.meeting).toEqual({ subject: 'Kickoff', recurring: false, join_url: 'https://teams.microsoft.com/l/meetup-join/x' });
    expect(out.occurrence).toBeUndefined();
    expect(out.recap.recordings).toEqual([
      expect.objectContaining({ duration: '0:06:17', url: 'https://contoso-my.sharepoint.com/:v:/p/x/Play' }),
    ]);
    expect(out.recap.files).toEqual([expect.objectContaining({ name: 'Plan.docx', shared_in: 'chat' })]);
    expect(out.hint).toMatch(/teams_get_transcript/);
  });

  it('for a recurring meeting, defaults to the newest day the chat shows a call on that has a recap', async () => {
    const msgs = [
      callEvent('2026-10-28T14:30:00.000Z'),
      callEvent('2026-10-28T14:19:00.000Z', 'RichText/Media_CallRecording'),
      callEvent('2026-11-04T14:30:00.000Z'),
      rawMsg(Date.parse('2026-11-05T09:00:00Z'), { conv: MEETING, text: 'hi' }),
    ];
    const recap = recapWith({ [occ('2026-10-28')]: [RECORDING], [SERIES]: [FILE] });
    const out = parseToolResult(await call(chatsvcWith(msgs), recap, { conversation: MEETING })) as Out;

    // Newest first: Nov 4 (empty) was tried before Oct 28; the plain message's day never was.
    const uids = recap.readCollabObject.mock.calls.map((c) => c[1]);
    expect(uids).toEqual(expect.arrayContaining([occ('2026-11-04'), occ('2026-10-28'), SERIES]));
    expect(uids).not.toContain(occ('2026-11-05'));
    expect(out.occurrence).toBe('2026-10-28');
    expect(out.recap.recordings).toHaveLength(1);
    expect(out.series.files).toEqual([expect.objectContaining({ name: 'Plan.docx' })]);
    expect(out.other_occurrences).toEqual(['2026-11-04']);
    expect(out.meeting.recurring).toBe(true);
  });

  it('reads just the occurrence named by date', async () => {
    const recap = recapWith({ [occ('2026-11-04')]: [AI] });
    const out = parseToolResult(await call(chatsvcWith([callEvent('2026-10-28T14:30:00Z')]), recap, { conversation: MEETING, date: '2026-11-04' })) as Out;
    expect(recap.readCollabObject.mock.calls.map((c) => c[1])).toEqual(expect.arrayContaining([occ('2026-11-04')]));
    expect(recap.readCollabObject.mock.calls.map((c) => c[1])).not.toContain(occ('2026-10-28'));
    expect(out.occurrence).toBe('2026-11-04');
    expect(out.recap.ai_summary).toEqual({ exists: true, generated: '2026-10-28T15:20:15.608Z' });
    expect(out.hint).toMatch(/AI/);
  });

  it('says so, rather than erroring, when an occurrence has nothing recorded', async () => {
    const out = parseToolResult(await call(chatsvcWith(), recapWith({}), { conversation: MEETING, date: '2026-11-04' })) as Out;
    expect(out.recap).toEqual({});
    expect(out.hint).toMatch(/[Nn]othing/);
  });

  it('asks for a date when a recurring chat shows no call to start from', async () => {
    const recap = recapWith({});
    const out = parseToolResult(await call(chatsvcWith([]), recap, { conversation: MEETING })) as Out;
    expect(out.occurrence).toBeUndefined();
    expect(out.hint).toMatch(/date/);
  });

  it('refuses a chat that is not a meeting', async () => {
    const result = await call(chatsvcWith(), recapWith({}), { conversation: 'lunch' });
    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/not a meeting/i);
  });

  it('rejects a date that is not YYYY-MM-DD', async () => {
    const result = await call(chatsvcWith(), recapWith({}), { conversation: MEETING, date: '11/04/2026' });
    expect(result.isError).toBe(true);
  });

  it('frames the result as untrusted content', async () => {
    const result = await call(chatsvcWith(), recapWith({ [ONE_OFF_UID]: [FILE] }), { conversation: 'kickoff' });
    expect(JSON.stringify(result.content)).toMatch(/untrusted/i);
  });
});
