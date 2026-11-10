import { readFileSync } from 'node:fs';
import { describe, expect, it, vi } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import {
  RecapClient,
  collabObjectUrl,
  mcpsBase,
  occurrenceUid,
  parseCollabObject,
  parseMeetingInfo,
  summarizeRecap,
  ticksToIso,
  type MeetingInfo,
} from '../src/recap.js';

/** A JWT carrying `exp`; opaque filler, never a real token. */
function jwt(secondsFromNow = 3000): string {
  const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'none' })}.${b64({ exp: Math.floor(Date.now() / 1000) + secondsFromNow })}.x`;
}

/** Cache off, so no test touches the developer's real `~/.microsoft-teams-mcp`. */
const ENV = { TEAMS_TOKEN_CACHE: 'false' } as NodeJS.ProcessEnv;

const FIXTURE = JSON.parse(readFileSync(new URL('./fixtures/recap-collab-object.json', import.meta.url), 'utf8')) as unknown;

const THREAD = '19:meeting_ZXhhbXBsZQ@thread.v2';
const ORG = '00000000-0000-0000-0000-0000000000c3';
const TENANT = '00000000-0000-0000-0000-0000000000d4';
const SERIES = '040000008200E00074C5B7101A82E00800000000D0D6A79F81A2DC010000000000000000100000002E6B83DFACC35643A5AE3E04B7C84C36';

const MEETING: MeetingInfo = { iCalUid: SERIES, organizerId: ORG, tenantId: TENANT, recurring: true, subject: 'Team Standup' };

type FetchMock = ReturnType<typeof vi.fn<(input: RequestInfo | URL, init?: RequestInit) => Promise<Response>>>;

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

function client(fetchImpl: FetchMock, capture: () => Promise<string> = async () => jwt(), env = ENV) {
  return new RecapClient({ env, captureToken: capture, fetchImpl: fetchImpl as unknown as typeof fetch });
}

describe('parseMeetingInfo', () => {
  const meetingProp = {
    subject: 'Team Standup',
    startTime: '2026-02-25T15:00:00+00:00',
    iCalUid: SERIES,
    meetingType: 'Recurring',
    organizerId: ORG,
    tenantId: TENANT,
    meetingJoinUrl: 'https://teams.microsoft.com/l/meetup-join/x',
  };

  it("reads the meeting a thread's `meeting` property (a JSON string, as Teams sends it) describes", () => {
    expect(parseMeetingInfo({ id: THREAD, properties: { meeting: JSON.stringify(meetingProp) } })).toEqual({
      subject: 'Team Standup',
      iCalUid: SERIES,
      organizerId: ORG,
      tenantId: TENANT,
      recurring: true,
      seriesStart: '2026-02-25T15:00:00.000Z',
      joinUrl: 'https://teams.microsoft.com/l/meetup-join/x',
    });
  });

  it('treats any meetingType but Recurring as a one-off', () => {
    const m = parseMeetingInfo({ id: THREAD, properties: { meeting: { ...meetingProp, meetingType: 'Scheduled' } } });
    expect(m?.recurring).toBe(false);
  });

  it('is undefined for a thread with no usable meeting (a group chat, or ids that are not ids)', () => {
    expect(parseMeetingInfo({ id: THREAD, properties: {} })).toBeUndefined();
    expect(parseMeetingInfo({ id: THREAD, properties: { meeting: 'not json' } })).toBeUndefined();
    expect(parseMeetingInfo({ id: THREAD, properties: { meeting: JSON.stringify({ ...meetingProp, organizerId: '../x' }) } })).toBeUndefined();
    expect(parseMeetingInfo({ id: THREAD, properties: { meeting: JSON.stringify({ ...meetingProp, iCalUid: 'zz' }) } })).toBeUndefined();
  });
});

describe('occurrenceUid', () => {
  it("writes the occurrence's year, month and day into bytes 16-19 of the series id (verified live)", () => {
    expect(occurrenceUid(SERIES, '2026-11-04')).toBe(SERIES.slice(0, 32) + '07EA0B04' + SERIES.slice(40));
    expect(occurrenceUid(SERIES, '2026-10-28').slice(32, 40)).toBe('07EA0A1C');
  });

  it('rejects a date that is not YYYY-MM-DD or not a real day', () => {
    expect(() => occurrenceUid(SERIES, '2026-13-01')).toThrow();
    expect(() => occurrenceUid(SERIES, 'Nov 4')).toThrow();
  });
});

describe('collabObjectUrl', () => {
  it('builds the V2 read path Teams calls: organizer@tenant / occurrence id / thread', () => {
    expect(collabObjectUrl(mcpsBase(ENV), MEETING, SERIES, THREAD)).toBe(
      `https://teams.cloud.microsoft/api/mcps/us/collab/readcollabobject/V2/${ORG}@${TENANT}/${SERIES}/${THREAD}`,
    );
  });

  it('takes the region from TEAMS_MCPS_REGION', () => {
    expect(mcpsBase({ TEAMS_MCPS_REGION: 'emea' } as NodeJS.ProcessEnv)).toBe('https://teams.cloud.microsoft/api/mcps/emea');
  });

  it('refuses anything that could steer the path: a non-hex id or an odd thread id', () => {
    expect(() => collabObjectUrl(mcpsBase(ENV), MEETING, '../x', THREAD)).toThrow();
    expect(() => collabObjectUrl(mcpsBase(ENV), MEETING, SERIES, '19:x/../y@thread.v2')).toThrow();
  });
});

describe('ticksToIso', () => {
  it('converts .NET ticks (100 ns since 0001-01-01) to UTC ISO', () => {
    expect(ticksToIso('639287934631201136')).toBe('2026-10-28T14:11:03.120Z');
    expect(ticksToIso('nope')).toBeUndefined();
    expect(ticksToIso(undefined)).toBeUndefined();
  });
});

describe('parseCollabObject + summarizeRecap', () => {
  const recap = summarizeRecap(parseCollabObject(FIXTURE));

  it('lists the recording with its start, end, duration and play link', () => {
    expect(recap.recordings).toEqual([
      {
        started: '2026-10-28T14:11:03.120Z',
        ended: '2026-10-28T14:17:20.620Z',
        duration: '0:06:17',
        url: 'https://contoso-my.sharepoint.com/:v:/p/organizer/Play1',
      },
    ]);
  });

  it('lists transcripts, new and legacy, without their storage URLs', () => {
    expect(recap.transcripts).toEqual([
      { started: '2026-10-28T14:11:03.120Z', duration: '0:06:17', url: 'https://contoso-my.sharepoint.com/:v:/p/organizer/Play1' },
      { started: '2026-10-28T14:04:57.728Z', state: 'Started' },
    ]);
  });

  it('gives the meeting notes, attendance reports, and shared files and whiteboards', () => {
    expect(recap.notes).toEqual([{ url: 'https://contoso-my.sharepoint.com/:fl:/g/personal/someone/Notes1', added: '2026-10-28T13:59:46.557Z' }]);
    expect(recap.attendance_reports).toEqual([
      {
        started: '2026-10-28T14:10:53.447Z',
        url: 'https://teams.microsoft.com/api/attendancereport/prod/attendancereport/file?ThreadId=19:meeting_ZXhhbXBsZQ@thread.v2&CallId=00000000-0000-0000-0000-0000000000ca',
      },
    ]);
    expect(recap.files).toEqual([
      {
        name: 'Plan.docx',
        file_type: 'docx',
        app: 'Word',
        shared_in: 'chat',
        url: 'https://contoso-my.sharepoint.com/personal/someone/Documents/Microsoft%20Teams%20Chat%20Files/Plan.docx',
        added: '2026-10-28T14:05:37.169Z',
      },
      {
        name: 'Team Standup',
        file_type: 'whiteboard',
        app: 'Whiteboard',
        shared_in: 'meeting',
        url: 'https://contoso-my.sharepoint.com/:wb:/p/someone/Board1Share',
        added: '2026-10-28T14:08:41.697Z',
      },
    ]);
  });

  it('says an AI summary exists without claiming to read it', () => {
    expect(recap.ai_summary).toEqual({ exists: true, generated: '2026-10-28T15:20:15.608Z' });
  });

  it('keeps an unknown type by name, and never passes on a non-https link', () => {
    expect(recap.other).toEqual([{ type: 'SomethingNew', added: '2026-10-28T14:05:12.000Z' }]);
  });

  it('counts a resource once even when the object lists it twice', () => {
    const resources = (FIXTURE as { resources: unknown[] }).resources;
    expect(summarizeRecap(parseCollabObject({ resources: [...resources, ...resources] })).recordings).toHaveLength(1);
  });

  it('is empty, not an error, for an object with no resources or an odd shape', () => {
    expect(summarizeRecap(parseCollabObject({ changeKey: '', resources: [] }))).toEqual({});
    expect(parseCollabObject('nope')).toEqual([]);
  });
});

describe('RecapClient', () => {
  it('GETs the collab object with the captured bearer and returns its resources', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => json(FIXTURE));
    const resources = await client(fetchImpl).readCollabObject(MEETING, SERIES, THREAD);
    expect(resources).toHaveLength(9);
    const [url, init] = fetchImpl.mock.calls[0];
    expect(String(url)).toBe(collabObjectUrl(mcpsBase(ENV), MEETING, SERIES, THREAD));
    expect((init?.headers as Record<string, string>).authorization).toMatch(/^Bearer /);
    expect(init?.redirect).toBe('error');
  });

  it('re-captures once on a 401, then succeeds', async () => {
    const capture = vi.fn(async () => jwt());
    const fetchImpl: FetchMock = vi.fn().mockResolvedValueOnce(new Response('', { status: 401 })).mockResolvedValueOnce(json(FIXTURE));
    await client(fetchImpl, capture).readCollabObject(MEETING, SERIES, THREAD);
    expect(capture).toHaveBeenCalledTimes(2);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('maps a 403 to a plain refusal that never echoes the body', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => json({ message: 'secret detail' }, 403));
    const err = await client(fetchImpl).readCollabObject(MEETING, SERIES, THREAD).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(McpToolError);
    expect((err as McpToolError).status).toBe(403);
    expect((err as Error).message).not.toContain('secret');
  });

  it('is unconfigured, and says why, with the bridge disabled', async () => {
    const c = client(vi.fn(), undefined, { ...ENV, TEAMS_DISABLE_FETCHPROXY: 'true' } as NodeJS.ProcessEnv);
    expect(c.isConfigured()).toBe(false);
    await expect(c.readCollabObject(MEETING, SERIES, THREAD)).rejects.toThrow(/disabled/);
  });

  it('passes a failed capture through without making a request', async () => {
    const capture = vi.fn(async () => {
      throw new Error('capture timed out');
    });
    const fetchImpl: FetchMock = vi.fn();
    await expect(client(fetchImpl, capture).readCollabObject(MEETING, SERIES, THREAD)).rejects.toThrow(/capture timed out/);
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});
