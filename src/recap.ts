/**
 * Meeting recaps: the Meeting Collab Platform Service (MCPS) behind Teams'
 * Recap tab.
 *
 *   GET https://teams.cloud.microsoft/api/mcps/<region>/collab/readcollabobject/V2/
 *         <organizerOid>@<tenantId>/<occurrence id>/<meeting thread id>
 *       [Authorization: Bearer <MCPS token> — see `recap-auth.ts`]
 *   → { changeKey, id?, participants: [], resources: [{ id, type, location,
 *        metadata: { … }, subType, addedBy, additionDateTime }], … }
 *
 * Verified live 2026-11-10 against a real tenant (40 meeting chats sampled):
 *   - the organizer, tenant and series `iCalUid` come from the meeting thread's
 *     `meeting` property (`GET /threads/<id>`), as does `meetingType`
 *     (`Recurring` / `Scheduled`);
 *   - the occurrence id is the `iCalUid` (an Outlook GlobalObjectId, hex) with
 *     bytes 16-19 — hex characters 32-39 — set to the occurrence's year (2
 *     bytes), month and day; a one-off meeting, and a recurring series as a
 *     whole, use the `iCalUid` as it is (those bytes zero). The day was the
 *     same in UTC and the organizer's zone in every sample: which one MCPS
 *     means is not established;
 *   - an occurrence with nothing recorded answers 200 with `resources: []`;
 *   - resource types seen: `Recording` (location: the OneDrive/Stream play
 *     link; .NET-tick start/end), `TranscriptV2` (location: a SharePoint
 *     `/media/transcripts/<id>/content` API URL; `metadata.duration`),
 *     `transcript` (an older kind, ISO start), `notes` (the Loop meeting-notes
 *     page), `AttendanceReport`, `MeetingChat` (a file shared in the meeting
 *     chat), `MeetingStage` (shared on stage, e.g. a Whiteboard) and
 *     `AISummary` (the Intelligent Recap AI notes);
 *   - the AI notes themselves are a per-user file read through MCPS
 *     `groundingdata/drives/<d>/items/<i>/recapAISummary`: that answered 403
 *     ("Graph denied access") for anyone but the file's owner, and an
 *     unlicensed user sees no notes in Teams either. So this module reports
 *     that a summary exists, and never reads it;
 *   - the chat-service token is a 401 here: MCPS needs its own token.
 *
 * The bearer only ever goes to a URL this module BUILDS from vetted parts
 * (GUIDs, a hex occurrence id, a `19:…@thread.v2` id) — never to a URL taken
 * from a response.
 */
import { fetchBounded, McpToolError, parseBoolEnv, readEnvVar } from '@chrischall/mcp-utils';
import { createCapturedTokenManager, type TokenManager } from '@chrischall/mcp-utils/session';
import type { RawThread } from './chatsvc.js';
import { describeTransportFailure } from './network-error.js';
import { createRecapTokenCache } from './recap-token-cache.js';

/** The MCPS region this tenant was verified on. */
export const DEFAULT_MCPS_REGION = 'us';

/** Per-request deadline, covering the body read. */
const REQUEST_TIMEOUT_MS = 30_000;

/** A collab object is a list of links; anything near this is not one. */
const MAX_RESPONSE_BYTES = 4 * 1024 * 1024;

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
/** An Outlook GlobalObjectId as hex: the 16-byte class id, the date, then the rest. */
const GOID = /^[0-9a-f]{40,256}$/i;
const THREAD_ID = /^19:[A-Za-z0-9_-]+@thread\.(?:v2|tacv2|skype)$/;
const DATE = /^(\d{4})-(\d{2})-(\d{2})$/;

/** .NET ticks at the Unix epoch. */
const EPOCH_TICKS = 621_355_968_000_000_000n;

/** Capture lazily, so boot never loads the bridge. */
async function captureRecapTokenLazily(): Promise<string> {
  const mod = await import('./recap-auth.js');
  return mod.captureRecapTokenViaFetchproxy();
}

/** What a meeting chat's thread says about its meeting. */
export interface MeetingInfo {
  subject?: string;
  /** The series (or one-off meeting) GlobalObjectId, hex. */
  iCalUid: string;
  organizerId: string;
  tenantId: string;
  recurring: boolean;
  /** The first occurrence's start (UTC ISO). */
  seriesStart?: string;
  joinUrl?: string;
}

type Obj = Record<string, unknown>;

function obj(v: unknown): Obj | undefined {
  return v !== null && typeof v === 'object' && !Array.isArray(v) ? (v as Obj) : undefined;
}

function str(v: unknown): string | undefined {
  return typeof v === 'string' && v.trim() !== '' ? v.trim() : undefined;
}

function isoOf(v: unknown): string | undefined {
  const s = str(v);
  if (s === undefined) return undefined;
  const ms = Date.parse(s);
  return Number.isFinite(ms) ? new Date(ms).toISOString() : undefined;
}

function httpsUrl(v: unknown): string | undefined {
  const s = str(v);
  return s !== undefined && /^https:\/\//i.test(s) ? s : undefined;
}

/**
 * The meeting a thread's `meeting` property describes (a JSON string, as
 * Teams sends it, or an object), or `undefined` when there is none or its ids
 * are not ids.
 */
export function parseMeetingInfo(thread: RawThread): MeetingInfo | undefined {
  const raw = thread.properties?.meeting;
  let m: Obj | undefined;
  if (typeof raw === 'string') {
    try {
      m = obj(JSON.parse(raw));
    } catch {
      return undefined;
    }
  } else {
    m = obj(raw);
  }
  const iCalUid = str(m?.iCalUid);
  const organizerId = str(m?.organizerId);
  const tenantId = str(m?.tenantId);
  if (m === undefined || iCalUid === undefined || organizerId === undefined || tenantId === undefined) return undefined;
  if (!GOID.test(iCalUid) || !GUID.test(organizerId) || !GUID.test(tenantId)) return undefined;
  const subject = str(m.subject);
  const seriesStart = isoOf(m.startTime);
  const joinUrl = httpsUrl(m.meetingJoinUrl);
  return {
    ...(subject !== undefined ? { subject } : {}),
    iCalUid,
    organizerId,
    tenantId,
    recurring: m.meetingType === 'Recurring',
    ...(seriesStart !== undefined ? { seriesStart } : {}),
    ...(joinUrl !== undefined ? { joinUrl } : {}),
  };
}

/** Whether `date` is a real `YYYY-MM-DD` day. */
export function isOccurrenceDate(date: string): boolean {
  const m = DATE.exec(date);
  if (!m) return false;
  const d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  return d.toISOString().slice(0, 10) === date;
}

/** The id of the occurrence of series `iCalUid` on `date` (`YYYY-MM-DD`) — see the module note. */
export function occurrenceUid(iCalUid: string, date: string): string {
  if (!GOID.test(iCalUid)) throw new Error(`not a meeting id: ${JSON.stringify(iCalUid)}`);
  if (!isOccurrenceDate(date)) throw new Error(`not a YYYY-MM-DD date: ${JSON.stringify(date)}`);
  const [y, mo, d] = date.split('-').map(Number);
  const hex = [y >> 8, y & 0xff, mo, d].map((b) => b.toString(16).padStart(2, '0')).join('').toUpperCase();
  return iCalUid.slice(0, 32) + hex + iCalUid.slice(40);
}

/** The MCPS base for `TEAMS_MCPS_REGION` (default `us`). */
export function mcpsBase(env: NodeJS.ProcessEnv = process.env): string {
  const region = readEnvVar('TEAMS_MCPS_REGION', { env }) ?? DEFAULT_MCPS_REGION;
  return `https://teams.cloud.microsoft/api/mcps/${encodeURIComponent(region)}`;
}

/** The collab-object read URL, built only from vetted parts. */
export function collabObjectUrl(base: string, meeting: MeetingInfo, uid: string, threadId: string): string {
  if (!GUID.test(meeting.organizerId) || !GUID.test(meeting.tenantId)) throw new Error('not a meeting organizer');
  if (!GOID.test(uid)) throw new Error(`not a meeting id: ${JSON.stringify(uid)}`);
  if (!THREAD_ID.test(threadId)) throw new Error(`not a meeting thread id: ${JSON.stringify(threadId)}`);
  return `${base}/collab/readcollabobject/V2/${meeting.organizerId}@${meeting.tenantId}/${uid}/${threadId}`;
}

/** .NET ticks (100 ns since 0001-01-01, as a decimal string) as UTC ISO. */
export function ticksToIso(ticks: unknown): string | undefined {
  const s = str(ticks);
  if (s === undefined || !/^\d{15,20}$/.test(s)) return undefined;
  const ms = Number((BigInt(s) - EPOCH_TICKS) / 10_000n);
  return Number.isFinite(ms) && ms > 0 ? new Date(ms).toISOString() : undefined;
}

function hms(totalSeconds: number): string {
  const s = Math.floor(totalSeconds);
  return `${Math.floor(s / 3600)}:${String(Math.floor((s % 3600) / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`;
}

/** `00:06:17.5000000` → `0:06:17`. */
function timespan(v: unknown): string | undefined {
  const m = /^(\d+):(\d{2}):(\d{2})/.exec(str(v) ?? '');
  return m ? hms(Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3])) : undefined;
}

/** One resource of a collab object, normalised. */
export interface RecapResource {
  id: string;
  type: string;
  location?: string;
  metadata: Record<string, unknown>;
  /** When it was added (UTC ISO). */
  added?: string;
}

/** The resources of a collab-object response, one per id. Never throws on shape. */
export function parseCollabObject(body: unknown): RecapResource[] {
  const list = obj(body)?.resources;
  if (!Array.isArray(list)) return [];
  const byId = new Map<string, RecapResource>();
  for (const raw of list) {
    const r = obj(raw);
    const id = str(r?.id);
    const type = str(r?.type);
    if (r === undefined || id === undefined || type === undefined || byId.has(id)) continue;
    const location = str(r.location);
    const added = isoOf(r.additionDateTime);
    byId.set(id, {
      id,
      type,
      ...(location !== undefined ? { location } : {}),
      metadata: obj(r.metadata) ?? {},
      ...(added !== undefined ? { added } : {}),
    });
  }
  return [...byId.values()];
}

/** A recap as the tool returns it: only the sections that have something. */
export interface RecapSummary {
  recordings?: { started?: string; ended?: string; duration?: string; url?: string }[];
  transcripts?: { started?: string; duration?: string; state?: string; url?: string }[];
  notes?: { url?: string; added?: string }[];
  attendance_reports?: { started?: string; url?: string }[];
  files?: { name?: string; file_type?: string; app?: string; shared_in: 'chat' | 'meeting'; url?: string; added?: string }[];
  ai_summary?: { exists: true; generated?: string };
  other?: { type: string; url?: string; added?: string }[];
}

/** Drop the undefined fields of a flat record. */
function compact<T extends Record<string, unknown>>(o: T): T {
  return Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined)) as T;
}

/** Group a collab object's resources into the recap's sections. Links are passed on only when https. */
export function summarizeRecap(resources: readonly RecapResource[]): RecapSummary {
  const out: Required<Omit<RecapSummary, 'ai_summary'>> & Pick<RecapSummary, 'ai_summary'> = {
    recordings: [],
    transcripts: [],
    notes: [],
    attendance_reports: [],
    files: [],
    other: [],
  };
  for (const r of resources) {
    const md = r.metadata;
    switch (r.type) {
      case 'Recording': {
        const started = ticksToIso(md.startTime);
        const ended = ticksToIso(md.endTime);
        const duration = started && ended ? hms((Date.parse(ended) - Date.parse(started)) / 1000) : undefined;
        out.recordings.push(compact({ started, ended, duration, url: httpsUrl(r.location) }));
        break;
      }
      case 'TranscriptV2':
        // Its location is a SharePoint API URL, not something a person opens; webUrl is the recording.
        out.transcripts.push(compact({ started: ticksToIso(md.startTime), duration: timespan(md.duration), url: httpsUrl(md.webUrl) }));
        break;
      case 'transcript':
        out.transcripts.push(compact({ started: isoOf(md.startTime), state: str(md.state) }));
        break;
      case 'notes':
        out.notes.push(compact({ url: httpsUrl(r.location), added: r.added }));
        break;
      case 'AttendanceReport':
        out.attendance_reports.push(compact({ started: ticksToIso(md.startTime), url: httpsUrl(md.webUrl) ?? httpsUrl(r.location) }));
        break;
      case 'MeetingChat':
      case 'MeetingStage':
        out.files.push(
          compact({
            name: str(md.fileTitle),
            file_type: str(md.fileType),
            app: str(md.appTitle),
            shared_in: r.type === 'MeetingChat' ? ('chat' as const) : ('meeting' as const),
            url: httpsUrl(md.shareUrl) ?? httpsUrl(r.location),
            added: r.added,
          }),
        );
        break;
      case 'AISummary':
        out.ai_summary = compact({ exists: true as const, generated: r.added });
        break;
      default:
        out.other.push(compact({ type: r.type, url: httpsUrl(r.location), added: r.added }));
    }
  }
  return Object.fromEntries(
    Object.entries(out).filter(([, v]) => v !== undefined && (!Array.isArray(v) || v.length > 0)),
  ) as RecapSummary;
}

/** What `teams_get_meeting_recap` needs of a recap client — tests pass a fake. */
export interface RecapReader {
  isConfigured(): boolean;
  readCollabObject(meeting: MeetingInfo, uid: string, threadId: string): Promise<RecapResource[]>;
}

export interface RecapClientOptions {
  env?: NodeJS.ProcessEnv;
  /** Injected for tests so no suite ever touches the real bridge. */
  captureToken?: () => Promise<string>;
  fetchImpl?: typeof fetch;
}

export class RecapClient implements RecapReader {
  readonly #fetchImpl: typeof fetch;
  readonly #base: string;
  readonly #configError: McpToolError | null = null;
  readonly #manager: TokenManager;

  constructor(opts: RecapClientOptions = {}) {
    const env = opts.env ?? process.env;
    // A receiver-safe wrapper, never the bare global (see `sharepoint.ts`).
    this.#fetchImpl = opts.fetchImpl ?? ((input, init) => fetch(input, init));
    this.#base = mcpsBase(env);
    this.#manager = createCapturedTokenManager({
      capture: opts.captureToken ?? captureRecapTokenLazily,
      persistence: createRecapTokenCache(env),
    });
    if (parseBoolEnv('TEAMS_DISABLE_FETCHPROXY', { env })) {
      this.#configError = new McpToolError('Meeting recaps are disabled (no token source).', {
        kind: 'no_credential',
        hint: 'Unset TEAMS_DISABLE_FETCHPROXY so a recap token can be captured from a signed-in Teams tab.',
      });
    }
  }

  isConfigured(): boolean {
    return this.#configError === null;
  }

  async readCollabObject(meeting: MeetingInfo, uid: string, threadId: string): Promise<RecapResource[]> {
    if (this.#configError) throw this.#configError;
    const url = collabObjectUrl(this.#base, meeting, uid, threadId);
    let text = '';
    const res = await this.#manager.withAuth(async (token) => {
      let r: Awaited<ReturnType<typeof fetchBounded<'text'>>>;
      try {
        r = await fetchBounded(
          url,
          { method: 'GET', headers: { authorization: `Bearer ${token}`, accept: 'application/json' }, redirect: 'error' },
          { read: 'text', maxBytes: MAX_RESPONSE_BYTES, service: 'Teams meeting recap', timeoutMs: REQUEST_TIMEOUT_MS, fetchImpl: this.#fetchImpl },
        );
      } catch (e) {
        throw describeTransportFailure(e, url, 'Teams meeting recap') ?? e;
      }
      text = r.body;
      return r.res;
    });
    checkStatus(res.status);
    try {
      return parseCollabObject(JSON.parse(text));
    } catch {
      // Never echo the body: it holds links to other people's files.
      throw new McpToolError('The Teams meeting-recap service returned a response that is not JSON.', { status: res.status, kind: 'http' });
    }
  }
}

/** Map a non-2xx to an `McpToolError` carrying its status. The body is never included. */
function checkStatus(status: number): void {
  if (status >= 200 && status < 300) return;
  if (status === 401) {
    throw new McpToolError('The Teams meeting-recap service rejected the recap token (401) even after a fresh capture.', {
      status,
      kind: 'credential_rejected',
      hint: "Open a meeting chat's Recap tab in a signed-in Teams tab to mint a fresh recap token, then retry.",
    });
  }
  if (status === 403) {
    throw new McpToolError("You are not allowed to read this meeting's recap (403).", { status, kind: 'http' });
  }
  if (status === 429) {
    throw new McpToolError('The Teams meeting-recap service is rate-limiting requests (429).', {
      status,
      kind: 'http',
      hint: 'Wait a minute before retrying.',
    });
  }
  throw new McpToolError(`The Teams meeting-recap service returned HTTP ${status}.`, {
    status,
    kind: 'http',
    ...(status >= 500 ? { hint: 'A Teams-side failure; retry shortly.' } : {}),
  });
}
