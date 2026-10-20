/**
 * `teams_export` — bulk-export chats and channels to files, so thousands of
 * messages can be kept, searched or handed on without any of them entering
 * the model's context.
 *
 * The tool pages the chat service server-side and writes one file per
 * conversation (`<title>-<short id>.jsonl|md`, `export-format.ts`) through
 * mcp-utils' export sink: on a local (stdio) install into a fresh directory
 * under `dest` → `TEAMS_OUTPUT_DIR` → `~/Downloads/microsoft-teams-mcp`, with
 * ABSOLUTE paths in the result; on mcp-host (which sets `MCP_BLOB_BASE_URL`
 * and `MCP_BLOB_SIGNING_KEY`) as one zip in the hosted blob store behind a
 * signed link. The result is only a manifest — where the files are, counts,
 * the date span, a resume cursor — never a message.
 *
 * A call does bounded work: it stops at `max_items` messages or
 * `budget_seconds`, whichever comes first, and returns `has_more` with a
 * `cursor`; passing the cursor back continues the SAME export (the same
 * directory on disk, the next `part-NNNN.zip` in the blob store). The cursor
 * carries the original arguments, the conversations still to do (or, when
 * that list is too long for a cursor, the id to go on from — conversations
 * are exported in id order), and where in the current conversation it
 * stopped, as a `from.ts` position. A rejected credential, persistent rate
 * limiting or a Teams-side failure stops the run the same way, keeping what
 * was written; any other failure to read one conversation is recorded in
 * `errors` and the export moves on.
 *
 * Reading a conversation is `readFrom`'s rule: the chat service pages newest
 * CHANGE first, so the read stops at the first message whose last change is
 * before `since` (or the resume position), and messages are then taken by
 * arrival, oldest first (`takePage`).
 */
import { createHash } from 'node:crypto';
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import {
  currentCallSignal,
  McpToolError,
  reportProgress,
  throwIfCancelled,
  toolAnnotations,
} from '@chrischall/mcp-utils';
import {
  decodeExportCursor,
  encodeExportCursor,
  exportPartName,
  openExportSink,
  sanitizeExportFileName,
} from '@chrischall/mcp-utils/export';
import type { ChatServiceClient, RawConversation, RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import {
  changedMs,
  conversationSummary,
  isUserMessage,
  lastChangeMs,
  normaliseMessage,
  summariseAll,
  type ConversationKind,
  type ConversationSummary,
  type NormalisedMessage,
} from '../model.js';
import { AmbiguousConversationError, ambiguousResult, arrivalMs, listedResolver } from './conversation.js';
import { wrapApiError } from './errors.js';
import {
  EXPORT_EXT,
  EXPORT_FORMATS,
  EXPORT_UNTRUSTED_NOTE,
  renderConversation,
  type ExportConversation,
  type ExportFormat,
  type ExportHeader,
} from './export-format.js';
import { takePage, type FromPosition } from './from.js';
import { conversationName, enrichTitles } from './titles.js';
import { resolveTimeZone, tzArg } from './timezone.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

export const EXPORT_SERVER_NAME = 'microsoft-teams-mcp';
export const EXPORT_OUTPUT_ENV = 'TEAMS_OUTPUT_DIR';
export const DEFAULT_MAX_ITEMS = 5000;
export const MAX_ITEMS = 50_000;
export const DEFAULT_BUDGET_SECONDS = 90;
export const MAX_BUDGET_SECONDS = 600;
/** Conversations one call may name. */
export const MAX_NAMED_CONVERSATIONS = 200;
/** Raw messages read per conversation (250 pages of 200) before the oldest are reported as not exported. */
export const EXPORT_MAX_RAW = 50_000;
export const EXPORT_MAX_PAGES = 250;
/** Waits before each retry of a rate-limited read; after the last, the export stops with a cursor. */
export const RATE_LIMIT_RETRY_MS = [2_000, 4_000, 8_000] as const;
/** Files listed in the result; the manifest file lists them all. */
const RESULT_FILES_CAP = 50;
/** Errors listed in the result; the manifest file lists them all. */
const RESULT_ERRORS_CAP = 20;

export const EXPORT_KINDS = ['chat', 'channel', 'meeting'] as const;
type ExportKind = (typeof EXPORT_KINDS)[number];
const KINDS: Record<ExportKind, readonly ConversationKind[]> = {
  chat: ['oneOnOne', 'group', 'self'],
  channel: ['channel'],
  meeting: ['meeting'],
};

/** Test seams: the wait between rate-limit retries, the clock the budget runs on, the blob PUT. */
export interface ExportToolOptions {
  sleep?: (ms: number) => Promise<void>;
  now?: () => number;
  fetchImpl?: typeof fetch;
}

/** An ISO 8601 date or date-time; one with a time but no zone is read as UTC. */
function parseIsoTime(s: string): number {
  const zoned = /T\d{2}:\d{2}(:\d{2}(\.\d+)?)?$/.test(s) ? `${s}Z` : s;
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? Date.parse(zoned) : Number.NaN;
}

const isoTimeArg = (what: string) =>
  z
    .string()
    .refine((s) => !Number.isNaN(parseIsoTime(s)), { message: `${what} must be an ISO 8601 date or date-time` })
    .optional();

const conversationsArg = z.union([
  z.string().min(1),
  z.array(z.string().min(1)).min(1).max(MAX_NAMED_CONVERSATIONS),
]);
const kindsArg = z.array(z.enum(EXPORT_KINDS)).min(1);

/** The arguments an export keeps across its parts (in the cursor). */
const storedArgs = z
  .object({
    conversations: conversationsArg.optional(),
    kinds: kindsArg.optional(),
    since: isoTimeArg('since'),
    until: isoTimeArg('until'),
    format: z.enum(EXPORT_FORMATS),
    tz: tzArg,
  })
  .strict();
type StoredArgs = z.infer<typeof storedArgs>;

/** Where an export stopped, as its cursor carries it. */
interface ResumeState {
  exportId: string;
  part: number;
  dir?: string;
  args: StoredArgs;
  /** The conversations still to do, the first being the one it stopped in. */
  ids?: string[];
  /** Instead of `ids` when they did not fit: go on from this id, in id order. */
  from_id?: string;
  /** The position inside the first of them, when it stopped part-way. */
  pos?: { id: string; t: number; s?: string[] };
}

const resumeSchema = z.object({
  exportId: z.string(),
  part: z.number().int().min(2),
  dir: z.string().optional(),
  args: storedArgs,
  ids: z.array(z.string().min(1)).min(1).optional(),
  from_id: z.string().min(1).optional(),
  pos: z
    .object({ id: z.string().min(1), t: z.number().finite().min(0), s: z.array(z.string()).optional() })
    .optional(),
});

function badCursor(): McpToolError {
  return new McpToolError('The export cursor is not valid.', {
    hint: 'Pass the cursor exactly as the previous call returned it, or start a new export without one.',
  });
}

function decodeResume(cursor: string): ResumeState {
  const { v: _v, ...state } = decodeExportCursor(cursor);
  const parsed = resumeSchema.safeParse(state);
  if (!parsed.success || (parsed.data.ids === undefined) === (parsed.data.from_id === undefined)) throw badCursor();
  return parsed.data;
}

/** A short, stable, filename-safe tag for a conversation id. */
export function shortId(id: string): string {
  return createHash('sha256').update(id).digest('hex').slice(0, 8);
}

/** Whether a failure to read one conversation should stop the export (with a cursor) rather than skip it. */
function stopsExport(err: unknown): boolean {
  if (!(err instanceof McpToolError)) return true; // a token capture failure, or something unexpected
  if (err.status === 401 || err.status === 429 || (err.status !== undefined && err.status >= 500)) return true;
  return ['no_credential', 'credential_rejected', 'session_expired', 'transport', 'timeout'].includes(err.kind ?? '');
}

const isRateLimited = (err: unknown): boolean => err instanceof McpToolError && err.status === 429;

async function withRateLimitRetry<T>(fn: () => Promise<T>, sleep: (ms: number) => Promise<void>): Promise<T> {
  for (let attempt = 0; ; attempt++) {
    try {
      return await fn();
    } catch (err) {
      if (!isRateLimited(err) || attempt >= RATE_LIMIT_RETRY_MS.length) throw err;
      await sleep(RATE_LIMIT_RETRY_MS[attempt]);
      throwIfCancelled();
    }
  }
}

const realSleep = (ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms));

/** The conversations an export covers, in id order: the named ones, or every one listed, narrowed by kind and `since`. */
async function selectConversations(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  args: StoredArgs,
  since: number | undefined,
): Promise<string[]> {
  const kinds = new Set((args.kinds ?? EXPORT_KINDS).flatMap((k) => KINDS[k]));
  let summaries: ConversationSummary[];
  if (args.conversations !== undefined) {
    const resolve = listedResolver(chatsvc, raw, me, 'export');
    const named = Array.isArray(args.conversations) ? args.conversations : [args.conversations];
    summaries = [];
    for (const n of new Set(named)) summaries.push(await resolve(n));
  } else {
    summaries = raw.map((c) => conversationSummary(c, me, raw));
  }
  const byId = new Map(raw.map((c) => [c.id, c]));
  const ids = new Set<string>();
  for (const s of summaries) {
    if (!kinds.has(s.kind)) continue;
    const conv = byId.get(s.id);
    // Nothing changed in it since `since`, so nothing arrived since: no need to read it.
    const changed = conv !== undefined ? lastChangeMs(conv) : undefined;
    if (since !== undefined && changed !== undefined && changed < since) continue;
    ids.add(s.id);
  }
  return [...ids].sort();
}

/** The people other than the user known to be in a chat: its titled participants, then who wrote what is exported. */
function chatParticipants(titled: ConversationSummary, messages: readonly NormalisedMessage[]): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  const add = (name: string) => {
    const key = name.trim().toLowerCase();
    if (key === '' || seen.has(key)) return;
    seen.add(key);
    out.push(name);
  };
  for (const p of titled.participants ?? []) add(p);
  for (const m of [...messages].reverse()) if (!m.fromMe) add(m.sender);
  return out;
}

interface WrittenFile {
  name: string;
  path?: string;
  bytes: number;
  records: number;
  conversation: { id: string; kind: ConversationKind; title: string };
}

interface ExportError {
  conversation: { id: string; title: string };
  error: string;
  hint?: string;
}

interface ExportInput {
  conversations?: string | string[];
  kinds?: ExportKind[];
  since?: string;
  until?: string;
  format?: ExportFormat;
  dest?: string;
  cursor?: string;
  max_items?: number;
  budget_seconds?: number;
  tz?: string;
}

/** Arguments that start an export, which a cursor already carries. */
const STARTING_ARGS = ['conversations', 'kinds', 'since', 'until', 'format', 'dest', 'tz'] as const;

async function runExport(chatsvc: ChatServiceClient, input: ExportInput, opts: ExportToolOptions) {
  const clock = opts.now ?? Date.now;
  const sleep = opts.sleep ?? realSleep;
  const started = clock();

  let resume: ResumeState | undefined;
  if (input.cursor !== undefined) {
    const clash = STARTING_ARGS.find((k) => input[k] !== undefined);
    if (clash !== undefined) {
      throw new McpToolError(`cursor cannot be combined with ${clash}.`, {
        hint: 'A cursor carries the export\'s original arguments: pass only cursor (and optionally max_items / budget_seconds).',
      });
    }
    resume = decodeResume(input.cursor);
  }
  const args: StoredArgs = resume?.args ?? {
    ...(input.conversations !== undefined ? { conversations: input.conversations } : {}),
    ...(input.kinds !== undefined ? { kinds: input.kinds } : {}),
    ...(input.since !== undefined ? { since: input.since } : {}),
    ...(input.until !== undefined ? { until: input.until } : {}),
    format: input.format ?? 'jsonl',
    ...(input.tz !== undefined ? { tz: input.tz } : {}),
  };
  const since = args.since !== undefined ? parseIsoTime(args.since) : undefined;
  const until = args.until !== undefined ? parseIsoTime(args.until) : undefined;
  if (since !== undefined && until !== undefined && since > until) {
    throw new McpToolError('since is after until.', { hint: 'Pass a since at or before until.' });
  }
  const timeZone = resolveTimeZone(args.tz);
  const maxItems = input.max_items ?? DEFAULT_MAX_ITEMS;
  const budgetMs = (input.budget_seconds ?? DEFAULT_BUDGET_SECONDS) * 1000;

  let raw: RawConversation[];
  let me: string;
  try {
    [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
  } catch (err) {
    throw wrapApiError(err, 'list conversations to export');
  }
  let account: string | undefined;
  try {
    account = (await chatsvc.meDisplayName()) || undefined;
  } catch {
    account = undefined;
  }

  let queue: string[];
  if (resume?.ids !== undefined) {
    queue = resume.ids;
  } else {
    let selected: string[];
    try {
      selected = await selectConversations(chatsvc, raw, me, args, since);
    } catch (err) {
      // A fresh export hands back the candidates, having written nothing.
      if (err instanceof AmbiguousConversationError && resume === undefined) return untrustedResult(ambiguousResult(err));
      throw wrapApiError(err, 'resolve the conversations to export');
    }
    queue = resume?.from_id !== undefined ? selected.filter((id) => id >= resume.from_id!) : selected;
  }

  const sink = await openExportSink({
    name: EXPORT_SERVER_NAME,
    envVar: EXPORT_OUTPUT_ENV,
    baseName: 'teams-export',
    ...(input.dest !== undefined ? { dest: input.dest } : {}),
    ...(resume !== undefined ? { exportId: resume.exportId, part: resume.part } : {}),
    ...(resume?.dir !== undefined ? { existingDir: resume.dir } : {}),
    ...(opts.fetchImpl !== undefined ? { fetchImpl: opts.fetchImpl } : {}),
    ...(currentCallSignal() !== undefined ? { signal: currentCallSignal() } : {}),
  });

  const header: ExportHeader = {
    source: 'microsoft-teams',
    ...(account !== undefined ? { account } : {}),
    generated_at: new Date(clock()).toISOString(),
    export_id: sink.exportId,
    part: sink.part,
    args,
  };
  const summaries = new Map(summariseAll(raw, me).map((s) => [s.id, s]));
  const files: WrittenFile[] = [];
  const errors: ExportError[] = [];
  let written = 0;
  let oldest: string | undefined;
  let newest: string | undefined;
  let hasMore = false;
  let pos = resume?.pos;
  let stopPos: ResumeState['pos'];
  const total = queue.length;
  let done = 0;

  while (queue.length > 0) {
    if (written >= maxItems || clock() - started >= budgetMs) {
      hasMore = true;
      break;
    }
    throwIfCancelled();
    const id = queue[0];
    const resumed = pos !== undefined && pos.id === id ? pos : undefined;
    pos = undefined;
    const summary = summaries.get(id) ?? conversationSummary({ id }, me);
    const [titled] = await enrichTitles(chatsvc, raw, me, [summary]);
    const start: FromPosition = resumed !== undefined ? { t: resumed.t, seen: resumed.s ?? [] } : { t: since ?? 0, seen: [] };

    let fetched: RawMessage[];
    let capped = false;
    try {
      fetched = await withRateLimitRetry(
        () =>
          chatsvc.getMessages(id, {
            stopAt: (m) => {
              throwIfCancelled();
              return start.t > 0 && changedMs(m) < start.t;
            },
            max: EXPORT_MAX_RAW,
            maxPages: EXPORT_MAX_PAGES,
            onCapped: () => (capped = true),
          }),
        sleep,
      );
    } catch (err) {
      if (currentCallSignal()?.aborted) throw err;
      const e = wrapApiError(err, 'read this conversation');
      errors.push({
        conversation: { id, title: titled.title },
        error: e.message,
        ...(e.hint !== undefined ? { hint: e.hint } : {}),
      });
      if (stopsExport(err)) {
        hasMore = true;
        stopPos = resumed;
        break;
      }
      queue.shift();
      done++;
      continue;
    }

    const wanted = fetched.filter((m) => isUserMessage(m) && (until === undefined || arrivalMs(m) <= until));
    const { page, has_more, next } = takePage(
      wanted.map((m) => ({ ms: arrivalMs(m), key: m.id, item: m })),
      start,
      maxItems - written,
    );
    if (capped) {
      errors.push({
        conversation: { id, title: titled.title },
        error:
          `Only the newest ${EXPORT_MAX_RAW} messages of this conversation could be read; older ones were not ` +
          'exported.',
        hint: 'Export it again with a later since to cover the rest in a separate export.',
      });
    }
    if (page.length > 0) {
      const messages = page.map((m) => normaliseMessage(m, me, { timeZone }));
      const { participants: _named, ...named } = conversationName(titled);
      const conversation: ExportConversation = {
        ...named,
        ...(titled.kind !== 'channel' ? { participants: chatParticipants(titled, messages) } : {}),
      };
      const ext = EXPORT_EXT[args.format];
      const stem = `${sanitizeExportFileName(titled.title, 80)}-${shortId(id)}`;
      // A conversation continued from an earlier part goes in a part-numbered file beside its first.
      const name = resumed !== undefined ? `${stem}.${exportPartName(sink.part, ext)}` : `${stem}.${ext}`;
      const rec = await sink.addFile(name, renderConversation(args.format, header, conversation, messages));
      files.push({
        name,
        ...(rec.path !== undefined ? { path: rec.path } : {}),
        bytes: rec.bytes,
        records: messages.length,
        conversation: { id, kind: titled.kind, title: titled.title },
      });
      written += messages.length;
      const first = messages[0].time;
      const last = messages.at(-1)!.time;
      if (oldest === undefined || first < oldest) oldest = first;
      if (newest === undefined || last > newest) newest = last;
    }
    if (has_more) {
      hasMore = true;
      stopPos = { id, t: next.t, ...(next.seen.length > 0 ? { s: [...next.seen] } : {}) };
      break;
    }
    queue.shift();
    done++;
    await reportProgress(done, total, `${done}/${total} conversations`);
  }

  const counts = { conversations: files.length, messages: written, attachments: 0, attachments_skipped: 0 };
  const span = oldest !== undefined && newest !== undefined ? { oldest, newest } : undefined;
  let cursor: string | undefined;
  if (hasMore && queue.length > 0) {
    const base = {
      exportId: sink.exportId,
      part: sink.part + 1,
      ...(sink.dir !== undefined ? { dir: sink.dir } : {}),
      args,
      ...(stopPos !== undefined ? { pos: stopPos } : {}),
    };
    try {
      cursor = encodeExportCursor({ ...base, ids: queue });
    } catch (err) {
      if (!(err instanceof TypeError)) throw err;
      cursor = encodeExportCursor({ ...base, from_id: queue[0] });
    }
  } else {
    hasMore = false;
  }

  const manifestName = sink.part === 1 ? 'manifest.json' : `manifest.${exportPartName(sink.part, 'json')}`;
  const manifest = await sink.addFile(
    manifestName,
    `${JSON.stringify(
      {
        ...header,
        note: EXPORT_UNTRUSTED_NOTE,
        files: files.map(({ path: _p, ...f }) => f),
        counts,
        ...(span !== undefined ? { span } : {}),
        has_more: hasMore,
        ...(errors.length > 0 ? { errors } : {}),
      },
      null,
      2,
    )}\n`,
  );
  const loc = await sink.finish();

  return untrustedResult({
    mode: loc.mode,
    location: loc.location,
    ...(loc.urlExpiresAt !== undefined ? { url_expires_at: loc.urlExpiresAt } : {}),
    export_id: loc.exportId,
    part: loc.part,
    manifest: manifest.path ?? manifest.name,
    files: files
      .slice(0, RESULT_FILES_CAP)
      .map((f) => ({ ...(f.path !== undefined ? { path: f.path } : { name: f.name }), bytes: f.bytes, records: f.records })),
    ...(files.length > RESULT_FILES_CAP ? { files_truncated: files.length - RESULT_FILES_CAP } : {}),
    counts,
    ...(span !== undefined ? { span } : {}),
    has_more: hasMore,
    ...(cursor !== undefined ? { cursor } : {}),
    ...(errors.length > 0 ? { errors: errors.slice(0, RESULT_ERRORS_CAP) } : {}),
    ...(errors.length > RESULT_ERRORS_CAP ? { errors_truncated: errors.length - RESULT_ERRORS_CAP } : {}),
  });
}

export function registerExportTools(
  server: McpServer,
  { chatsvc }: Pick<TeamsDeps, 'chatsvc'>,
  opts: ExportToolOptions = {},
): void {
  server.registerTool(
    'teams_export',
    {
      title: 'Export Teams chats and channels to files',
      description:
        'Bulk-export the messages of chats and channels to files — for archiving, or to search or analyse ' +
        'thousands of messages with shell tools — without any message text entering this conversation. The ' +
        'result holds no message content: only where the files are (on a local install, absolute paths in a ' +
        `fresh folder under dest, else ${EXPORT_OUTPUT_ENV}, else ~/Downloads/${EXPORT_SERVER_NAME}; on a hosted ` +
        'install, a signed download link to a zip that expires, url_expires_at), each file\'s size and record ' +
        'count, counts, the time span, errors, and has_more / cursor. One file per conversation, named after it ' +
        '(an untitled group chat after its participants), plus a manifest; JSONL by default (a header record ' +
        'with the export\'s details and an untrusted-content note, then one record per message, oldest first, ' +
        'each with its conversation\'s id, kind, title, team and — for chats — participants), or format ' +
        '"markdown" for readable transcripts (a channel threaded under its root posts). Use jq, grep or rg on ' +
        'the files rather than reading them whole. conversations names chats/channels by id or name (omit for ' +
        'every one), kinds narrows to chat / channel / meeting, since and until bound messages by arrival ' +
        '(default: all history). Each call stops at max_items messages ' +
        `(default ${DEFAULT_MAX_ITEMS}, max ${MAX_ITEMS}) or budget_seconds (default ${DEFAULT_BUDGET_SECONDS}, ` +
        `max ${MAX_BUDGET_SECONDS}); when has_more is true pass cursor back (alone, or with max_items / ` +
        'budget_seconds) to continue the same export into the same folder (a hosted install returns a new link ' +
        'per part). A rejected credential or persistent rate limiting stops early with a cursor and keeps what ' +
        'was written; a conversation that cannot be read is listed in errors and skipped. Never marks anything ' +
        'read. A name matching several conversations returns {ambiguous: true, candidates} and writes nothing.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        conversations: conversationsArg
          .optional()
          .describe(
            'Chats/channels to export: an id ("48:notes" for the self-chat) or a case-insensitive substring of a ' +
              `title, channel, team or participant name, or an array of up to ${MAX_NAMED_CONVERSATIONS}. Omit to ` +
              'export every conversation.',
          ),
        kinds: kindsArg
          .optional()
          .describe('Only these kinds: chat (1:1, group chats, the self-chat), channel, meeting. Default all.'),
        since: isoTimeArg('since').describe(
          'Only messages that arrived at or after this ISO 8601 date/time (UTC when it names no zone). Default: all history.',
        ),
        until: isoTimeArg('until').describe('Only messages that arrived at or before this ISO 8601 date/time. Default: now.'),
        format: z
          .enum(EXPORT_FORMATS)
          .optional()
          .describe('jsonl (default: one JSON record per message) or markdown (readable transcripts).'),
        dest: z
          .string()
          .min(1)
          .optional()
          .describe(
            `Local install only: the folder to create the export folder in (under ${EXPORT_OUTPUT_ENV} when set, ` +
              'else under ~/Downloads, ~/Documents or ~/Desktop). Default: ' +
              `${EXPORT_OUTPUT_ENV}, else ~/Downloads/${EXPORT_SERVER_NAME}.`,
          ),
        cursor: z
          .string()
          .min(1)
          .optional()
          .describe('A previous result\'s cursor, unchanged, to continue that export. Pass no other argument but max_items / budget_seconds.'),
        max_items: z
          .number()
          .int()
          .min(1)
          .max(MAX_ITEMS)
          .optional()
          .describe(`Most messages this call writes (default ${DEFAULT_MAX_ITEMS}, max ${MAX_ITEMS}).`),
        budget_seconds: z
          .number()
          .int()
          .min(1)
          .max(MAX_BUDGET_SECONDS)
          .optional()
          .describe(
            `Seconds this call may spend before stopping with a cursor (default ${DEFAULT_BUDGET_SECONDS}, max ${MAX_BUDGET_SECONDS}).`,
          ),
        tz: tzArg,
      },
    },
    async (input) => runExport(chatsvc, input, opts),
  );
}
