import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createHash, createHmac } from 'node:crypto';
import { mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { isAbsolute, join } from 'node:path';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { decodeExportCursor, encodeExportCursor } from '@chrischall/mcp-utils/export';
import { McpToolError, surfaceToolHints } from '@chrischall/mcp-utils';
import type { ChatServiceClient, GetMessagesOptions, RawConversation, RawMessage } from '../../src/chatsvc.js';
import { registerExportTools, type ExportToolOptions } from '../../src/tools/export.js';
import { fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;

const ALPHA = '19:alpha0001@thread.v2';
const UNTITLED = '19:untitled01@thread.v2';
const CHANNEL = '19:design01@thread.tacv2';
const TEAM = '19:team0001@thread.tacv2';
const MEETING = '19:meeting_abc@thread.v2';

/** A message from a named person (MRI derived from the name). */
function said(name: string, at: number, conv: string, opts: { text?: string; rootId?: number; version?: number } = {}): RawMessage {
  const mri = `8:orgid:${name.toLowerCase().replace(/[^a-z]/g, '')}`;
  return { ...rawMsg(at, { conv, sender: name, ...opts }), from: FROM(mri) };
}

const CONVS: RawConversation[] = [
  rawConv(ALPHA, { topic: 'Project Alpha', at: T0 + 3 * MIN }),
  rawConv(UNTITLED, { at: T0 + 4 * MIN, sender: 'Rivera, Sam' }),
  rawConv(CHANNEL, { topic: 'Design', spaceId: TEAM, at: T0 + 5 * MIN }),
  rawConv(TEAM, { topic: 'Example Team', spaceId: TEAM, at: T0 - 100 * MIN }),
  { ...rawConv(MEETING, { topic: 'Weekly sync', at: T0 + 2 * MIN }) },
];

const MESSAGES: Record<string, RawMessage[]> = {
  [ALPHA]: [
    said('Lee, Pat', T0, ALPHA, { text: 'alpha one' }),
    rawMsg(T0 + MIN, { conv: ALPHA, type: 'Event/Call' }),
    rawMsg(T0 + 2 * MIN, { conv: ALPHA, fromMe: true, text: 'alpha mine' }),
    said('Lee, Pat', T0 + 3 * MIN, ALPHA, { text: 'alpha three' }),
  ],
  [UNTITLED]: [
    said('Kim, Jo', T0 + MIN, UNTITLED, { text: 'untitled one' }),
    said('Rivera, Sam', T0 + 4 * MIN, UNTITLED, { text: 'untitled two' }),
  ],
  [CHANNEL]: [
    said('Lee, Pat', T0, CHANNEL, { text: 'root post', rootId: T0 }),
    said('Kim, Jo', T0 + MIN, CHANNEL, { text: 'a reply', rootId: T0 }),
    said('Lee, Pat', T0 + 5 * MIN, CHANNEL, { text: 'second root', rootId: T0 + 5 * MIN }),
  ],
  [TEAM]: [],
  [MEETING]: [said('Lee, Pat', T0 + 2 * MIN, MEETING, { text: 'meeting chat' })],
};

function perConversation(byId: Record<string, RawMessage[] | Error | (() => Error | RawMessage[])>) {
  return vi.fn(async (id: string, opts: GetMessagesOptions) => {
    let v = byId[id] ?? [];
    if (typeof v === 'function') v = v();
    if (v instanceof Error) throw v;
    return messagesFrom(v)(id, opts);
  });
}

function chatsvcWith(
  convs: RawConversation[] = CONVS,
  byId: Record<string, RawMessage[] | Error | (() => Error | RawMessage[])> = MESSAGES,
): ChatServiceClient {
  return fakeChatsvc(convs, { getMessages: perConversation(byId), meDisplayName: vi.fn().mockResolvedValue('Me Myself') });
}

let outDir: string;
beforeEach(() => {
  outDir = realpathSync(mkdtempSync(join(tmpdir(), 'teams-export-')));
  vi.stubEnv('TEAMS_OUTPUT_DIR', outDir);
  vi.stubEnv('MCP_BLOB_BASE_URL', '');
  vi.stubEnv('MCP_BLOB_SIGNING_KEY', '');
  vi.stubEnv('TEAMS_TIME_ZONE', '');
});
afterEach(() => {
  vi.unstubAllEnvs();
  rmSync(outDir, { recursive: true, force: true });
});

const noSleep: ExportToolOptions = { sleep: async () => {} };

async function call(chatsvc: ChatServiceClient, args: Record<string, unknown> = {}, opts: ExportToolOptions = noSleep) {
  const harness = await createTestHarness((server) => registerExportTools(server, { chatsvc }, opts));
  try {
    const result = await harness.callTool('teams_export_messages', args);
    const text = (result.content as { type: string; text: string }[])[0]?.text ?? '';
    return {
      result,
      text,
      data: result.isError ? undefined : (parseToolResult(result) as Record<string, any>),
    };
  } finally {
    await harness.close();
  }
}

const short = (id: string) => createHash('sha256').update(id).digest('hex').slice(0, 8);
const lines = (path: string) =>
  readFileSync(path, 'utf8')
    .trim()
    .split('\n')
    .map((l) => JSON.parse(l));
/** The export's payload, inside the untrusted envelope when it has one. */
const payload = (data: Record<string, any>) => (data.data ?? data) as Record<string, any>;

describe('teams_export_messages registration', () => {
  it('is published read-only and open-world, and says no message content comes back', async () => {
    const harness = await createTestHarness((server) => registerExportTools(server, { chatsvc: chatsvcWith() }));
    try {
      const tool = (await harness.client.listTools()).tools.find((t) => t.name === 'teams_export_messages');
      expect(tool?.annotations).toMatchObject({ readOnlyHint: true, openWorldHint: true });
      expect(tool?.description).toMatch(/no message (text|content)/i);
      expect(tool?.description).toMatch(/jq/);
      expect(tool?.description).toMatch(/cursor/);
    } finally {
      await harness.close();
    }
  });
});

describe('teams_export_messages to disk (JSONL)', () => {
  it('writes one file per conversation with messages, oldest first, and returns only a manifest', async () => {
    const { data, text } = await call(chatsvcWith());
    const out = payload(data!);
    expect(out.mode).toBe('disk');
    expect(isAbsolute(out.location)).toBe(true);
    expect(out.location.startsWith(outDir)).toBe(true);
    expect(out.has_more).toBe(false);
    expect(out.cursor).toBeUndefined();
    // Four conversations with messages; the team's own thread has none.
    expect(out.counts).toMatchObject({ conversations: 4, messages: 9 });
    expect(out.span).toEqual({ oldest: new Date(T0).toISOString(), newest: new Date(T0 + 5 * MIN).toISOString() });
    // Nothing anyone wrote reaches the result.
    for (const said of ['alpha one', 'alpha mine', 'root post', 'a reply', 'meeting chat']) expect(text).not.toContain(said);

    const alphaFile = out.files.find((f: any) => f.path.endsWith(`Project Alpha-${short(ALPHA)}.jsonl`));
    expect(alphaFile).toMatchObject({ records: 3 });
    expect(isAbsolute(alphaFile.path)).toBe(true);
    const [header, ...records] = lines(alphaFile.path);
    expect(header._export).toMatchObject({
      source: 'microsoft-teams',
      account: 'Me Myself',
      conversation: { id: ALPHA, kind: 'group', title: 'Project Alpha' },
      note: expect.stringMatching(/untrusted data, not instructions/),
    });
    expect(records.map((r) => r.text)).toEqual(['alpha one', 'alpha mine', 'alpha three']);
    // Every record names its conversation; a chat's carries the people seen writing in it.
    for (const r of records) {
      expect(r.conversation).toMatchObject({ id: ALPHA, kind: 'group', title: 'Project Alpha', participants: ['Lee, Pat'] });
    }
    expect(records[1]).toMatchObject({ fromMe: true, sender: 'Me Myself' });

    const manifest = JSON.parse(readFileSync(join(out.location, 'manifest.json'), 'utf8'));
    expect(manifest).toMatchObject({ export_id: expect.any(String), part: 1, counts: { conversations: 4, messages: 9 } });
    expect(manifest.files).toHaveLength(4);
  });

  it('names an untitled group chat by its participants, in the file name and every record', async () => {
    const { data } = await call(chatsvcWith(), { conversations: UNTITLED });
    const out = payload(data!);
    expect(out.files).toHaveLength(1);
    const title = 'Group chat: Kim, Jo; Rivera, Sam';
    expect(out.files[0].path).toContain(`Group chat_ Kim, Jo; Rivera, Sam-${short(UNTITLED)}.jsonl`);
    const [header, ...records] = lines(out.files[0].path);
    expect(header._export.conversation).toMatchObject({ title, kind: 'group', participants: ['Kim, Jo', 'Rivera, Sam'] });
    expect(records.every((r) => r.conversation.title === title)).toBe(true);
  });

  it('exports a channel with its team and no participants', async () => {
    const { data } = await call(chatsvcWith(), { kinds: ['channel'] });
    const out = payload(data!);
    expect(out.files).toHaveLength(1);
    const [header, ...records] = lines(out.files[0].path);
    expect(header._export.conversation).toEqual({ id: CHANNEL, kind: 'channel', title: 'Design', team: 'Example Team' });
    expect(records.map((r) => [r.text, r.replyTo])).toEqual([
      ['root post', undefined],
      ['a reply', String(T0)],
      ['second root', undefined],
    ]);
  });

  it('filters by kind: chat covers group chats, 1:1s and the self-chat but not meetings', async () => {
    const { data } = await call(chatsvcWith(), { kinds: ['chat'] });
    const ids = payload(data!).files.map((f: any) => lines(f.path)[0]._export.conversation.id).sort();
    expect(ids).toEqual([ALPHA, UNTITLED]);
    const { data: meetings } = await call(chatsvcWith(), { kinds: ['meeting'] });
    expect(payload(meetings!).files).toHaveLength(1);
    expect(lines(payload(meetings!).files[0].path)[0]._export.conversation).toMatchObject({ id: MEETING, kind: 'meeting' });
  });

  it('takes conversations by name or id, and an ambiguous name writes nothing', async () => {
    const { data } = await call(chatsvcWith(), { conversations: ['alpha', CHANNEL] });
    expect(payload(data!).files).toHaveLength(2);

    const before = readdirSync(outDir).length;
    const amb = await call(chatsvcWith(), { conversations: ['a'] });
    expect(payload(amb.data!)).toMatchObject({ ambiguous: true, query: 'a' });
    expect(readdirSync(outDir)).toHaveLength(before);
  });

  it('refuses an unknown conversation name before writing anything', async () => {
    const before = readdirSync(outDir).length;
    const { result } = await call(chatsvcWith(), { conversations: 'no such chat' });
    expect(result.isError).toBe(true);
    expect(readdirSync(outDir)).toHaveLength(before);
  });

  it('applies since and until by arrival, reading back only to since by last change', async () => {
    const chatsvc = chatsvcWith();
    const { data } = await call(chatsvc, {
      conversations: ALPHA,
      since: new Date(T0 + MIN).toISOString(),
      until: new Date(T0 + 2 * MIN).toISOString(),
    });
    const [, ...records] = lines(payload(data!).files[0].path);
    expect(records.map((r) => r.text)).toEqual(['alpha mine']);
    const opts = (chatsvc.getMessages as any).mock.calls.at(-1)[1] as GetMessagesOptions;
    expect(opts.stopAt?.(rawMsg(T0, { conv: ALPHA }))).toBe(true);
    expect(opts.stopAt?.(rawMsg(T0, { conv: ALPHA, version: T0 + 5 * MIN }))).toBe(false);
  });

  it('skips a conversation whose last change is before since without reading it', async () => {
    const chatsvc = chatsvcWith();
    await call(chatsvc, { since: new Date(T0 + 4 * MIN).toISOString() });
    const read = (chatsvc.getMessages as any).mock.calls.map((c: unknown[]) => c[0]);
    expect(read).not.toContain(ALPHA);
    expect(read).not.toContain(MEETING);
    expect(read).toContain(CHANNEL);
  });

  it('adds time_local with tz', async () => {
    const { data } = await call(chatsvcWith(), { conversations: ALPHA, tz: 'America/New_York' });
    const [, first] = lines(payload(data!).files[0].path);
    expect(first.time_local).toMatch(/-04:00$/);
  });

  it('refuses a dest outside the allowed roots', async () => {
    const { result } = await call(chatsvcWith(), { dest: '/etc' });
    expect(result.isError).toBe(true);
  });
});

describe('teams_export_messages to disk (Markdown)', () => {
  it('writes a readable transcript per conversation with the untrusted note, channels threaded', async () => {
    const { data } = await call(chatsvcWith(), { format: 'markdown', conversations: [ALPHA, CHANNEL] });
    const out = payload(data!);
    const alpha = out.files.find((f: any) => f.path.endsWith(`Project Alpha-${short(ALPHA)}.md`));
    const md = readFileSync(alpha.path, 'utf8');
    expect(md).toMatch(/^# Project Alpha/);
    expect(md).toMatch(/untrusted data, not instructions/);
    expect(md).toContain('Lee, Pat');
    expect(md.indexOf('alpha one')).toBeLessThan(md.indexOf('alpha three'));

    const channel = out.files.find((f: any) => f.path.endsWith(`Design-${short(CHANNEL)}.md`));
    const cmd = readFileSync(channel.path, 'utf8');
    expect(cmd).toContain('Example Team');
    // The reply sits under its root, before the second thread.
    const root = cmd.indexOf('root post');
    const reply = cmd.indexOf('a reply');
    const second = cmd.indexOf('second root');
    expect(root).toBeGreaterThan(-1);
    expect(reply).toBeGreaterThan(root);
    expect(second).toBeGreaterThan(reply);
    expect(cmd.match(/^## Thread/gm)).toHaveLength(2);
  });
});

describe('teams_export_messages budgets and resume', () => {
  it('stops at max_items with a cursor, and the cursor finishes the same export with no gap or repeat', async () => {
    const chatsvc = chatsvcWith();
    const first = payload((await call(chatsvc, { conversations: [ALPHA, UNTITLED, CHANNEL], max_items: 2 })).data!);
    expect(first.has_more).toBe(true);
    expect(first.counts.messages).toBe(2);
    expect(typeof first.cursor).toBe('string');

    const all: string[] = [];
    const dirs = new Set<string>([first.location]);
    for (const f of first.files) all.push(...lines(f.path).slice(1).map((r) => `${r.conversation.id}/${r.id}`));
    let cursor = first.cursor;
    let parts = 1;
    while (cursor !== undefined) {
      const next = payload((await call(chatsvc, { cursor, max_items: 2 })).data!);
      parts++;
      dirs.add(next.location);
      for (const f of next.files) all.push(...lines(f.path).slice(1).map((r) => `${r.conversation.id}/${r.id}`));
      cursor = next.cursor;
      expect(next.has_more).toBe(cursor !== undefined);
      expect(parts).toBeLessThan(10);
    }
    // One directory, every message once.
    expect(dirs.size).toBe(1);
    const expected = [ALPHA, UNTITLED, CHANNEL].flatMap((id) =>
      MESSAGES[id].filter((m) => (m.messagetype ?? 'Text') === 'Text').map((m) => `${id}/${m.id}`),
    );
    expect(all.sort()).toEqual(expected.sort());
    // A conversation split across parts continues in a part-numbered file beside the first.
    const names = readdirSync(first.location);
    expect(names.some((n) => /\.part-0002\.jsonl$/.test(n))).toBe(true);
    expect(names).toContain('manifest.json');
    expect(names).toContain('manifest.part-0002.json');
  });

  it('stops between conversations once budget_seconds is spent', async () => {
    let now = T0;
    const chatsvc = chatsvcWith(CONVS, {
      ...MESSAGES,
      [ALPHA]: () => {
        now += 61_000;
        return MESSAGES[ALPHA];
      },
    });
    const { data } = await call(chatsvc, { budget_seconds: 60 }, { ...noSleep, now: () => now });
    const out = payload(data!);
    expect(out.has_more).toBe(true);
    expect(out.files).toHaveLength(1);
    expect(out.cursor).toBeDefined();
  });

  it('stops inside a long read once budget_seconds is spent, and the cursor re-reads that conversation', async () => {
    let now = T0;
    // The channel arrives a page (one message) at a time, and each page takes 61s.
    const paged = messagesFrom(MESSAGES[CHANNEL], 1);
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) =>
      id === CHANNEL
        ? paged(id, {
            ...opts,
            onPage: async (n, m) => {
              now += 61_000;
              await opts.onPage?.(n, m);
            },
          })
        : messagesFrom(MESSAGES[id] ?? [])(id, opts),
    );
    const chatsvc = fakeChatsvc(CONVS, { getMessages, meDisplayName: vi.fn().mockResolvedValue('Me Myself') });
    const first = payload(
      (await call(chatsvc, { conversations: [ALPHA, CHANNEL], budget_seconds: 60 }, { ...noSleep, now: () => now })).data!,
    );
    expect(first.has_more).toBe(true);
    expect(first.files.map((f: any) => f.path)).toEqual([expect.stringContaining(short(ALPHA))]);
    expect(first.errors ?? []).toEqual([]);
    // Only one 61s page went by before the read was cut, not all three.
    expect(now).toBe(T0 + 61_000);

    const next = payload((await call(chatsvc, { cursor: first.cursor })).data!);
    expect(next.has_more).toBe(false);
    expect(next.files.map((f: any) => f.path)).toEqual([expect.stringContaining(short(CHANNEL))]);
    expect(next.counts.messages).toBe(3);
  });

  it('never cuts the first conversation of a call, so every call makes progress', async () => {
    let now = T0;
    const paged = messagesFrom(MESSAGES[CHANNEL], 1);
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) =>
      paged(id, {
        ...opts,
        onPage: async (n, m) => {
          now += 61_000;
          await opts.onPage?.(n, m);
        },
      }),
    );
    const chatsvc = fakeChatsvc(CONVS, { getMessages, meDisplayName: vi.fn().mockResolvedValue('Me Myself') });
    const out = payload(
      (await call(chatsvc, { conversations: [CHANNEL], budget_seconds: 60 }, { ...noSleep, now: () => now })).data!,
    );
    expect(out.has_more).toBe(false);
    expect(out.counts.messages).toBe(3);
  });

  it('reports progress per page of a long read, not only per conversation', async () => {
    const paged = messagesFrom(MESSAGES[CHANNEL], 1);
    const chatsvc = fakeChatsvc(CONVS, {
      getMessages: vi.fn(async (id: string, opts: GetMessagesOptions) => paged(id, opts)),
      meDisplayName: vi.fn().mockResolvedValue('Me Myself'),
    });
    // surfaceToolHints is what makes the caller's progress token ambient, as the real server does.
    const harness = await createTestHarness((server) => {
      surfaceToolHints(server);
      registerExportTools(server, { chatsvc }, noSleep);
    });
    const messages: string[] = [];
    try {
      await harness.client.callTool({ name: 'teams_export_messages', arguments: { conversations: [CHANNEL] } }, {
        onprogress: (p) => messages.push(p.message ?? ''),
      });
    } finally {
      await harness.close();
    }
    expect(messages.filter((m) => /page/.test(m))).toHaveLength(2);
    expect(messages.at(-1)).toMatch(/1\/1 conversations/);
  });

  it('refuses other arguments beside a cursor', async () => {
    const first = payload((await call(chatsvcWith(), { max_items: 1 })).data!);
    const { result } = await call(chatsvcWith(), { cursor: first.cursor, kinds: ['channel'] });
    expect(result.isError).toBe(true);
  });

  it('never trusts the directory in a cursor', async () => {
    const first = payload((await call(chatsvcWith(), { max_items: 1 })).data!);
    const state = decodeExportCursor(first.cursor);
    const { v: _v, ...rest } = state;
    const forged = encodeExportCursor({ ...rest, dir: '/etc' });
    const { result } = await call(chatsvcWith(), { cursor: forged });
    expect(result.isError).toBe(true);
  });

  it('refuses a cursor that is not one', async () => {
    const { result } = await call(chatsvcWith(), { cursor: 'not-a-cursor' });
    expect(result.isError).toBe(true);
  });
});

describe('teams_export_messages failures', () => {
  it('records a conversation that cannot be read and exports the rest', async () => {
    const forbidden = new McpToolError('Teams chat service returned HTTP 403.', { status: 403, kind: 'http' });
    const { data } = await call(chatsvcWith(CONVS, { ...MESSAGES, [ALPHA]: forbidden }));
    const out = payload(data!);
    expect(out.has_more).toBe(false);
    expect(out.files).toHaveLength(3);
    expect(out.errors).toEqual([expect.objectContaining({ conversation: expect.objectContaining({ id: ALPHA }), error: expect.stringMatching(/403/) })]);
  });

  it('stops on a rejected credential, keeps what was written, and resumes from there', async () => {
    let rejected = true;
    const chatsvc = chatsvcWith(CONVS, {
      ...MESSAGES,
      [CHANNEL]: () =>
        rejected
          ? new McpToolError('Teams rejected the access token (401) even after a fresh capture.', { status: 401, kind: 'credential_rejected' })
          : MESSAGES[CHANNEL],
    });
    const first = payload((await call(chatsvc, { conversations: [ALPHA, CHANNEL, MEETING] })).data!);
    expect(first.has_more).toBe(true);
    expect(first.cursor).toBeDefined();
    expect(first.files).toHaveLength(1);
    expect(first.errors?.[0]?.error).toMatch(/401/);
    rejected = false;
    const second = payload((await call(chatsvc, { cursor: first.cursor })).data!);
    expect(second.has_more).toBe(false);
    expect(second.location).toBe(first.location);
    expect(second.files).toHaveLength(2);
  });

  it('retries a rate-limited read with backoff, then goes on', async () => {
    let tries = 0;
    const sleep = vi.fn(async () => {});
    const chatsvc = chatsvcWith(CONVS, {
      ...MESSAGES,
      [ALPHA]: () =>
        ++tries < 3 ? new McpToolError('Teams chat service is rate-limiting requests (429).', { status: 429, kind: 'http' }) : MESSAGES[ALPHA],
    });
    const out = payload((await call(chatsvc, { conversations: ALPHA }, { sleep })).data!);
    expect(out.has_more).toBe(false);
    expect(out.files).toHaveLength(1);
    expect(sleep).toHaveBeenCalledTimes(2);
  });

  it('stops with a cursor when rate limiting outlasts the retries', async () => {
    const chatsvc = chatsvcWith(CONVS, {
      ...MESSAGES,
      [UNTITLED]: new McpToolError('Teams chat service is rate-limiting requests (429).', { status: 429, kind: 'http' }),
    });
    const out = payload((await call(chatsvc, { conversations: [ALPHA, UNTITLED] })).data!);
    expect(out.has_more).toBe(true);
    expect(out.cursor).toBeDefined();
    expect(out.errors?.[0]?.error).toMatch(/429/);
  });
});

describe('teams_export_messages to the hosted blob store', () => {
  it('uploads one zip and returns a signed link', async () => {
    vi.stubEnv('MCP_BLOB_BASE_URL', 'https://host.example/b/reg123');
    vi.stubEnv('MCP_BLOB_SIGNING_KEY', 'k'.repeat(32));
    const puts: { url: string; body: Uint8Array; type: string }[] = [];
    const fetchImpl = vi.fn(async (url: string | URL | Request, init?: RequestInit) => {
      puts.push({ url: String(url), body: init?.body as Uint8Array, type: new Headers(init?.headers).get('content-type') ?? '' });
      return new Response(null, { status: 201 });
    }) as unknown as typeof fetch;
    const out = payload((await call(chatsvcWith(), { conversations: ALPHA }, { ...noSleep, fetchImpl })).data!);
    expect(out.mode).toBe('blob');
    expect(out.location).toMatch(/^https:\/\/host\.example\/b\/reg123\/exp\/[A-Za-z0-9_-]+\/part-0001\.zip\?exp=\d+&sig=/);
    expect(out.url_expires_at).toBeDefined();
    expect(out.files.map((f: any) => f.name)).toContain(`Project Alpha-${short(ALPHA)}.jsonl`);
    expect(puts).toHaveLength(1);
    expect(puts[0].type).toBe('application/zip');
    // The link is the gateway's GET signature over `<prefix>/<key>\n<exp>`.
    const url = new URL(out.location);
    const key = url.pathname.replace('/b/reg123/', '');
    const sig = createHmac('sha256', 'k'.repeat(32)).update(`reg123/${key}\n${url.searchParams.get('exp')}`).digest('base64url');
    expect(url.searchParams.get('sig')).toBe(sig);
    // The archive holds the conversation file (a local header names it).
    const zip = Buffer.from(puts[0].body);
    expect(zip.includes(Buffer.from(`Project Alpha-${short(ALPHA)}.jsonl`))).toBe(true);
    expect(readdirSync(outDir)).toHaveLength(0);
  });
});

describe('teams_export_messages cursor size', () => {
  it('keeps the cursor within the codec limit however many conversations remain', async () => {
    const many = Array.from({ length: 600 }, (_, i) => `19:${'x'.repeat(40)}${String(i).padStart(4, '0')}@thread.v2`);
    const convs = many.map((id, i) => rawConv(id, { topic: `Chat ${i}`, at: T0 + i }));
    const byId = Object.fromEntries(many.map((id, i) => [id, [rawMsg(T0 + i, { conv: id })]]));
    const chatsvc = chatsvcWith(convs, byId);
    const first = payload((await call(chatsvc, { max_items: 1 })).data!);
    expect(first.has_more).toBe(true);
    const second = payload((await call(chatsvc, { cursor: first.cursor, max_items: 1 })).data!);
    expect(second.has_more).toBe(true);
    const a = lines(first.files[0].path)[0]._export.conversation.id;
    const b = lines(second.files[0].path)[0]._export.conversation.id;
    expect(a).not.toBe(b);
    expect([a, b].sort()).toEqual([...many].sort().slice(0, 2));
  });
});
