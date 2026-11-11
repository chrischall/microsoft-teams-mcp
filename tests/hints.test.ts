import { afterEach, beforeEach, describe, it, expect, vi } from 'vitest';
import { createTestHarness } from '@chrischall/mcp-utils/test';
import { servedTools } from './served-tools.js';
import { registerChangesTools } from '../src/tools/changes.js';
import { DAY, fakeChatsvc, messagesFrom, pinClock, rawConv, rawMsg } from './tools/fake-chatsvc.js';

/**
 * Context-group hints (`_meta.mint`, mcp-host spec 2026-11-03 D4), read off the
 * SERVED `tools/list` — what a host sees — so a new tool cannot skip the
 * decision.
 */
type Listed = {
  name: string;
  annotations?: { readOnlyHint?: boolean };
  inputSchema: { properties?: Record<string, { enum?: unknown[] }> };
  _meta?: { mint?: Record<string, unknown> };
};

const DOMAINS = ['school', 'home', 'work', 'family', 'health', 'money', 'travel', 'food'];
const INTENTS = ['whats-new', 'whats-due', 'whats-on', 'lookup', 'turn-on-off', 'send', 'change', 'other'];
const KEYS = ['domain', 'intent', 'canonical', 'subject', 'when', 'aliases'];
const CANONICAL_INTENTS = ['whats-new', 'whats-due', 'whats-on'];

/** The intent every tool declares — this table is the assertion. */
const EXPECTED_INTENT: Record<string, string> = {
  teams_list_conversations: 'lookup',
  teams_list_channels: 'lookup',
  teams_list_messages: 'lookup',
  teams_list_unread_messages: 'whats-new',
  teams_list_changed_conversations: 'whats-new',
  teams_list_activity: 'whats-new',
  teams_search_messages: 'lookup',
  teams_get_attachment: 'other',
  teams_export_messages: 'other',
  teams_get_transcript: 'lookup',
  teams_get_meeting_recap: 'lookup',
  teams_send_message: 'send',
  teams_react: 'send',
  teams_rename_chat: 'send',
  teams_add_chat_members: 'send',
  teams_remove_chat_members: 'send',
  teams_create_group_chat: 'send',
  teams_mark_read: 'change',
  teams_edit_message: 'change',
  teams_delete_message: 'change',
};

const served = async () => (await servedTools()) as unknown as Listed[];
const mint = (t: Listed) => t._meta?.mint as Record<string, unknown> | undefined;

describe('_meta.mint on tools/list', () => {
  it('routes "transcript" to the meeting-transcript tool, not the chat export', async () => {
    const tools = await served();
    const owners = tools.filter((t) => ((mint(t)?.aliases as string[] | undefined) ?? []).includes('transcript')).map((t) => t.name);
    expect(owners).toEqual(['teams_get_transcript']);
  });

  it('hints every served tool except the healthcheck, which joins no group', async () => {
    const tools = await served();
    const unhinted = tools.filter((t) => mint(t) === undefined).map((t) => t.name);
    expect(unhinted).toEqual(['teams_healthcheck']);
    expect(tools.filter((t) => mint(t) !== undefined).map((t) => t.name).sort()).toEqual(Object.keys(EXPECTED_INTENT).sort());
  });

  it.each(Object.entries(EXPECTED_INTENT))('%s declares domain work and intent %s', async (name, intent) => {
    const tool = (await served()).find((t) => t.name === name)!;
    expect(mint(tool)?.domain).toBe('work');
    expect(mint(tool)?.intent).toBe(intent);
  });

  it('every hint is shaped per the host contract', async () => {
    for (const t of await served()) {
      const m = mint(t);
      if (m === undefined) continue;
      const props = t.inputSchema.properties ?? {};
      expect(Object.keys(m).filter((k) => !KEYS.includes(k)), t.name).toEqual([]);
      expect(DOMAINS, t.name).toContain(m.domain);
      expect(INTENTS, t.name).toContain(m.intent);
      if ('canonical' in m) expect(m.canonical, t.name).toBe(true);
      if ('subject' in m) expect(props, t.name).toHaveProperty(m.subject as string);
      if ('when' in m) {
        const when = m.when as Record<string, string>;
        expect(Object.keys(when).sort(), t.name).toEqual(['end', 'start']);
        expect(props, t.name).toHaveProperty(when.start);
        expect(props, t.name).toHaveProperty(when.end);
      }
      if ('aliases' in m) {
        const aliases = m.aliases as string[];
        expect(aliases.length, t.name).toBeGreaterThan(0);
        expect(aliases.length, t.name).toBeLessThanOrEqual(8);
        for (const a of aliases) {
          expect(a.length, t.name).toBeLessThanOrEqual(40);
          expect(a, t.name).toBe(a.toLowerCase());
          expect(a, t.name).not.toMatch(/^teams_/);
        }
      }
    }
  });

  it('marks canonical only read-only what\'s-new/due/on reads, at most one per intent', async () => {
    const canonical = (await served()).filter((t) => mint(t)?.canonical === true);
    for (const t of canonical) {
      expect(t.annotations?.readOnlyHint, t.name).toBe(true);
      expect(CANONICAL_INTENTS, t.name).toContain(mint(t)?.intent);
    }
    for (const intent of CANONICAL_INTENTS) {
      expect(canonical.filter((t) => mint(t)?.intent === intent).length, intent).toBeLessThanOrEqual(1);
    }
    expect(canonical.map((t) => [t.name, mint(t)?.intent])).toEqual([['teams_list_changed_conversations', 'whats-new']]);
  });

  it('every canonical read accepts view: brief', async () => {
    for (const t of (await served()).filter((x) => mint(x)?.canonical === true)) {
      expect(t.inputSchema.properties?.view?.enum, t.name).toContain('brief');
    }
  });

  it('the canonical what\'s-new names no subject (one account, no listable subjects) and no range (it takes a cursor, not a start and end)', async () => {
    const tool = (await served()).find((t) => t.name === 'teams_list_changed_conversations')!;
    expect(mint(tool)).not.toHaveProperty('subject');
    expect(mint(tool)).not.toHaveProperty('when');
  });
});

describe('teams_list_changed_conversations view: brief', () => {
  const NOW = 1_760_000_000_000;
  const HOUR = 3_600_000;
  const GROUP = '19:group0001@thread.v2';
  const LONG = '19:group0002@thread.v2';
  const BROKEN = '19:group0003@thread.v2';
  const NOTE_LIKE = '19:group0004@thread.v2';
  const SLASHED = '19:group0005@thread.v2';
  const LONG_TITLE = `Quarterly planning\nfor the ${'very '.repeat(30)}long project`;

  beforeEach(() => pinClock(NOW));
  afterEach(() => vi.useRealTimers());

  async function call(args: Record<string, unknown>, convs = CONVS) {
    const messages: Record<string, ReturnType<typeof rawMsg>[]> = {
      [GROUP]: [rawMsg(NOW - 2 * HOUR, { text: 'one' }), rawMsg(NOW - HOUR, { text: 'two', fromMe: true })],
      [LONG]: [rawMsg(NOW - 3 * HOUR, { conv: LONG })],
      [NOTE_LIKE]: [rawMsg(NOW - 5 * HOUR, { conv: NOTE_LIKE })],
      [SLASHED]: [rawMsg(NOW - 6 * HOUR, { conv: SLASHED })],
    };
    const getMessages = vi.fn(async (id: string, opts: Parameters<ReturnType<typeof messagesFrom>>[1]) => {
      if (id === BROKEN) throw new Error('boom');
      return messagesFrom(messages[id] ?? [])(id, opts);
    });
    const chatsvc = fakeChatsvc(convs, { getMessages });
    const h = await createTestHarness((server) => registerChangesTools(server, { chatsvc }));
    try {
      const result = await h.callTool('teams_list_changed_conversations', args);
      return ((result.content as { type: string; text: string }[])[0]).text;
    } finally {
      await h.close();
    }
  }

  const CONVS = [
    rawConv(GROUP, { topic: 'Project Falcon', at: NOW - HOUR }),
    rawConv(LONG, { topic: LONG_TITLE, at: NOW - 3 * HOUR }),
    rawConv(BROKEN, { topic: 'Broken one', at: NOW - 4 * HOUR }),
    rawConv('19:quiet@thread.v2', { topic: 'Quiet', at: NOW - 2 * DAY }),
  ];

  /** The brief answer's raw text, and its lines out of the untrusted-data envelope. */
  async function brief(args: Record<string, unknown> = { view: 'brief' }, convs = CONVS) {
    const text = await call(args, convs);
    const parsed = JSON.parse(text) as { untrusted_content?: unknown; note?: unknown; data?: unknown };
    return { text, parsed, lines: parsed.data as string[] };
  }

  it('answers inside the untrusted-data envelope, minified, note first, lines as data: string[]', async () => {
    const { text, parsed, lines } = await brief();
    expect(parsed.untrusted_content).toBe(true);
    expect(parsed.note).toMatch(/other people in Microsoft Teams/);
    expect(Object.keys(parsed)).toEqual(['untrusted_content', 'note', 'data']);
    expect(text).toBe(JSON.stringify(parsed));
    expect(Array.isArray(lines) && lines.every((l) => typeof l === 'string')).toBe(true);
  });

  it('answers one line per conversation, "title — <ISO date | no date> — <status>", each ≤ 120 chars', async () => {
    const { lines } = await brief();
    expect(lines.slice(0, 2)).toEqual([
      `Project Falcon — ${new Date(NOW - HOUR).toISOString()} — 2 new, 1 from others`,
      expect.stringMatching(/^Quarterly planning for the very .*… — \S+ — 1 new, 1 from others$/),
    ]);
    for (const line of lines.filter((l) => !l.startsWith('-- '))) {
      expect(line.length).toBeLessThanOrEqual(120);
      expect(line).toMatch(/^.+ — (\d{4}-\d{2}-\d{2}T\S+|no date) — .+$/);
    }
  });

  it('reports a conversation whose read failed as an "-- error:" note, not as an item', async () => {
    const { lines } = await brief();
    expect(lines).toHaveLength(3);
    expect(lines[2]).toBe('-- error: 1 conversation could not be read.');
    expect(lines.join('\n')).not.toContain('Broken one');
  });

  it('answers the single note "-- none" when nothing is new', async () => {
    const { lines } = await brief({ view: 'brief', cursor: new Date(NOW).toISOString() });
    expect(lines).toEqual(['-- none']);
  });

  it('escapes a title that starts with "-- " or "\\" with one leading backslash', async () => {
    const { lines } = await brief({ view: 'brief' }, [
      rawConv(NOTE_LIKE, { topic: '-- none', at: NOW - 5 * HOUR }),
      rawConv(SLASHED, { topic: '\\server\\share', at: NOW - 6 * HOUR }),
    ]);
    expect(lines).toEqual([
      `\\-- none — ${new Date(NOW - 5 * HOUR).toISOString()} — 1 new, 1 from others`,
      `\\\\server\\share — ${new Date(NOW - 6 * HOUR).toISOString()} — 1 new, 1 from others`,
    ]);
  });

  it('keeps the full JSON envelope without view, and with view: full', async () => {
    for (const args of [{}, { view: 'full' }]) {
      const parsed = JSON.parse(await call(args)) as Record<string, unknown>;
      expect(parsed.untrusted_content, JSON.stringify(args)).toBe(true);
      expect(parsed.conversations, JSON.stringify(args)).toBeDefined();
    }
  });
});
