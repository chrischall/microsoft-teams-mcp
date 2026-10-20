import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError, RequestTimeoutError, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import { CHECKMARK_KEY, reactionKey, registerReactTools } from '../../src/tools/react.js';
import { UNTRUSTED_CONTENT_NOTE } from '../../src/tools/untrusted.js';
import type { ChatServiceClient, RawMessage } from '../../src/chatsvc.js';
import { ME, fakeChatsvc, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;

const GROUP = '19:group0001@thread.v2';
const OTHER_GROUP = '19:group0002@thread.v2';
const MSG = String(T0 + 1 * MIN);

const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * MIN }),
  rawConv(OTHER_GROUP, { topic: 'Falcon launch party', at: T0 }),
];

type Emotion = { key: string; users: { mri: string; time: number }[] };

/**
 * A chat service holding one message whose `emotions` property the reaction
 * calls really change — unless `ignore` is set, which models a write Teams
 * accepted but that does not show on reading back.
 */
function fakeWith(opts: { ignore?: boolean; emotions?: Emotion[]; overrides?: Record<string, unknown> } = {}) {
  let emotions: Emotion[] = opts.emotions ?? [];
  const message = (): RawMessage => ({
    ...rawMsg(Number(MSG), { text: 'ship it?', sender: 'Bea Example' }),
    properties: { emotions: JSON.stringify(emotions) },
  });
  const setReaction = vi.fn(async (_c: string, _m: string, key: string) => {
    if (opts.ignore) return;
    emotions = [...emotions.filter((e) => e.key !== key), { key, users: [{ mri: ME, time: T0 }] }];
  });
  const removeReaction = vi.fn(async (_c: string, _m: string, key: string) => {
    if (opts.ignore) return;
    emotions = emotions.filter((e) => e.key !== key);
  });
  const getMessage = vi.fn(async (_c: string, id: string) => (id === MSG ? message() : undefined));
  const chatsvc = fakeChatsvc(CONVS, { setReaction, removeReaction, getMessage, ...opts.overrides });
  return { chatsvc, setReaction, removeReaction, getMessage };
}

async function harnessFor(chatsvc: ChatServiceClient) {
  return createTestHarness((server) => registerReactTools(server, { chatsvc }));
}

type PhaseOne = {
  confirmToken?: string;
  preview?: { method?: string; path?: string; willSend?: unknown } & Record<string, unknown>;
};

function errorText(result: { content?: unknown[] }): string {
  return (result.content?.[0] as { text?: string })?.text ?? '';
}

/** Phase 1 then phase 2 with its token. */
async function reactConfirmed(chatsvc: ChatServiceClient, args: Record<string, unknown>) {
  const h = await harnessFor(chatsvc);
  try {
    const first = await h.callTool('teams_react', args);
    if (first.isError) return { first, second: first, preview: undefined, data: undefined };
    const preview = parseToolResult<PhaseOne>(first);
    const second = await h.callTool('teams_react', { ...args, confirmToken: preview.confirmToken });
    return {
      first,
      second,
      preview,
      data: second.isError ? undefined : (parseToolResult(second) as Record<string, unknown>),
    };
  } finally {
    await h.close();
  }
}

/** A confirm-preview field fenced as third-party text; returns what it fences. */
function fenced(field: unknown): Record<string, unknown> {
  const f = field as Record<string, unknown> | undefined;
  expect(f?.untrusted_content).toBe(true);
  expect(f?.note).toBe(UNTRUSTED_CONTENT_NOTE);
  const { untrusted_content: _u, note: _n, ...rest } = f ?? {};
  return rest;
}

function payload(data: Record<string, unknown> | undefined): Record<string, unknown> {
  expect(data?.untrusted_content).toBe(true);
  const { untrusted_content: _u, note: _n, ...rest } = data ?? {};
  return rest;
}

const ENV_KEYS = ['MCP_CONFIRM_MODE', 'MCP_CONFIRM_TTL_SECONDS', 'MCP_CONFIRM_SECRET'] as const;
let savedEnv: Record<string, string | undefined>;
beforeEach(() => {
  savedEnv = Object.fromEntries(ENV_KEYS.map((k) => [k, process.env[k]]));
  for (const k of ENV_KEYS) delete process.env[k];
});
afterEach(() => {
  for (const k of ENV_KEYS) {
    if (savedEnv[k] === undefined) delete process.env[k];
    else process.env[k] = savedEnv[k];
  }
});

describe('reactionKey', () => {
  it.each([
    ['thumbsup', 'like'],
    ['👍', 'like'],
    ['like', 'like'],
    ['heart', 'heart'],
    ['❤️', 'heart'],
    ['❤', 'heart'],
    ['laugh', 'laugh'],
    ['😆', 'laugh'],
    ['surprised', 'surprised'],
    ['😮', 'surprised'],
    ['sad', 'sad'],
    ['😢', 'sad'],
    ['angry', 'angry'],
    ['😡', 'angry'],
    ['checkmark', CHECKMARK_KEY],
    ['✅', CHECKMARK_KEY],
  ])('maps %s to %s', (input, key) => {
    expect(reactionKey(input)).toBe(key);
  });

  it('defaults the checkmark to the unicode-style key (to be confirmed live)', () => {
    expect(CHECKMARK_KEY).toBe('2705_whiteheavycheckmark');
  });

  it('passes a raw key through, ignoring case, spaces and :colons:', () => {
    expect(reactionKey('fire')).toBe('fire');
    expect(reactionKey('1f62c_grimacingface')).toBe('1f62c_grimacingface');
    expect(reactionKey(' ThumbsUp ')).toBe('like');
    expect(reactionKey(':party:')).toBe('party');
  });

  it('asks for the key of an emoji it has no name for, with the key rule and its code point', () => {
    let err: unknown;
    try {
      reactionKey('😬');
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(McpToolError);
    const e = err as McpToolError;
    expect(e.message).toMatch(/😬/);
    expect(e.hint).toMatch(/1f62c_/);
    expect(e.hint).toMatch(/1f4af_hundredpointssymbol/);
  });

  it('lists every code point of a multi-code-point emoji (a skin tone) instead of guessing a key', () => {
    let err: unknown;
    try {
      reactionKey('👍🏽');
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(McpToolError);
    const e = err as McpToolError;
    expect(e.hint).toMatch(/Its code points are U\+1F44D U\+1F3FD\./);
    expect(e.hint).not.toMatch(/so its key is/);
    expect(e.hint).toMatch(/1f4af_hundredpointssymbol/);
  });

  it('refuses something that is neither an alias, a key nor an emoji', () => {
    expect(() => reactionKey('thumbs up!')).toThrow(McpToolError);
  });
});

describe('teams_react — confirm gate', () => {
  it('phase 1 previews the conversation, message and reaction and writes nothing', async () => {
    const { chatsvc, setReaction } = fakeWith();
    const h = await harnessFor(chatsvc);
    try {
      const first = parseToolResult<PhaseOne>(
        await h.callTool('teams_react', { conversation: GROUP, message_id: MSG, reaction: '👍' }),
      );
      expect(first.confirmToken).toEqual(expect.any(String));
      expect(first.preview?.method).toBe('PUT');
      expect(first.preview?.path).toBe(
        `/users/ME/conversations/${encodeURIComponent(GROUP)}/messages/${MSG}/properties?name=emotions`,
      );
      expect(fenced(first.preview?.conversation)).toEqual({ id: GROUP, title: 'Project Falcon', kind: 'group' });
      expect(fenced(first.preview?.message)).toEqual({ id: MSG, sender: 'Bea Example', preview: 'ship it?' });
      expect(first.preview?.reaction).toBe('like');
      expect(first.preview?.already_reacted).toBe(false);
      expect(setReaction).not.toHaveBeenCalled();
    } finally {
      await h.close();
    }
  });

  it('previews already_reacted: true when the user already has that reaction', async () => {
    const { chatsvc } = fakeWith({ emotions: [{ key: 'like', users: [{ mri: ME, time: T0 }] }] });
    const { preview } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'like' });
    expect(preview?.preview?.already_reacted).toBe(true);
  });

  it('previews already_reacted: false when only someone else has that reaction', async () => {
    const { chatsvc } = fakeWith({ emotions: [{ key: 'like', users: [{ mri: '8:orgid:someone-else', time: T0 }] }] });
    const { preview } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'like' });
    expect(preview?.preview?.already_reacted).toBe(false);
  });

  it('clips the message preview without splitting an emoji in half', async () => {
    const { chatsvc } = fakeWith({
      overrides: { getMessage: vi.fn(async () => rawMsg(Number(MSG), { text: `${'x'.repeat(199)}🎉 tail` })) },
    });
    const { preview } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'like' });
    expect(fenced(preview?.preview?.message).preview).toBe('x'.repeat(199));
  });

  it('phase 2 sets the reaction once and verifies it by reading the message back', async () => {
    const { chatsvc, setReaction } = fakeWith();
    const { data } = await reactConfirmed(chatsvc, { conversation: 'project falcon', message_id: MSG, reaction: 'thumbsup' });
    expect(setReaction).toHaveBeenCalledTimes(1);
    expect(setReaction.mock.calls[0]?.slice(0, 3)).toEqual([GROUP, MSG, 'like']);
    expect(payload(data)).toEqual({
      status: 'reacted',
      conversation: { id: GROUP, title: 'Project Falcon', kind: 'group' },
      message_id: MSG,
      reaction: 'like',
      verified: true,
    });
  });

  it('names an untitled group chat by its participants in the preview and the result', async () => {
    const UNTITLED = '19:untitled0001@thread.v2';
    const { chatsvc } = fakeWith({
      overrides: {
        listConversations: vi.fn().mockResolvedValue([rawConv(UNTITLED, { at: T0 + 4 * MIN, sender: 'Bea Example' })]),
        getMessages: vi.fn().mockResolvedValue([
          rawMsg(T0 + 4 * MIN, { sender: 'Bea Example' }),
          { ...rawMsg(T0 + MIN, { sender: 'Sam One' }), from: 'https://x/contacts/8:orgid:sam' },
        ]),
      },
    });
    const { preview, data } = await reactConfirmed(chatsvc, { conversation: UNTITLED, message_id: MSG, reaction: 'like' });
    const named = { id: UNTITLED, title: 'Group chat: Bea Example; Sam One', kind: 'group', participants: ['Bea Example', 'Sam One'] };
    expect(fenced(preview?.preview?.conversation)).toEqual(named);
    expect(payload(data)).toMatchObject({ status: 'reacted', conversation: named });
  });

  it('remove: DELETEs the reaction and verifies it is gone', async () => {
    const { chatsvc, setReaction, removeReaction } = fakeWith({
      emotions: [{ key: 'heart', users: [{ mri: ME, time: T0 }] }],
    });
    const { preview, data } = await reactConfirmed(chatsvc, {
      conversation: GROUP,
      message_id: MSG,
      reaction: 'heart',
      remove: true,
    });
    expect(preview?.preview?.method).toBe('DELETE');
    expect(setReaction).not.toHaveBeenCalled();
    expect(removeReaction).toHaveBeenCalledWith(GROUP, MSG, 'heart');
    expect(payload(data)).toMatchObject({ status: 'removed', reaction: 'heart', verified: true });
  });

  it('remove: a message gone on reading back (404) is verified: null with a warning, not verified: true', async () => {
    let written = false;
    const removeReaction = vi.fn(async () => {
      written = true;
    });
    const { chatsvc } = fakeWith({
      emotions: [{ key: 'heart', users: [{ mri: ME, time: T0 }] }],
      overrides: {
        removeReaction,
        getMessage: vi.fn(async () =>
          written
            ? undefined
            : { ...rawMsg(Number(MSG)), properties: { emotions: JSON.stringify([{ key: 'heart', users: [{ mri: ME, time: T0 }] }]) } },
        ),
      },
    });
    const { data } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'heart', remove: true });
    expect(removeReaction).toHaveBeenCalledTimes(1);
    expect(payload(data)).toMatchObject({
      status: 'removed',
      verified: null,
      warning: expect.stringMatching(/not found on reading it back/),
    });
  });

  it('add: a message gone on reading back (404) is verified: null too', async () => {
    let written = false;
    const { chatsvc } = fakeWith({
      overrides: {
        setReaction: vi.fn(async () => {
          written = true;
        }),
        getMessage: vi.fn(async () => (written ? undefined : rawMsg(Number(MSG)))),
      },
    });
    const { data } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'like' });
    expect(payload(data)).toMatchObject({ status: 'reacted', verified: null, warning: expect.any(String) });
  });

  it('settles an unconfirmed write (a 503) by the read-back: shown means reacted, with a warning', async () => {
    const { chatsvc, setReaction } = fakeWith();
    const real = setReaction.getMockImplementation()!;
    setReaction.mockImplementation(async (...args) => {
      await real(...args);
      throw new WriteOutcomeUnknownError('Teams', 'PUT', {
        cause: new McpToolError('Teams chat service returned HTTP 503.', { status: 503 }),
      });
    });
    const { second, data } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'like' });
    expect(second.isError).toBeFalsy();
    expect(payload(data)).toMatchObject({
      status: 'reacted',
      verified: true,
      warning: expect.stringMatching(/did not confirm/),
    });
  });

  it('reports status "unknown" when an unconfirmed write does not show on reading back', async () => {
    const { chatsvc } = fakeWith({
      overrides: {
        setReaction: vi.fn().mockRejectedValue(
          new WriteOutcomeUnknownError('Teams', 'PUT', {
            cause: new McpToolError('Teams chat service returned HTTP 504.', { status: 504 }),
          }),
        ),
      },
    });
    const { second, data } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'like' });
    expect(second.isError).toBeFalsy();
    expect(payload(data)).toMatchObject({
      status: 'unknown',
      verified: false,
      warning: expect.stringMatching(/HTTP 504.*check the message/),
    });
  });

  it('reports a timed-out remove as status "unknown" when the reaction is still there', async () => {
    const { chatsvc } = fakeWith({
      emotions: [{ key: 'heart', users: [{ mri: ME, time: T0 }] }],
      overrides: {
        removeReaction: vi.fn().mockRejectedValue(
          new WriteOutcomeUnknownError('Teams', 'DELETE', { timeoutMs: 60_000, cause: new RequestTimeoutError('Teams', 60_000) }),
        ),
      },
    });
    const { data } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'heart', remove: true });
    expect(payload(data)).toMatchObject({ status: 'unknown', warning: expect.stringMatching(/did not answer in time/) });
  });

  it('still fails plainly on a 4xx refusal', async () => {
    const { chatsvc } = fakeWith({
      overrides: { setReaction: vi.fn().mockRejectedValue(new McpToolError('Teams chat service returned HTTP 403.', { status: 403 })) },
    });
    const { second } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'like' });
    expect(second.isError).toBe(true);
  });

  it('reports verified: false with a warning when the read-back does not show the change', async () => {
    const { chatsvc } = fakeWith({ ignore: true });
    const { data } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'like' });
    expect(payload(data)).toMatchObject({ status: 'reacted', verified: false, warning: expect.stringMatching(/check/i) });
  });

  it('still reports the reaction when the read-back itself fails', async () => {
    // Reads before the write succeed; the read-back after it fails.
    let written = false;
    const setReaction = vi.fn(async () => {
      written = true;
    });
    const { chatsvc } = fakeWith({
      overrides: {
        setReaction,
        getMessage: vi.fn(async () => {
          if (!written) return rawMsg(Number(MSG), { text: 'ship it?' });
          throw new Error('boom');
        }),
      },
    });
    const { data } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: 'like' });
    expect(setReaction).toHaveBeenCalledTimes(1);
    expect(payload(data)).toMatchObject({
      status: 'reacted',
      verified: false,
      warning: expect.stringMatching(/reading the message back failed/),
    });
  });

  it('declares itself a non-destructive, idempotent, open-world write', async () => {
    const h = await harnessFor(fakeWith().chatsvc);
    try {
      const tool = (await h.client.listTools()).tools.find((t) => t.name === 'teams_react')!;
      expect(tool.annotations).toMatchObject({
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      });
      expect(tool.description).toMatch(/confirmToken/);
      expect(tool.inputSchema.properties).toHaveProperty('confirmToken');
    } finally {
      await h.close();
    }
  });
});

describe('teams_react — refusals before any preview', () => {
  it('refuses a message that is not in the conversation', async () => {
    const { chatsvc, setReaction } = fakeWith();
    const { first } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: '999', reaction: 'like' });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(/999/);
    expect(setReaction).not.toHaveBeenCalled();
  });

  it('refuses an ambiguous conversation name', async () => {
    const { chatsvc } = fakeWith();
    const { first } = await reactConfirmed(chatsvc, { conversation: 'falcon', message_id: MSG, reaction: 'like' });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toContain(OTHER_GROUP);
  });

  it('refuses an unnamed emoji, asking for its key', async () => {
    const { chatsvc, getMessage } = fakeWith();
    const { first } = await reactConfirmed(chatsvc, { conversation: GROUP, message_id: MSG, reaction: '😬' });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(/1f62c_/);
    expect(getMessage).not.toHaveBeenCalled();
  });

  it('resolves the self-chat by the name "notes"', async () => {
    const { chatsvc, getMessage } = fakeWith();
    await reactConfirmed(chatsvc, { conversation: 'notes', message_id: MSG, reaction: 'like' });
    expect(getMessage.mock.calls[0]?.[0]).toBe('48:notes');
  });

  it('refuses a 48: system conversation other than the self-chat', async () => {
    const { chatsvc } = fakeWith();
    const { first } = await reactConfirmed(chatsvc, { conversation: '48:calendar', message_id: MSG, reaction: 'like' });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(/48:notes/);
  });
});
