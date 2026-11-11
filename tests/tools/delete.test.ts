import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import { registerDeleteTools } from '../../src/tools/delete.js';
import { UNTRUSTED_CONTENT_NOTE } from '../../src/tools/untrusted.js';
import type { ChatServiceClient, RawMessage } from '../../src/chatsvc.js';
import { fakeChatsvc, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
const GROUP = '19:group0001@thread.v2';
const CONVS = [rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * MIN })];

const MINE = String(T0 + MIN);
const THEIRS = String(T0 + 2 * MIN);
const GONE = String(T0 + 3 * MIN);
const EVENT = String(T0 + 4 * MIN);

function history(): RawMessage[] {
  return [
    rawMsg(T0 + MIN, { text: '<p>oops, wrong chat</p>', fromMe: true }),
    rawMsg(T0 + 2 * MIN, { text: '<p>question</p>', sender: 'Bea Example' }),
    rawMsg(T0 + 3 * MIN, { text: '', fromMe: true, deleted: true }),
    rawMsg(T0 + 4 * MIN, { text: '<p>call ended</p>', fromMe: true, type: 'Event/Call' }),
  ];
}

/** A chat service whose deletes mark the stored message deleted, unless `land` is false. */
function fakeWith(opts: { fail?: Error; land?: boolean; vanish?: boolean } = {}) {
  const stored = history();
  const deleteMessage = vi.fn(async (_conv: string, id: string) => {
    if (opts.land ?? opts.fail === undefined) {
      const i = stored.findIndex((x) => x.id === id);
      if (opts.vanish) stored.splice(i, 1);
      else stored[i] = { ...stored[i], content: '', properties: { deletetime: String(T0 + 9 * MIN) } };
    }
    if (opts.fail) throw opts.fail;
  });
  const chatsvc = fakeChatsvc(CONVS, {
    deleteMessage,
    getMessage: vi.fn(async (_conv: string, id: string) => {
      const m = stored.find((x) => x.id === id);
      return m === undefined ? undefined : structuredClone(m);
    }),
  });
  return { chatsvc, deleteMessage };
}

type PhaseOne = { confirmToken?: string; preview?: Record<string, any> };

async function del(chatsvc: ChatServiceClient, args: Record<string, unknown>) {
  const h = await createTestHarness((server) => registerDeleteTools(server, { chatsvc }));
  try {
    const first = await h.callTool('teams_delete_message', args);
    if (first.isError) return { first, error: (first.content?.[0] as { text?: string })?.text ?? '' };
    const one = parseToolResult<PhaseOne & Record<string, any>>(first);
    if (one.confirmToken === undefined) return { first, data: one };
    const second = await h.callTool('teams_delete_message', { ...args, confirmToken: one.confirmToken });
    return { first, preview: one.preview, data: parseToolResult(second) as Record<string, any> };
  } finally {
    await h.close();
  }
}

function fenced(field: unknown): unknown {
  const f = field as Record<string, unknown> | undefined;
  expect(f?.untrusted_content).toBe(true);
  expect(f?.note).toBe(UNTRUSTED_CONTENT_NOTE);
  const { untrusted_content: _u, note: _n, ...rest } = f ?? {};
  return 'data' in rest ? rest.data : rest;
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

describe('teams_delete_message', () => {
  it('previews the message to be deleted and deletes nothing in phase 1', async () => {
    const { chatsvc, deleteMessage } = fakeWith();
    const h = await createTestHarness((server) => registerDeleteTools(server, { chatsvc }));
    try {
      const one = parseToolResult<PhaseOne>(await h.callTool('teams_delete_message', { conversation: GROUP, message_id: MINE }));
      expect(one.confirmToken).toEqual(expect.any(String));
      expect(one.preview).toMatchObject({
        method: 'DELETE',
        path: `/users/ME/conversations/${encodeURIComponent(GROUP)}/messages/${MINE}`,
        message_id: MINE,
        text: 'oops, wrong chat',
        time: new Date(T0 + MIN).toISOString(),
      });
      expect(fenced(one.preview?.conversation)).toEqual({ id: GROUP, title: 'Project Falcon', kind: 'group' });
      expect(deleteMessage).not.toHaveBeenCalled();
    } finally {
      await h.close();
    }
  });

  it('deletes once confirmed and verifies it by reading the message back', async () => {
    const { chatsvc, deleteMessage } = fakeWith();
    const { data } = await del(chatsvc, { conversation: GROUP, message_id: MINE });
    expect(deleteMessage).toHaveBeenCalledExactlyOnceWith(GROUP, MINE);
    expect(data).toMatchObject({ status: 'deleted', message_id: MINE, verified: true });
  });

  it('counts a message that is gone altogether (404) as deleted', async () => {
    const { chatsvc } = fakeWith({ vanish: true });
    const { data } = await del(chatsvc, { conversation: GROUP, message_id: MINE });
    expect(data).toMatchObject({ status: 'deleted', verified: true });
  });

  it('reports an already-deleted message without asking or deleting', async () => {
    const { chatsvc, deleteMessage } = fakeWith();
    const { data } = await del(chatsvc, { conversation: GROUP, message_id: GONE });
    expect(data).toMatchObject({ status: 'already_deleted', message_id: GONE });
    expect(deleteMessage).not.toHaveBeenCalled();
  });

  it.each([
    [THEIRS, /only your own messages/i],
    ['1', /no message 1/i],
    [EVENT, /not a message anyone wrote/i],
  ])('refuses message %s before any preview', async (id, why) => {
    const { chatsvc, deleteMessage } = fakeWith();
    const { error, first } = await del(chatsvc, { conversation: GROUP, message_id: id });
    expect(first.isError).toBe(true);
    expect(error).toMatch(why);
    expect(deleteMessage).not.toHaveBeenCalled();
  });

  it('settles an unconfirmed delete by reading back: deleted with a warning, or unknown', async () => {
    const lost = () => new WriteOutcomeUnknownError('Teams', 'DELETE', { hint: 'x' });
    const landed = await del(fakeWith({ fail: lost(), land: true }).chatsvc, { conversation: GROUP, message_id: MINE });
    expect(landed.data).toMatchObject({ status: 'deleted', verified: true, warning: expect.stringMatching(/did not confirm/) });
    const missed = await del(fakeWith({ fail: lost(), land: false }).chatsvc, { conversation: GROUP, message_id: MINE });
    expect(missed.data).toMatchObject({ status: 'unknown', verified: false });
  });

  it('is declared a destructive, idempotent write', async () => {
    const { chatsvc } = fakeWith();
    const h = await createTestHarness((server) => registerDeleteTools(server, { chatsvc }));
    try {
      const tool = (await h.client.listTools()).tools.find((t) => t.name === 'teams_delete_message');
      expect(tool?.annotations).toMatchObject({ readOnlyHint: false, destructiveHint: true, idempotentHint: true });
    } finally {
      await h.close();
    }
  });
});
