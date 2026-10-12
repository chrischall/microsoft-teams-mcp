import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { RequestTimeoutError, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import { newestMessageId, registerMarkReadTools } from '../../src/tools/markread.js';
import { UNTRUSTED_CONTENT_NOTE } from '../../src/tools/untrusted.js';
import type { ChatServiceClient, RawMessage } from '../../src/chatsvc.js';
import { fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;

const GROUP = '19:group0001@thread.v2';
const OTHER_GROUP = '19:group0002@thread.v2';
const READ_GROUP = '19:group0003@thread.v2';

/** GROUP is read up to +2 min; its newest message is at +4 min. */
const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * MIN, horizon: T0 + 2 * MIN }),
  rawConv(OTHER_GROUP, { topic: 'Falcon launch party', at: T0 }),
  rawConv(READ_GROUP, { topic: 'Weekly sync', at: T0 + 4 * MIN, horizon: T0 + 4 * MIN }),
];

const MSGS = [
  rawMsg(T0 + 1 * MIN, { text: 'first' }),
  rawMsg(T0 + 2 * MIN, { text: 'second' }),
  rawMsg(T0 + 3 * MIN, { text: 'third', sender: 'Bea Example' }),
  rawMsg(T0 + 4 * MIN, { text: 'fourth', sender: 'Bea Example' }),
];

const NEWEST = String(T0 + 4 * MIN);

function chatsvcWith(overrides: Record<string, unknown> = {}, messages: RawMessage[] = MSGS) {
  const setConsumptionHorizon = vi.fn(async () => undefined);
  const getMessage = vi.fn(async (_c: string, id: string) => messages.find((m) => m.id === id));
  const chatsvc = fakeChatsvc(CONVS, {
    getMessages: messagesFrom(messages),
    getMessage,
    setConsumptionHorizon,
    ...overrides,
  });
  return { chatsvc, setConsumptionHorizon: (overrides.setConsumptionHorizon ?? setConsumptionHorizon) as typeof setConsumptionHorizon };
}

async function harnessFor(chatsvc: ChatServiceClient) {
  return createTestHarness((server) => registerMarkReadTools(server, { chatsvc }));
}

type PhaseOne = {
  confirmToken?: string;
  preview?: { method?: string; path?: string } & Record<string, unknown>;
};

function errorText(result: { content?: unknown[] }): string {
  return (result.content?.[0] as { text?: string })?.text ?? '';
}

/** Phase 1, then (when phase 1 asked for one) phase 2 with its token. */
async function markConfirmed(chatsvc: ChatServiceClient, args: Record<string, unknown>) {
  const h = await harnessFor(chatsvc);
  try {
    const first = await h.callTool('teams_mark_read', args);
    if (first.isError) return { first, second: first, preview: undefined, data: undefined };
    const preview = parseToolResult<PhaseOne>(first);
    if (preview.confirmToken === undefined) {
      return { first, second: first, preview: undefined, data: parseToolResult(first) as Record<string, unknown> };
    }
    const second = await h.callTool('teams_mark_read', { ...args, confirmToken: preview.confirmToken });
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

describe('newestMessageId', () => {
  it('takes the largest id, since a page is not strictly ordered', () => {
    expect(newestMessageId([rawMsg(T0 + 2), rawMsg(T0 + 5), rawMsg(T0 + 3)])).toBe(String(T0 + 5));
  });

  it('is undefined for no messages', () => {
    expect(newestMessageId([])).toBeUndefined();
  });
});

describe('teams_mark_read — confirm gate', () => {
  it('phase 1 previews the conversation, the message and the marker move, and writes nothing', async () => {
    const { chatsvc, setConsumptionHorizon } = chatsvcWith();
    const h = await harnessFor(chatsvc);
    try {
      const first = parseToolResult<PhaseOne>(await h.callTool('teams_mark_read', { conversation: GROUP }));
      expect(first.confirmToken).toEqual(expect.any(String));
      expect(first.preview?.method).toBe('PUT');
      expect(first.preview?.path).toBe(
        `/users/ME/conversations/${encodeURIComponent(GROUP)}/properties?name=consumptionhorizon`,
      );
      expect(fenced(first.preview?.conversation)).toEqual({ id: GROUP, title: 'Project Falcon', kind: 'group' });
      expect(fenced(first.preview?.up_to_message)).toEqual({ id: NEWEST, sender: 'Bea Example', preview: 'fourth' });
      expect(first.preview?.previous_horizon_message_id).toBe(String(T0 + 2 * MIN));
      expect(setConsumptionHorizon).not.toHaveBeenCalled();
    } finally {
      await h.close();
    }
  });

  it('phase 2 moves the marker to the newest message and reports both horizons', async () => {
    const { chatsvc, setConsumptionHorizon } = chatsvcWith();
    const { data } = await markConfirmed(chatsvc, { conversation: GROUP });
    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(GROUP, NEWEST);
    expect(payload(data)).toEqual({
      status: 'marked_read',
      conversation: expect.objectContaining({ id: GROUP, title: 'Project Falcon', unread: false }),
      previous_horizon_message_id: String(T0 + 2 * MIN),
      new_horizon_message_id: NEWEST,
    });
  });

  it('reads one page for the newest message', async () => {
    const { chatsvc } = chatsvcWith();
    await markConfirmed(chatsvc, { conversation: GROUP });
    expect(chatsvc.getMessages).toHaveBeenCalledWith(GROUP, expect.objectContaining({ max: expect.any(Number) }));
    const opts = vi.mocked(chatsvc.getMessages).mock.calls[0][1];
    expect(opts.max).toBeLessThanOrEqual(200);
  });

  it('marks up to an explicit up_to_message_id', async () => {
    const { chatsvc, setConsumptionHorizon } = chatsvcWith();
    const target = String(T0 + 3 * MIN);
    const { data } = await markConfirmed(chatsvc, { conversation: GROUP, up_to_message_id: target });
    expect(setConsumptionHorizon).toHaveBeenCalledExactlyOnceWith(GROUP, target);
    expect(payload(data)).toMatchObject({ status: 'marked_read', new_horizon_message_id: target });
    // Something newer is still unread, so the summary does not claim otherwise.
    expect((payload(data).conversation as { unread: boolean }).unread).toBe(true);
  });

  it('refuses an up_to_message_id that is not in the conversation', async () => {
    const { chatsvc, setConsumptionHorizon } = chatsvcWith();
    const { first } = await markConfirmed(chatsvc, { conversation: GROUP, up_to_message_id: '123' });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(/no message 123/i);
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
  });

  it('resolves a conversation by a unique name, and the self-chat by "notes"', async () => {
    const { chatsvc, setConsumptionHorizon } = chatsvcWith();
    await markConfirmed(chatsvc, { conversation: 'project falcon' });
    expect(setConsumptionHorizon).toHaveBeenLastCalledWith(GROUP, NEWEST);
    await markConfirmed(chatsvc, { conversation: 'notes' });
    expect(setConsumptionHorizon).toHaveBeenLastCalledWith('48:notes', NEWEST);
  });

  it('refuses an ambiguous name, naming the candidates', async () => {
    const { chatsvc } = chatsvcWith();
    const { first } = await markConfirmed(chatsvc, { conversation: 'falcon' });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(/matches 2 conversations/);
  });
});

describe('teams_mark_read — never backwards', () => {
  it('is a no-op (no gate, no write) when the marker is already at the newest message', async () => {
    const { chatsvc, setConsumptionHorizon } = chatsvcWith();
    const { data, preview } = await markConfirmed(chatsvc, { conversation: READ_GROUP });
    expect(preview).toBeUndefined();
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(payload(data)).toMatchObject({
      status: 'already_read',
      previous_horizon_message_id: NEWEST,
      new_horizon_message_id: NEWEST,
    });
    expect(String(payload(data).hint)).toMatch(/already/i);
  });

  it('is a no-op when up_to_message_id is older than the marker', async () => {
    const { chatsvc, setConsumptionHorizon } = chatsvcWith();
    const older = String(T0 + 1 * MIN);
    const { data } = await markConfirmed(chatsvc, { conversation: GROUP, up_to_message_id: older });
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(payload(data)).toMatchObject({
      status: 'already_read',
      previous_horizon_message_id: String(T0 + 2 * MIN),
      new_horizon_message_id: String(T0 + 2 * MIN),
    });
  });

  it('is a no-op for a conversation with no messages', async () => {
    const { chatsvc, setConsumptionHorizon } = chatsvcWith({}, []);
    const { data } = await markConfirmed(chatsvc, { conversation: OTHER_GROUP });
    expect(setConsumptionHorizon).not.toHaveBeenCalled();
    expect(payload(data)).toMatchObject({ status: 'already_read' });
  });

  it('reports previous_horizon_message_id null for a conversation never opened', async () => {
    const neverOpened = { ...rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * MIN }), properties: {} };
    const { chatsvc, setConsumptionHorizon } = chatsvcWith({ listConversations: vi.fn().mockResolvedValue([neverOpened]) });
    const { data } = await markConfirmed(chatsvc, { conversation: GROUP });
    expect(setConsumptionHorizon).toHaveBeenCalledWith(GROUP, NEWEST);
    expect(payload(data)).toMatchObject({ previous_horizon_message_id: null, new_horizon_message_id: NEWEST });
  });
});

describe('teams_mark_read — an unconfirmed write', () => {
  it('reports status unknown with a warning, not a failure', async () => {
    const setConsumptionHorizon = vi.fn(async () => {
      throw new WriteOutcomeUnknownError('Teams', 'PUT', { cause: new RequestTimeoutError('Teams', 60_000), timeoutMs: 60_000 });
    });
    const { chatsvc } = chatsvcWith({ setConsumptionHorizon });
    const { data, second } = await markConfirmed(chatsvc, { conversation: GROUP });
    expect(second.isError).toBeFalsy();
    expect(payload(data)).toMatchObject({ status: 'unknown', new_horizon_message_id: NEWEST });
    expect(String(payload(data).warning)).toMatch(/did not confirm/i);
  });

  it('passes a refusal through as an error', async () => {
    const setConsumptionHorizon = vi.fn(async () => {
      throw Object.assign(new Error('Teams chat service returned HTTP 403.'), { status: 403 });
    });
    const { chatsvc } = chatsvcWith({ setConsumptionHorizon });
    const { second } = await markConfirmed(chatsvc, { conversation: GROUP });
    expect(second.isError).toBe(true);
  });
});
