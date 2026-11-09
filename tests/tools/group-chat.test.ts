import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import { registerGroupChatTools } from '../../src/tools/group-chat.js';
import { UNTRUSTED_CONTENT_NOTE } from '../../src/tools/untrusted.js';
import type { ChatServiceClient, RawConversation, RawThread } from '../../src/chatsvc.js';
import { ME, fakeChatsvc, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const ME_OID = ME.slice('8:orgid:'.length);
const GROUP = '19:group0001@thread.v2';
const GREG_OID = '00000000-0000-0000-0000-0000000000c3';
const GREG = `8:orgid:${GREG_OID}`;
const SAM_OID = '00000000-0000-0000-0000-0000000000d4';
const SAM = `8:orgid:${SAM_OID}`;
const BEA = '8:orgid:00000000-0000-0000-0000-0000000000e5';
const ONE_ON_ONE = `19:${ME_OID}_${GREG_OID}@unq.gbl.spaces`;
const MEETING = '19:meeting_abc@thread.v2';
const CHANNEL = '19:chan0001@thread.tacv2';
const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;
const ENC = encodeURIComponent(GROUP);

const CONVS: RawConversation[] = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * 60_000 }),
  { ...rawConv(ONE_ON_ONE, { at: T0 }), lastMessage: { ...rawConv(ONE_ON_ONE).lastMessage, from: FROM(GREG), imdisplayname: 'King, Greg' } },
  rawConv(MEETING, { topic: 'Weekly sync', at: T0 }),
  rawConv(CHANNEL, { topic: 'Releases', at: T0 }),
];

/**
 * A chat service whose GROUP thread really changes under setTopic / addMember
 * / removeMember — unless `ignore` is set (accepted, but not shown on reading back).
 */
function fakeWith(opts: { members?: string[]; topic?: string; ignore?: boolean; overrides?: Record<string, unknown> } = {}) {
  let topic = opts.topic ?? 'Project Falcon';
  let members = opts.members ?? [ME, BEA];
  const thread = (): RawThread => ({
    id: GROUP,
    properties: { topic },
    members: members.map((id) => ({ id, role: id === ME ? 'Admin' : 'User' })),
  });
  const getThread = vi.fn(async () => thread());
  const setTopic = vi.fn(async (_id: string, t: string) => {
    if (!opts.ignore) topic = t;
  });
  const addMember = vi.fn(async (_id: string, mri: string) => {
    if (!opts.ignore && !members.includes(mri)) members = [...members, mri];
  });
  const removeMember = vi.fn(async (_id: string, mri: string) => {
    if (!opts.ignore) members = members.filter((m) => m !== mri);
  });
  const getMessages = vi.fn(async (id: string) =>
    id === GROUP ? [{ ...rawMsg(T0, { sender: 'Bea Example' }), from: FROM(BEA) }] : [],
  );
  const chatsvc = fakeChatsvc(CONVS, { getThread, setTopic, addMember, removeMember, getMessages, ...opts.overrides });
  return { chatsvc, getThread, setTopic, addMember, removeMember };
}

type PhaseOne = {
  confirmToken?: string;
  preview?: { method?: string; path?: string; willSend?: unknown } & Record<string, unknown>;
};

function errorText(result: { content?: unknown[] }): string {
  return (result.content?.[0] as { text?: string })?.text ?? '';
}

/** Phase 1 then (when it minted a token) phase 2 with it. */
async function confirmed(chatsvc: ChatServiceClient, tool: string, args: Record<string, unknown>) {
  const h = await createTestHarness((server) => registerGroupChatTools(server, { chatsvc }));
  try {
    const first = await h.callTool(tool, args);
    if (first.isError) return { first, second: first, preview: undefined, data: undefined };
    const preview = parseToolResult<PhaseOne>(first);
    if (preview.confirmToken === undefined) {
      return { first, second: first, preview, data: preview as Record<string, unknown> };
    }
    const second = await h.callTool(tool, { ...args, confirmToken: preview.confirmToken });
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

async function phaseOne(chatsvc: ChatServiceClient, tool: string, args: Record<string, unknown>) {
  const h = await createTestHarness((server) => registerGroupChatTools(server, { chatsvc }));
  try {
    const first = await h.callTool(tool, args);
    return { first, preview: first.isError ? undefined : parseToolResult<PhaseOne>(first) };
  } finally {
    await h.close();
  }
}

function fenced(field: unknown): unknown {
  const f = field as Record<string, unknown> | undefined;
  expect(f?.untrusted_content).toBe(true);
  expect(f?.note).toBe(UNTRUSTED_CONTENT_NOTE);
  const { untrusted_content: _u, note: _n, ...rest } = f ?? {};
  return 'data' in rest && Object.keys(rest).length === 1 ? rest.data : rest;
}

function payload(data: Record<string, unknown> | undefined): Record<string, unknown> {
  expect(data?.untrusted_content).toBe(true);
  const { untrusted_content: _u, note: _n, ...rest } = data ?? {};
  return rest;
}

const unknownOutcome = (method: string) =>
  new WriteOutcomeUnknownError('Teams', method, {
    cause: new McpToolError('Teams chat service returned HTTP 504.', { status: 504 }),
  });

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

const NON_GROUPS = [
  ['a 1:1', ONE_ON_ONE, /1:1/],
  ['a meeting chat', MEETING, /meeting/],
  ['a channel', CHANNEL, /channel/],
  ['the self-chat', 'notes', /self-chat/],
] as const;

describe('teams_rename_chat', () => {
  it('phase 1 previews the current and new name and writes nothing', async () => {
    const { chatsvc, setTopic } = fakeWith();
    const { preview } = await phaseOne(chatsvc, 'teams_rename_chat', { conversation: GROUP, name: '  Falcon v2  ' });
    expect(preview?.confirmToken).toEqual(expect.any(String));
    expect(preview?.preview?.method).toBe('PUT');
    expect(preview?.preview?.path).toBe(`/threads/${ENC}/properties?name=topic`);
    expect(preview?.preview?.willSend).toEqual({ topic: 'Falcon v2' });
    expect(fenced(preview?.preview?.conversation)).toEqual({ id: GROUP, title: 'Project Falcon', kind: 'group' });
    expect(fenced(preview?.preview?.current_name)).toBe('Project Falcon');
    expect(preview?.preview?.new_name).toBe('Falcon v2');
    expect(setTopic).not.toHaveBeenCalled();
  });

  it('phase 2 sets the trimmed topic and verifies it by reading the chat back', async () => {
    const { chatsvc, setTopic } = fakeWith();
    const { data } = await confirmed(chatsvc, 'teams_rename_chat', { conversation: 'project falcon', name: ' Falcon v2 ' });
    expect(setTopic).toHaveBeenCalledTimes(1);
    expect(setTopic).toHaveBeenCalledWith(GROUP, 'Falcon v2');
    expect(payload(data)).toEqual({
      status: 'renamed',
      conversation: { id: GROUP, title: 'Project Falcon', kind: 'group' },
      previous_name: 'Project Falcon',
      name: 'Falcon v2',
      verified: true,
    });
  });

  it('writes nothing, and asks for no confirmation, when the chat already has that name', async () => {
    const { chatsvc, setTopic } = fakeWith();
    const { preview, data } = await confirmed(chatsvc, 'teams_rename_chat', { conversation: GROUP, name: 'Project Falcon' });
    expect(preview?.confirmToken).toBeUndefined();
    expect(payload(data)).toMatchObject({ status: 'unchanged', name: 'Project Falcon' });
    expect(setTopic).not.toHaveBeenCalled();
  });

  it('refuses an empty name', async () => {
    const { chatsvc, setTopic } = fakeWith();
    const { first } = await phaseOne(chatsvc, 'teams_rename_chat', { conversation: GROUP, name: '   ' });
    expect(first.isError).toBe(true);
    expect(setTopic).not.toHaveBeenCalled();
  });

  it.each(NON_GROUPS)('refuses %s', async (_label, conversation, why) => {
    const { chatsvc, getThread, setTopic } = fakeWith();
    const { first } = await phaseOne(chatsvc, 'teams_rename_chat', { conversation, name: 'x' });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(why);
    expect(errorText(first)).toMatch(/group chat/);
    expect(getThread).not.toHaveBeenCalled();
    expect(setTopic).not.toHaveBeenCalled();
  });

  it('settles an unconfirmed write by the read-back: shown means renamed, with a warning', async () => {
    const { chatsvc, setTopic } = fakeWith();
    const real = setTopic.getMockImplementation()!;
    setTopic.mockImplementation(async (...args) => {
      await real(...args);
      throw unknownOutcome('PUT');
    });
    const { data } = await confirmed(chatsvc, 'teams_rename_chat', { conversation: GROUP, name: 'Falcon v2' });
    expect(payload(data)).toMatchObject({ status: 'renamed', verified: true, warning: expect.stringMatching(/did not confirm/) });
  });

  it('reports status "unknown" when an unconfirmed write does not show on reading back', async () => {
    const { chatsvc } = fakeWith({ overrides: { setTopic: vi.fn().mockRejectedValue(unknownOutcome('PUT')) } });
    const { second, data } = await confirmed(chatsvc, 'teams_rename_chat', { conversation: GROUP, name: 'Falcon v2' });
    expect(second.isError).toBeFalsy();
    expect(payload(data)).toMatchObject({ status: 'unknown', verified: false, warning: expect.stringMatching(/HTTP 504/) });
  });

  it('reports verified: false with a warning when Teams accepted it but the read-back shows the old name', async () => {
    const { chatsvc } = fakeWith({ ignore: true });
    const { data } = await confirmed(chatsvc, 'teams_rename_chat', { conversation: GROUP, name: 'Falcon v2' });
    expect(payload(data)).toMatchObject({ status: 'renamed', verified: false, warning: expect.any(String) });
  });

  it('still fails plainly on a 4xx refusal', async () => {
    const { chatsvc } = fakeWith({
      overrides: { setTopic: vi.fn().mockRejectedValue(new McpToolError('Teams chat service returned HTTP 403.', { status: 403 })) },
    });
    const { second } = await confirmed(chatsvc, 'teams_rename_chat', { conversation: GROUP, name: 'Falcon v2' });
    expect(second.isError).toBe(true);
  });
});

describe('teams_add_chat_members', () => {
  it('phase 1 resolves every person, previews who is added and who is already in, and writes nothing', async () => {
    const { chatsvc, addMember } = fakeWith();
    const { preview } = await phaseOne(chatsvc, 'teams_add_chat_members', {
      conversation: GROUP,
      people: ['Greg King', SAM_OID, 'bea'],
    });
    expect(preview?.confirmToken).toEqual(expect.any(String));
    expect(preview?.preview?.method).toBe('PUT');
    expect(preview?.preview?.path).toBe(`/threads/${ENC}/members`);
    expect(fenced(preview?.preview?.conversation)).toEqual({ id: GROUP, title: 'Project Falcon', kind: 'group' });
    expect(fenced(preview?.preview?.add)).toEqual([
      { person: 'Greg King', mri: GREG, name: 'King, Greg' },
      { person: SAM_OID, mri: SAM },
    ]);
    expect(fenced(preview?.preview?.already_members)).toEqual([{ person: 'bea', mri: BEA, name: 'Bea Example' }]);
    expect(addMember).not.toHaveBeenCalled();
  });

  it('phase 2 adds each new person as a User, skips members, and verifies by reading the chat back', async () => {
    const { chatsvc, addMember } = fakeWith();
    const { data } = await confirmed(chatsvc, 'teams_add_chat_members', {
      conversation: GROUP,
      people: ['King, Greg', SAM, 'Bea Example'],
    });
    expect(addMember.mock.calls).toEqual([
      [GROUP, GREG, 'User'],
      [GROUP, SAM, 'User'],
    ]);
    expect(payload(data)).toEqual({
      status: 'done',
      conversation: { id: GROUP, title: 'Project Falcon', kind: 'group' },
      results: [
        { person: 'King, Greg', mri: GREG, name: 'King, Greg', status: 'added', verified: true },
        { person: SAM, mri: SAM, status: 'added', verified: true },
        { person: 'Bea Example', mri: BEA, name: 'Bea Example', status: 'already_member' },
      ],
    });
  });

  it('returns at once, with no confirmation, when everyone is already a member', async () => {
    const { chatsvc, addMember } = fakeWith();
    const { preview, data } = await confirmed(chatsvc, 'teams_add_chat_members', { conversation: GROUP, people: [BEA] });
    expect(preview?.confirmToken).toBeUndefined();
    expect(payload(data)).toMatchObject({
      status: 'unchanged',
      results: [{ person: BEA, mri: BEA, name: 'Bea Example', status: 'already_member' }],
    });
    expect(addMember).not.toHaveBeenCalled();
  });

  it('refuses the whole call when one name is ambiguous, before any preview', async () => {
    const { chatsvc, addMember } = fakeWith({
      overrides: {
        getMessages: vi.fn(async () => [
          { ...rawMsg(T0, { sender: 'Greg Kingsley' }), from: FROM('8:orgid:gk2') },
        ]),
      },
    });
    const { first } = await phaseOne(chatsvc, 'teams_add_chat_members', { conversation: GROUP, people: [SAM, 'greg'] });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toContain('"King, Greg"');
    expect(errorText(first)).toContain('"Greg Kingsley"');
    expect(addMember).not.toHaveBeenCalled();
  });

  it('refuses an unknown name, asking for the MRI', async () => {
    const { chatsvc, addMember } = fakeWith();
    const { first } = await phaseOne(chatsvc, 'teams_add_chat_members', { conversation: GROUP, people: ['Zed Nobody'] });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(/8:orgid:/);
    expect(addMember).not.toHaveBeenCalled();
  });

  it.each(NON_GROUPS)('refuses %s', async (_label, conversation, why) => {
    const { chatsvc, addMember } = fakeWith();
    const { first } = await phaseOne(chatsvc, 'teams_add_chat_members', { conversation, people: [SAM] });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(why);
    expect(addMember).not.toHaveBeenCalled();
  });

  it('settles unconfirmed adds by the read-back: shown is added, not shown is unknown', async () => {
    const { chatsvc, addMember } = fakeWith();
    const real = addMember.getMockImplementation()!;
    addMember.mockImplementation(async (id: string, mri: string) => {
      if (mri === GREG) await real(id, mri);
      throw unknownOutcome('PUT');
    });
    const { data } = await confirmed(chatsvc, 'teams_add_chat_members', { conversation: GROUP, people: [GREG, SAM] });
    const p = payload(data);
    expect(p.results).toEqual([
      { person: GREG, mri: GREG, name: 'King, Greg', status: 'added', verified: true },
      { person: SAM, mri: SAM, status: 'unknown', verified: false },
    ]);
    expect(p.warning).toMatch(/check the chat/i);
  });

  it('reports a refused add as failed with its error and still adds the rest', async () => {
    const { chatsvc, addMember } = fakeWith();
    const real = addMember.getMockImplementation()!;
    addMember.mockImplementation(async (id: string, mri: string) => {
      if (mri === GREG) throw new McpToolError('Teams chat service returned HTTP 403.', { status: 403 });
      await real(id, mri);
    });
    const { data } = await confirmed(chatsvc, 'teams_add_chat_members', { conversation: GROUP, people: [GREG, SAM] });
    expect(payload(data).results).toEqual([
      { person: GREG, mri: GREG, name: 'King, Greg', status: 'failed', error: 'Teams chat service returned HTTP 403.' },
      { person: SAM, mri: SAM, status: 'added', verified: true },
    ]);
  });

  it('refuses more than 20 people or none', async () => {
    const { chatsvc } = fakeWith();
    expect((await phaseOne(chatsvc, 'teams_add_chat_members', { conversation: GROUP, people: [] })).first.isError).toBe(true);
    const many = Array.from({ length: 21 }, (_, i) => `8:orgid:p${i}`);
    expect((await phaseOne(chatsvc, 'teams_add_chat_members', { conversation: GROUP, people: many })).first.isError).toBe(true);
  });
});

describe('teams_remove_chat_members', () => {
  it('phase 1 previews who is removed and writes nothing', async () => {
    const { chatsvc, removeMember } = fakeWith({ members: [ME, BEA, GREG] });
    const { preview } = await phaseOne(chatsvc, 'teams_remove_chat_members', { conversation: GROUP, people: ['Example, Bea'] });
    expect(preview?.confirmToken).toEqual(expect.any(String));
    expect(preview?.preview?.method).toBe('DELETE');
    expect(preview?.preview?.path).toBe(`/threads/${ENC}/members`);
    expect(fenced(preview?.preview?.remove)).toEqual([{ person: 'Example, Bea', mri: BEA, name: 'Bea Example' }]);
    expect(removeMember).not.toHaveBeenCalled();
  });

  it('phase 2 removes each member, skips non-members, and verifies by reading the chat back', async () => {
    const { chatsvc, removeMember } = fakeWith({ members: [ME, BEA, GREG] });
    const { data } = await confirmed(chatsvc, 'teams_remove_chat_members', {
      conversation: GROUP,
      people: ['Example, Bea', 'greg king', SAM],
    });
    expect(removeMember.mock.calls).toEqual([
      [GROUP, BEA],
      [GROUP, GREG],
    ]);
    expect(payload(data)).toEqual({
      status: 'done',
      conversation: { id: GROUP, title: 'Project Falcon', kind: 'group' },
      results: [
        { person: 'Example, Bea', mri: BEA, name: 'Bea Example', status: 'removed', verified: true },
        { person: 'greg king', mri: GREG, name: 'King, Greg', status: 'removed', verified: true },
        { person: SAM, mri: SAM, status: 'not_member' },
      ],
    });
  });

  it('returns at once, with no confirmation, when no one named is a member', async () => {
    const { chatsvc, removeMember } = fakeWith();
    const { preview, data } = await confirmed(chatsvc, 'teams_remove_chat_members', { conversation: GROUP, people: [SAM] });
    expect(preview?.confirmToken).toBeUndefined();
    expect(payload(data)).toMatchObject({ status: 'unchanged', results: [{ mri: SAM, status: 'not_member' }] });
    expect(removeMember).not.toHaveBeenCalled();
  });

  it('refuses to remove the signed-in user, pointing at leaving the chat in Teams', async () => {
    const { chatsvc, removeMember } = fakeWith();
    const { first } = await phaseOne(chatsvc, 'teams_remove_chat_members', { conversation: GROUP, people: [BEA, ME_OID] });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(/leave/i);
    expect(removeMember).not.toHaveBeenCalled();
  });

  it.each(NON_GROUPS)('refuses %s', async (_label, conversation, why) => {
    const { chatsvc, removeMember } = fakeWith();
    const { first } = await phaseOne(chatsvc, 'teams_remove_chat_members', { conversation, people: [BEA] });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(why);
    expect(removeMember).not.toHaveBeenCalled();
  });

  it('reports an unconfirmed remove the read-back still shows as unknown', async () => {
    const { chatsvc } = fakeWith({ overrides: { removeMember: vi.fn().mockRejectedValue(unknownOutcome('DELETE')) } });
    const { data } = await confirmed(chatsvc, 'teams_remove_chat_members', { conversation: GROUP, people: [BEA] });
    const p = payload(data);
    expect(p.results).toEqual([{ person: BEA, mri: BEA, name: 'Bea Example', status: 'unknown', verified: false }]);
    expect(p.warning).toMatch(/HTTP 504/);
  });

  it('reports verified: null with a warning when the read-back fails', async () => {
    // Reads before the write succeed; the read-back after it fails.
    const { chatsvc, getThread, removeMember } = fakeWith();
    getThread.mockImplementation(async () => {
      if (removeMember.mock.calls.length > 0) throw new Error('boom');
      return { id: GROUP, properties: { topic: 'Project Falcon' }, members: [{ id: ME }, { id: BEA }] };
    });
    const { data } = await confirmed(chatsvc, 'teams_remove_chat_members', { conversation: GROUP, people: [BEA] });
    const p = payload(data);
    expect(p.results).toEqual([{ person: BEA, mri: BEA, name: 'Bea Example', status: 'removed', verified: null }]);
    expect(p.warning).toMatch(/reading the chat back failed/);
  });
});

describe('annotations', () => {
  it.each([
    ['teams_rename_chat', false],
    ['teams_add_chat_members', false],
    ['teams_remove_chat_members', true],
  ] as const)('%s is an idempotent, open-world write (destructive: %s) taking a confirmToken', async (name, destructive) => {
    const h = await createTestHarness((server) => registerGroupChatTools(server, { chatsvc: fakeWith().chatsvc }));
    try {
      const tool = (await h.client.listTools()).tools.find((t) => t.name === name)!;
      expect(tool.annotations).toMatchObject({
        readOnlyHint: false,
        destructiveHint: destructive,
        idempotentHint: true,
        openWorldHint: true,
      });
      expect(tool.description).toMatch(/confirmToken/);
      expect(tool.description).toMatch(/notice/);
      expect(tool.inputSchema.properties).toHaveProperty('confirmToken');
    } finally {
      await h.close();
    }
  });
});
