import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import { registerEditTools } from '../../src/tools/edit.js';
import { UNTRUSTED_CONTENT_NOTE } from '../../src/tools/untrusted.js';
import type { ChatServiceClient, RawMessage } from '../../src/chatsvc.js';
import { OTHER, fakeChatsvc, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
const GROUP = '19:group0001@thread.v2';
const CONVS = [rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * MIN })];

const MINE = String(T0 + MIN);
const THEIRS = String(T0 + 2 * MIN);
const QUOTING = String(T0 + 3 * MIN);
const GONE = String(T0 + 4 * MIN);
const EVENT = String(T0 + 5 * MIN);

const QUOTE_BLOCK =
  `<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="${THEIRS}">` +
  `<strong itemprop="mri" itemid="${OTHER}">Bea Example</strong><span itemprop="time" itemid="${THEIRS}"></span>` +
  '<p itemprop="preview">question</p></blockquote>';

function history(): RawMessage[] {
  return [
    rawMsg(T0 + MIN, { text: '<p>first draft</p>', fromMe: true }),
    rawMsg(T0 + 2 * MIN, { text: '<p>question</p>', sender: 'Bea Example' }),
    {
      ...rawMsg(T0 + 3 * MIN, { text: `${QUOTE_BLOCK}<p>my answer</p>`, fromMe: true }),
      properties: { qtdMsgs: JSON.stringify([{ messageId: THEIRS, sender: OTHER, time: Number(THEIRS) }]) },
    },
    rawMsg(T0 + 4 * MIN, { text: '', fromMe: true, deleted: true }),
    rawMsg(T0 + 5 * MIN, { text: '<p>call ended</p>', fromMe: true, type: 'Event/Call' }),
  ];
}

type EditBody = { content: string; messagetype: string; properties: Record<string, string> };

/** A chat service whose edits land on the stored message, unless `fail` says how the PUT goes wrong. */
function fakeWith(opts: { fail?: Error; land?: boolean } = {}) {
  const stored = history();
  const editMessage = vi.fn(async (_conv: string, id: string, body: EditBody) => {
    const land = opts.land ?? opts.fail === undefined;
    if (land) {
      const m = stored.find((x) => x.id === id)!;
      m.content = body.content;
      m.properties = { ...m.properties, ...body.properties, edittime: String(T0 + 9 * MIN) };
    }
    if (opts.fail) throw opts.fail;
  });
  const chatsvc = fakeChatsvc(CONVS, {
    editMessage,
    getMessage: vi.fn(async (_conv: string, id: string) => {
      const m = stored.find((x) => x.id === id);
      return m === undefined ? undefined : structuredClone(m);
    }),
  });
  return { chatsvc, editMessage };
}

type PhaseOne = { confirmToken?: string; preview?: Record<string, any> };

async function edit(chatsvc: ChatServiceClient, args: Record<string, unknown>) {
  const h = await createTestHarness((server) => registerEditTools(server, { chatsvc }));
  try {
    const first = await h.callTool('teams_edit_message', args);
    if (first.isError) return { first, error: (first.content?.[0] as { text?: string })?.text ?? '' };
    const one = parseToolResult<PhaseOne & Record<string, any>>(first);
    if (one.confirmToken === undefined) return { first, data: one };
    const second = await h.callTool('teams_edit_message', { ...args, confirmToken: one.confirmToken });
    return {
      first,
      preview: one.preview,
      second,
      error: second.isError ? ((second.content?.[0] as { text?: string })?.text ?? '') : undefined,
      data: second.isError ? undefined : (parseToolResult(second) as Record<string, any>),
    };
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

describe('teams_edit_message', () => {
  it('previews the current and new text and the rendered HTML, and changes nothing in phase 1', async () => {
    const { chatsvc, editMessage } = fakeWith();
    const h = await createTestHarness((server) => registerEditTools(server, { chatsvc }));
    try {
      const one = parseToolResult<PhaseOne>(
        await h.callTool('teams_edit_message', { conversation: GROUP, message_id: MINE, text: '**final**', format: 'markdown' }),
      );
      expect(one.confirmToken).toEqual(expect.any(String));
      expect(one.preview).toMatchObject({
        method: 'PUT',
        path: `/users/ME/conversations/${encodeURIComponent(GROUP)}/messages/${MINE}`,
        message_id: MINE,
        text: '**final**',
        format: 'markdown',
      });
      expect(fenced(one.preview?.conversation)).toEqual({ id: GROUP, title: 'Project Falcon', kind: 'group' });
      expect(fenced(one.preview?.current_text)).toBe('first draft');
      expect(one.preview?.willSend?.content).toBe('<p><strong>final</strong></p>');
      expect(editMessage).not.toHaveBeenCalled();
    } finally {
      await h.close();
    }
  });

  it('replaces the message once confirmed, and verifies it by reading it back', async () => {
    const { chatsvc, editMessage } = fakeWith();
    const { data } = await edit(chatsvc, {
      conversation: GROUP,
      message_id: MINE,
      text: 'now with @Bea and 1. a list',
      mentions: [{ name_in_text: 'Bea', person: OTHER }],
    });
    expect(editMessage).toHaveBeenCalledTimes(1);
    const [conv, id, body] = editMessage.mock.calls[0]!;
    expect([conv, id]).toEqual([GROUP, MINE]);
    expect(body).toMatchObject({ messagetype: 'RichText/Html', contenttype: 'text' });
    expect(body.content).toContain('itemtype="http://schema.skype.com/Mention"');
    expect(JSON.parse(body.properties.mentions)).toEqual([expect.objectContaining({ mri: OTHER, displayName: 'Bea' })]);
    expect(data).toMatchObject({ status: 'edited', message_id: MINE, verified: true });
  });

  it('keeps the quote of a quote-reply and its qtdMsgs', async () => {
    const { chatsvc, editMessage } = fakeWith();
    await edit(chatsvc, { conversation: GROUP, message_id: QUOTING, text: 'a better answer' });
    const [, , body] = editMessage.mock.calls[0]!;
    expect(body.content).toBe(`${QUOTE_BLOCK}<p>a better answer</p>`);
    expect(JSON.parse(body.properties.qtdMsgs)).toEqual([{ messageId: THEIRS, sender: OTHER, time: Number(THEIRS) }]);
  });

  it('fences the current text, since a quote-reply carries what someone else wrote', async () => {
    const { chatsvc } = fakeWith();
    const h = await createTestHarness((server) => registerEditTools(server, { chatsvc }));
    try {
      const one = parseToolResult<PhaseOne>(await h.callTool('teams_edit_message', { conversation: GROUP, message_id: QUOTING, text: 'x' }));
      expect(fenced(one.preview?.current_text)).toBe('> Bea Example\n> question\nmy answer');
    } finally {
      await h.close();
    }
  });

  it('counts text as unchanged when Teams stored it with CRLFs between blocks', async () => {
    const { chatsvc, editMessage } = fakeWith();
    const stored = await chatsvc.getMessage(GROUP, MINE);
    (chatsvc.getMessage as ReturnType<typeof vi.fn>).mockResolvedValue({
      ...stored,
      content: '<p><strong>Plan</strong></p>\r\n<ol>\r\n<li>one</li><li>two</li></ol>',
    });
    const { data } = await edit(chatsvc, { conversation: GROUP, message_id: MINE, text: '**Plan**\n\n1. one\n2. two', format: 'markdown' });
    expect(data).toMatchObject({ status: 'unchanged' });
    expect(editMessage).not.toHaveBeenCalled();
  });

  it('counts text as unchanged when Teams put a line break before a nested list (live 2026-11-11)', async () => {
    const { chatsvc, editMessage } = fakeWith();
    const stored = await chatsvc.getMessage(GROUP, MINE);
    (chatsvc.getMessage as ReturnType<typeof vi.fn>).mockResolvedValue({
      ...stored,
      content: '<ol>\r\n<li>top\r\n<ul>\r\n<li>inner</li></ul>\r\n</li><li>next</li></ol>',
    });
    const { data } = await edit(chatsvc, { conversation: GROUP, message_id: MINE, text: '1. top\n   - inner\n2. next', format: 'markdown' });
    expect(data).toMatchObject({ status: 'unchanged' });
    expect(editMessage).not.toHaveBeenCalled();
  });

  it('changes nothing, and asks nothing, when the new text renders the same', async () => {
    const { chatsvc, editMessage } = fakeWith();
    const { data } = await edit(chatsvc, { conversation: GROUP, message_id: MINE, text: 'first draft' });
    expect(data).toMatchObject({ status: 'unchanged', message_id: MINE });
    expect(editMessage).not.toHaveBeenCalled();
  });

  it.each([
    [THEIRS, /only your own messages/i],
    ['1', /no message 1/i],
    [GONE, /deleted/i],
    [EVENT, /not a message anyone wrote/i],
  ])('refuses message %s before any preview', async (id, why) => {
    const { chatsvc, editMessage } = fakeWith();
    const { error, first } = await edit(chatsvc, { conversation: GROUP, message_id: id, text: 'x' });
    expect(first.isError).toBe(true);
    expect(error).toMatch(why);
    expect(editMessage).not.toHaveBeenCalled();
  });

  it('reports an unconfirmed edit that the read-back shows as edited, with a warning', async () => {
    const { chatsvc } = fakeWith({ fail: new WriteOutcomeUnknownError('Teams', 'PUT', { hint: 'x' }), land: true });
    const { data } = await edit(chatsvc, { conversation: GROUP, message_id: MINE, text: 'changed' });
    expect(data).toMatchObject({ status: 'edited', verified: true, warning: expect.stringMatching(/did not confirm/) });
  });

  it('reports status unknown when Teams did not confirm and the read-back does not show the edit', async () => {
    const { chatsvc } = fakeWith({ fail: new WriteOutcomeUnknownError('Teams', 'PUT', { hint: 'x' }), land: false });
    const { data } = await edit(chatsvc, { conversation: GROUP, message_id: MINE, text: 'changed' });
    expect(data).toMatchObject({ status: 'unknown', verified: false });
  });

  it('is declared a destructive, idempotent write', async () => {
    const { chatsvc } = fakeWith();
    const h = await createTestHarness((server) => registerEditTools(server, { chatsvc }));
    try {
      const tool = (await h.client.listTools()).tools.find((t) => t.name === 'teams_edit_message');
      expect(tool?.annotations).toMatchObject({ readOnlyHint: false, destructiveHint: true, idempotentHint: true });
    } finally {
      await h.close();
    }
  });
});
