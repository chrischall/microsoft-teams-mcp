import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import { registerSendTools } from '../../src/tools/send.js';
import type { ChatServiceClient, RawMessage } from '../../src/chatsvc.js';
import { ME, OTHER, fakeChatsvc, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;

const GROUP = '19:group0001@thread.v2';
const OTHER_GROUP = '19:group0002@thread.v2';
const CHANNEL = '19:chan0001@thread.tacv2';
const THIRD = '8:orgid:00000000-0000-0000-0000-0000000000c3';
const FOURTH = '8:orgid:00000000-0000-0000-0000-0000000000d4';
const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;

const CONVS = [
  rawConv(GROUP, { topic: 'Project Falcon', at: T0 + 4 * MIN }),
  rawConv(OTHER_GROUP, { topic: 'Falcon launch party', at: T0 }),
  rawConv(CHANNEL, { topic: 'Design', at: T0 + 4 * MIN }),
];

/** Recent history: three people besides the user, two of them both called "Sam …". */
function history(): RawMessage[] {
  return [
    rawMsg(T0 + 1 * MIN, { text: 'question', sender: 'Bea Example' }),
    { ...rawMsg(T0 + 2 * MIN, { text: 'hi' }), from: FROM(THIRD), imdisplayname: 'Sam One' },
    { ...rawMsg(T0 + 3 * MIN, { text: 'yo' }), from: FROM(FOURTH), imdisplayname: 'Sam Two' },
    rawMsg(T0 + 4 * MIN, { text: 'mine', fromMe: true }),
  ];
}

type SendBody = {
  content: string;
  clientmessageid: string;
  from: string;
  imdisplayname: string;
  conversationid: string;
  properties: Record<string, unknown>;
};

/**
 * A chat service that "stores" what is sent: a sent message shows up in the
 * next read, carrying the clientmessageid it was sent with — unless `lose` is
 * set, which models a send whose read-back cannot find it.
 */
function fakeWith(opts: { lose?: boolean; overrides?: Record<string, unknown> } = {}) {
  const stored = history();
  const sendMessage = vi.fn(async (conv: string, body: SendBody) => {
    if (!opts.lose) {
      stored.push({ ...rawMsg(T0 + 10 * MIN, { text: body.content, fromMe: true, conv }), clientmessageid: body.clientmessageid });
    }
    return { id: String(T0 + 10 * MIN) };
  });
  const getMessages = vi.fn(async (id: string, o: Parameters<ChatServiceClient['getMessages']>[1]) =>
    messagesFrom(stored)(id, o),
  );
  const chatsvc = fakeChatsvc(CONVS, {
    meDisplayName: vi.fn().mockResolvedValue('Me Myself'),
    sendMessage,
    getMessages,
    getMessage: vi.fn(async (_conv: string, id: string) => stored.find((m) => m.id === id)),
    ...opts.overrides,
  });
  return { chatsvc, sendMessage, getMessages };
}

async function harnessFor(chatsvc: ChatServiceClient) {
  return createTestHarness((server) => registerSendTools(server, { chatsvc }));
}

type PhaseOne = {
  confirmToken?: string;
  preview?: { action?: string; method?: string; path?: string; willSend?: Record<string, unknown> } & Record<string, unknown>;
};

function errorText(result: { content?: unknown[] }): string {
  return (result.content?.[0] as { text?: string })?.text ?? '';
}

/** Phase 1 then phase 2 with its token; returns both. */
async function sendConfirmed(chatsvc: ChatServiceClient, args: Record<string, unknown>) {
  const h = await harnessFor(chatsvc);
  try {
    const first = await h.callTool('teams_send_message', args);
    if (first.isError) return { first, second: first, preview: undefined, data: undefined };
    const preview = parseToolResult<PhaseOne>(first);
    const second = await h.callTool('teams_send_message', { ...args, confirmToken: preview.confirmToken });
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

/** The payload inside the untrusted envelope (the title is someone else's text). */
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

describe('teams_send_message — confirm gate', () => {
  it('phase 1 previews the conversation and rendered text and sends nothing', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    const h = await harnessFor(chatsvc);
    try {
      const first = parseToolResult<PhaseOne>(
        await h.callTool('teams_send_message', { conversation: GROUP, text: 'hello <team>' }),
      );
      expect(first.confirmToken).toEqual(expect.any(String));
      expect(first.preview?.method).toBe('POST');
      expect(first.preview?.path).toBe(`/users/ME/conversations/${encodeURIComponent(GROUP)}/messages`);
      expect(first.preview?.conversation).toEqual({ id: GROUP, title: 'Project Falcon', kind: 'group' });
      expect(first.preview?.text).toBe('hello <team>');
      expect(JSON.stringify(first.preview?.willSend)).toContain('<p>hello &lt;team&gt;</p>');
      expect(sendMessage).not.toHaveBeenCalled();
    } finally {
      await h.close();
    }
  });

  it('phase 2 with the token sends once, in the captured shape, and verifies by read-back', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    const { data } = await sendConfirmed(chatsvc, { conversation: GROUP, text: 'hello\nworld' });

    expect(sendMessage).toHaveBeenCalledTimes(1);
    const [target, body] = sendMessage.mock.calls[0]!;
    expect(target).toBe(GROUP);
    expect(body).toMatchObject({
      type: 'Message',
      conversationid: GROUP,
      from: ME,
      imdisplayname: 'Me Myself',
      content: '<p>hello<br>world</p>',
      messagetype: 'RichText/Html',
    });
    expect(body.clientmessageid).toMatch(/^[1-9]\d{18}$/);

    expect(payload(data)).toEqual({
      status: 'sent',
      conversation: { id: GROUP, title: 'Project Falcon', kind: 'group' },
      message_id: String(T0 + 10 * MIN),
      verified: true,
    });
  });

  it('reports verified: false with a warning when the read-back cannot find the message', async () => {
    const { chatsvc } = fakeWith({ lose: true });
    const { data } = await sendConfirmed(chatsvc, { conversation: GROUP, text: 'hi' });
    expect(payload(data)).toMatchObject({ status: 'sent', verified: false, warning: expect.stringMatching(/check/i) });
  });

  it('still reports the send when the read-back itself fails', async () => {
    // No mentions, so the read-back is the only read of messages.
    const { chatsvc, sendMessage } = fakeWith({ overrides: { getMessages: vi.fn().mockRejectedValue(new Error('boom')) } });
    const { data } = await sendConfirmed(chatsvc, { conversation: GROUP, text: 'hi' });
    expect(sendMessage).toHaveBeenCalledTimes(1);
    expect(payload(data)).toMatchObject({
      status: 'sent',
      message_id: String(T0 + 10 * MIN),
      verified: false,
      warning: expect.stringMatching(/reading the conversation back failed/),
    });
  });

  it('turns an unknown outcome into status "unknown", never an error to retry', async () => {
    const { chatsvc } = fakeWith({
      overrides: {
        sendMessage: vi.fn().mockRejectedValue(new WriteOutcomeUnknownError('Teams', 'POST', { timeoutMs: 60_000 })),
      },
    });
    const { second, data } = await sendConfirmed(chatsvc, { conversation: GROUP, text: 'hi' });
    expect(second.isError).toBeFalsy();
    expect(payload(data)).toMatchObject({
      status: 'unknown',
      warning: expect.stringMatching(/check the conversation before retrying/i),
    });
  });

  it('declares itself a destructive, non-idempotent, open-world write', async () => {
    const h = await harnessFor(fakeWith().chatsvc);
    try {
      const tool = (await h.client.listTools()).tools.find((t) => t.name === 'teams_send_message')!;
      expect(tool.annotations).toMatchObject({
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: true,
      });
      expect(tool.description).toMatch(/confirmToken/);
      expect(tool.description).toMatch(/48:notes/);
      expect(tool.inputSchema.properties).toHaveProperty('confirmToken');
    } finally {
      await h.close();
    }
  });
});

describe('teams_send_message — conversation', () => {
  it('resolves a conversation by name', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    await sendConfirmed(chatsvc, { conversation: 'project falcon', text: 'hi' });
    expect(sendMessage.mock.calls[0]?.[0]).toBe(GROUP);
  });

  it('refuses an ambiguous name, naming the candidates, before any preview', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    const { first } = await sendConfirmed(chatsvc, { conversation: 'falcon', text: 'hi' });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toContain(GROUP);
    expect(errorText(first)).toContain(OTHER_GROUP);
    expect(sendMessage).not.toHaveBeenCalled();
  });

  it('sends to 48:notes, the self-chat', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    const { data } = await sendConfirmed(chatsvc, { conversation: '48:notes', text: 'note to self' });
    expect(sendMessage.mock.calls[0]?.[0]).toBe('48:notes');
    expect(payload(data)).toMatchObject({ status: 'sent', conversation: { id: '48:notes' } });
  });

  it('refuses every other 48: system conversation', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    const { first } = await sendConfirmed(chatsvc, { conversation: '48:calendar', text: 'hi' });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(/48:notes/);
    expect(sendMessage).not.toHaveBeenCalled();
  });
});

describe('teams_send_message — replies and quotes', () => {
  it('posts a channel reply_to into that thread', async () => {
    const { chatsvc, sendMessage, getMessages } = fakeWith();
    const root = String(T0 + 1 * MIN);
    await sendConfirmed(chatsvc, { conversation: CHANNEL, text: 'agreed', reply_to: root });
    const [target, body] = sendMessage.mock.calls[0]!;
    expect(target).toBe(`${CHANNEL};messageid=${root}`);
    expect(body.content).toBe('<p>agreed</p>');
    expect(body.properties.qtdMsgs).toBeUndefined();
    // The read-back reads the channel itself, where its replies are listed.
    expect(getMessages.mock.calls.at(-1)?.[0]).toBe(CHANNEL);
  });

  it('quote-replies a reply_to in a chat', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    const qid = String(T0 + 1 * MIN);
    const { preview } = await sendConfirmed(chatsvc, { conversation: GROUP, text: 'answer', reply_to: qid });
    const [target, body] = sendMessage.mock.calls[0]!;
    expect(target).toBe(GROUP);
    expect(body.content).toBe(
      `<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="${qid}">` +
        `<strong itemprop="mri" itemid="${OTHER}">Bea Example</strong>` +
        `<span itemprop="time" itemid="${qid}"></span><p itemprop="preview">question</p></blockquote>` +
        '<p>answer</p>',
    );
    expect(JSON.parse(body.properties.qtdMsgs as string)).toEqual([{ messageId: qid, sender: OTHER, time: Number(qid) }]);
    expect(preview?.preview?.quoting).toEqual({ id: qid, sender: 'Bea Example', preview: 'question' });
  });

  it('quote-replies with quote in a channel too, alongside the thread reply', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    const root = String(T0 + 1 * MIN);
    const qid = String(T0 + 2 * MIN);
    await sendConfirmed(chatsvc, { conversation: CHANNEL, text: 'x', reply_to: root, quote: qid });
    const [target, body] = sendMessage.mock.calls[0]!;
    expect(target).toBe(`${CHANNEL};messageid=${root}`);
    expect(body.content).toContain(`itemid="${qid}"`);
  });

  it('refuses a quote of a message that is not there', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    const { first } = await sendConfirmed(chatsvc, { conversation: GROUP, text: 'x', quote: '999' });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(/999/);
    expect(sendMessage).not.toHaveBeenCalled();
  });

  it('refuses reply_to and quote naming different messages in a chat', async () => {
    const { chatsvc } = fakeWith();
    const { first } = await sendConfirmed(chatsvc, {
      conversation: GROUP,
      text: 'x',
      reply_to: String(T0 + 1 * MIN),
      quote: String(T0 + 2 * MIN),
    });
    expect(first.isError).toBe(true);
  });
});

describe('teams_send_message — mentions and emoji', () => {
  it('resolves a mention by display name among recent senders', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    const { preview } = await sendConfirmed(chatsvc, {
      conversation: GROUP,
      text: 'thanks @Bea!',
      mentions: [{ name_in_text: 'Bea', person: 'bea example' }],
    });
    const body = sendMessage.mock.calls[0]![1];
    expect(body.content).toBe(
      '<p>thanks <span itemscope itemtype="http://schema.skype.com/Mention" itemid="0">Bea</span>!</p>',
    );
    expect(JSON.parse(body.properties.mentions as string)).toEqual([
      { '@type': 'http://schema.skype.com/Mention', itemid: '0', mri: OTHER, mentionType: 'person', displayName: 'Bea' },
    ]);
    expect(preview?.preview?.mentions).toEqual([{ name_in_text: 'Bea', person: 'Bea Example', mri: OTHER }]);
  });

  it('takes an MRI as-is', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    await sendConfirmed(chatsvc, {
      conversation: GROUP,
      text: '@Zed hello',
      mentions: [{ name_in_text: 'Zed', person: '8:orgid:00000000-0000-0000-0000-0000000000e5' }],
    });
    const body = sendMessage.mock.calls[0]![1];
    expect(JSON.parse(body.properties.mentions as string)[0].mri).toBe('8:orgid:00000000-0000-0000-0000-0000000000e5');
  });

  it('refuses an ambiguous person, listing the candidates', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    const { first } = await sendConfirmed(chatsvc, {
      conversation: GROUP,
      text: '@Sam hi',
      mentions: [{ name_in_text: 'Sam', person: 'sam' }],
    });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toContain('Sam One');
    expect(errorText(first)).toContain('Sam Two');
    expect(sendMessage).not.toHaveBeenCalled();
  });

  it('refuses an unknown person', async () => {
    const { chatsvc } = fakeWith();
    const { first } = await sendConfirmed(chatsvc, {
      conversation: GROUP,
      text: '@Nobody hi',
      mentions: [{ name_in_text: 'Nobody', person: 'Nobody Here' }],
    });
    expect(first.isError).toBe(true);
    expect(errorText(first)).toMatch(/Nobody Here/);
  });

  it('lets the user mention themselves by name (they are a recent sender too)', async () => {
    const { chatsvc } = fakeWith();
    const { first } = await sendConfirmed(chatsvc, {
      conversation: GROUP,
      text: '@Me hi',
      mentions: [{ name_in_text: 'Me', person: 'Me Myself' }],
    });
    expect(first.isError).toBeFalsy();
  });

  it('renders a :key: shortcode as an emoji span', async () => {
    const { chatsvc, sendMessage } = fakeWith();
    await sendConfirmed(chatsvc, { conversation: GROUP, text: 'shipped :1f389_partypopper:' });
    expect(sendMessage.mock.calls[0]![1].content).toContain('itemid="1f389_partypopper"');
  });
});
