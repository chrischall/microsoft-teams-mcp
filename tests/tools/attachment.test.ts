import { describe, it, expect, vi } from 'vitest';
import { createTestHarness, parseToolResult } from '@chrischall/mcp-utils/test';
import { McpToolError } from '@chrischall/mcp-utils';
import { registerAttachmentTools } from '../../src/tools/attachment.js';
import type { ChatServiceClient, GetMessagesOptions, RawMessage } from '../../src/chatsvc.js';
import type { SharePointClient } from '../../src/sharepoint.js';
import { fakeChatsvc, messagesFrom, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
const GROUP = '19:group0001@thread.v2';

const IMAGE_REF = '0-eus-d1-00000000000000000000000000000abc';
const IMAGE_URL = `https://us-api.asm.skype.com/v1/objects/${IMAGE_REF}/views/imgo`;
const PNG_BYTES = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const FILE_NAME = 'Example plan.docx';
const FILE_OPEN_URL = 'https://example.sharepoint.com/:w:/r/sites/demo/Example%20plan.docx';
const FILE_URL = 'https://example.sharepoint.com/sites/demo/Shared%20Documents/Example%20plan.docx';
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
/** 25 MiB — must match MAX_FILE_BYTES in the tool. */
const MAX_FILE_BYTES = 25 * 1024 * 1024;

const FILES = JSON.stringify([
  {
    fileName: FILE_NAME,
    fileType: 'docx',
    openUrl: FILE_OPEN_URL,
    objectUrl: FILE_URL,
  },
]);

/** A message carrying one inline image and one shared file. */
function withAttachments(at: number): RawMessage {
  return {
    ...rawMsg(at, { text: `<p>look at this</p><img src="${IMAGE_URL}" alt="image">` }),
    messagetype: 'RichText/Html',
    properties: { files: FILES },
  };
}

const TARGET = T0 + 2 * MIN;
const MSGS = [rawMsg(T0 + 1 * MIN), withAttachments(TARGET), rawMsg(T0 + 3 * MIN), rawMsg(T0 + 4 * MIN)];

function chatsvcWith(messages = MSGS, overrides: Record<string, unknown> = {}): ChatServiceClient {
  return fakeChatsvc([], {
    getMessages: messagesFrom(messages),
    fetchImage: vi.fn().mockResolvedValue({ mimeType: 'image/png', data: PNG_BYTES }),
    ...overrides,
  });
}

/**
 * A stub `SharePointClient` — the tool only ever calls `downloadFile`, so that
 * is all it carries. By default it rejects (no SP token in a test), which is the
 * graceful-degrade path; a test that exercises a download injects its own.
 */
function fakeSharepoint(
  downloadFile: SharePointClient['downloadFile'] = vi.fn().mockRejectedValue(new Error('no token in tests')),
): SharePointClient {
  return { downloadFile } as unknown as SharePointClient;
}

async function call(
  chatsvc: ChatServiceClient,
  args: Record<string, unknown>,
  sharepoint: SharePointClient = fakeSharepoint(),
) {
  const harness = await createTestHarness((server) => registerAttachmentTools(server, { chatsvc, sharepoint }));
  try {
    return await harness.callTool('teams_get_attachment', args);
  } finally {
    await harness.close();
  }
}

function errorText(result: { content?: unknown[] }): string {
  return (result.content?.[0] as { text?: string })?.text ?? '';
}

type Block = { type: string; text?: string; data?: string; mimeType?: string; resource?: { uri: string; mimeType?: string; blob?: string } };

describe('teams_get_attachment', () => {
  it('returns an inline image as MCP image content plus an untrusted caption', async () => {
    const chatsvc = chatsvcWith();
    const result = await call(chatsvc, { conversation: GROUP, message_id: String(TARGET), ref: IMAGE_REF });

    expect(result.isError).toBeFalsy();
    const content = result.content as Block[];
    expect(content[0]).toEqual({ type: 'image', data: PNG_BYTES.toString('base64'), mimeType: 'image/png' });
    expect(content[1].type).toBe('text');
    const caption = JSON.parse(content[1].text!) as Record<string, unknown>;
    expect(caption.untrusted_content).toBe(true);
    expect(JSON.stringify(caption)).toContain(IMAGE_REF);
    expect(chatsvc.fetchImage).toHaveBeenCalledWith(IMAGE_URL);
    expect(chatsvc.getMessages).toHaveBeenCalledWith(GROUP, expect.anything());
  });

  it('stops paging once it is past the message, so an old message does not read the whole history', async () => {
    const chatsvc = chatsvcWith();
    await call(chatsvc, { conversation: GROUP, message_id: String(TARGET), ref: IMAGE_REF });

    const opts = vi.mocked(chatsvc.getMessages).mock.calls[0][1] as GetMessagesOptions;
    expect(opts.stopAt?.(rawMsg(T0 + 3 * MIN))).toBe(false);
    expect(opts.stopAt?.(rawMsg(TARGET))).toBe(false);
    expect(opts.stopAt?.(rawMsg(T0 + 1 * MIN))).toBe(true);
  });

  it('downloads a shared file as an MCP embedded resource plus an untrusted caption', async () => {
    const bytes = Buffer.from('PK\u0003\u0004 pretend docx body');
    const downloadFile = vi.fn().mockResolvedValue({ mimeType: DOCX_MIME, data: bytes, size: bytes.length });
    const result = await call(
      chatsvcWith(),
      { conversation: GROUP, message_id: String(TARGET), ref: FILE_NAME },
      fakeSharepoint(downloadFile),
    );

    expect(result.isError).toBeFalsy();
    const content = result.content as Block[];
    expect(content[0].type).toBe('resource');
    expect(content[0].resource).toEqual({ uri: FILE_OPEN_URL, mimeType: DOCX_MIME, blob: bytes.toString('base64') });
    expect(content[1].type).toBe('text');
    const caption = JSON.parse(content[1].text!) as Record<string, unknown>;
    expect(caption.untrusted_content).toBe(true);
    expect(caption.mimeType).toBe(DOCX_MIME);
    expect(caption.bytes).toBe(bytes.length);
    // A successful download carries content_available:true (the degrade path is false).
    expect(caption.content_available).toBe(true);
    expect(JSON.stringify(caption)).toContain(FILE_NAME);
    expect(downloadFile).toHaveBeenCalledWith(FILE_URL, MAX_FILE_BYTES);
  });

  it('downloads a shared file that is an image as MCP image content', async () => {
    const downloadFile = vi.fn().mockResolvedValue({ mimeType: 'image/png', data: PNG_BYTES, size: PNG_BYTES.length });
    const result = await call(
      chatsvcWith(),
      { conversation: GROUP, message_id: String(TARGET), ref: FILE_NAME },
      fakeSharepoint(downloadFile),
    );

    expect(result.isError).toBeFalsy();
    const content = result.content as Block[];
    expect(content[0]).toEqual({ type: 'image', data: PNG_BYTES.toString('base64'), mimeType: 'image/png' });
    expect(content[1].type).toBe('text');
  });

  it('degrades a forbidden file to metadata and its link, not a thrown error', async () => {
    const downloadFile = vi.fn().mockRejectedValue(
      new McpToolError('SharePoint refused access to the file (403).', {
        status: 403,
        kind: 'http',
        hint: 'Your account cannot open this file.',
      }),
    );
    const result = await call(
      chatsvcWith(),
      { conversation: GROUP, message_id: String(TARGET), ref: FILE_NAME },
      fakeSharepoint(downloadFile),
    );

    expect(result.isError).toBeFalsy();
    const data = parseToolResult(result) as Record<string, unknown>;
    expect(data).toMatchObject({ name: FILE_NAME, openUrl: FILE_OPEN_URL, content_available: false });
    expect(data.hint).toMatch(/cannot open/i);
  });

  it('degrades a token-capture failure to metadata and its link, not a thrown error', async () => {
    const downloadFile = vi.fn().mockRejectedValue(new Error('could not capture a SharePoint token (timed out).'));
    const result = await call(
      chatsvcWith(),
      { conversation: GROUP, message_id: String(TARGET), ref: FILE_NAME },
      fakeSharepoint(downloadFile),
    );

    expect(result.isError).toBeFalsy();
    const data = parseToolResult(result) as Record<string, unknown>;
    expect(data).toMatchObject({ name: FILE_NAME, content_available: false });
    expect(data.hint).toMatch(/open the (shared )?file once/i);
  });

  it('errors when the message is not in the conversation', async () => {
    const chatsvc = chatsvcWith();
    const result = await call(chatsvc, { conversation: GROUP, message_id: String(T0 + 99 * MIN), ref: IMAGE_REF });

    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/message .* not found/i);
    expect(errorText(result)).toMatch(/teams_read_conversation/);
  });

  it('errors when the message has no attachment with that ref or name', async () => {
    const chatsvc = chatsvcWith();
    const result = await call(chatsvc, { conversation: GROUP, message_id: String(TARGET), ref: 'nope.pdf' });

    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/no attachment/i);
    expect(chatsvc.fetchImage).not.toHaveBeenCalled();
  });

  it('refuses to return a non-image response as image content', async () => {
    const chatsvc = chatsvcWith(MSGS, {
      fetchImage: vi.fn().mockResolvedValue({ mimeType: 'text/html', data: Buffer.from('<html></html>') }),
    });
    const result = await call(chatsvc, { conversation: GROUP, message_id: String(TARGET), ref: IMAGE_REF });

    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/not an image/i);
  });

  it('passes an already-actionable image-fetch error through unchanged', async () => {
    const chatsvc = chatsvcWith(MSGS, {
      fetchImage: vi.fn().mockRejectedValue(new McpToolError('Teams chat service returned HTTP 404.', { status: 404 })),
    });
    const result = await call(chatsvc, { conversation: GROUP, message_id: String(TARGET), ref: IMAGE_REF });

    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/HTTP 404/);
  });

  it('turns a token-capture failure into the open-a-Teams-tab hint', async () => {
    const chatsvc = chatsvcWith(MSGS, {
      getMessages: vi.fn().mockRejectedValue(new Error('could not capture a Teams token from the browser (timed out).')),
    });
    const result = await call(chatsvc, { conversation: GROUP, message_id: String(TARGET), ref: IMAGE_REF });

    expect(result.isError).toBe(true);
    expect(errorText(result)).toMatch(/click into any chat/i);
  });
});
