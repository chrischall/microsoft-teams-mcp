/**
 * How `teams_get_attachment` calls `convertAttachment`: the explicit limits and
 * the call's cancellation reach it, and a cancelled call stays cancelled rather
 * than degrading to plain delivery. The conversion itself is stubbed — the real
 * one is covered by attachment-convert.test.ts.
 */
import { afterEach, beforeEach, describe, it, expect, vi } from 'vitest';
import { createTestHarness } from '@chrischall/mcp-utils/test';
import type { RawMessage } from '../../src/chatsvc.js';
import type { SharePointClient } from '../../src/sharepoint.js';
import { CONVERT_LIMITS, registerAttachmentTools } from '../../src/tools/attachment.js';
import { fakeChatsvc, messagesFrom, rawMsg } from './fake-chatsvc.js';

const signal = vi.hoisted(() => ({ current: undefined as AbortSignal | undefined }));
const convertAttachment = vi.hoisted(() => vi.fn());

vi.mock('@chrischall/mcp-utils', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@chrischall/mcp-utils')>()),
  currentCallSignal: () => signal.current,
}));
vi.mock('@chrischall/mcp-utils/convert', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@chrischall/mcp-utils/convert')>()),
  convertAttachment,
}));

const T0 = 1_760_000_000_000;
const GROUP = '19:group0001@thread.v2';
const NAME = 'Report.pdf';
const FILE_URL = `https://example.sharepoint.com/sites/demo/Shared%20Documents/${NAME}`;
const PDF = Buffer.from('%PDF-1.4 stand-in');

function messages(): RawMessage[] {
  const files = JSON.stringify([{ fileName: NAME, fileType: 'pdf', objectUrl: FILE_URL }]);
  return [{ ...rawMsg(T0, { text: '<p>see attached</p>' }), messagetype: 'RichText/Html', properties: { files } }];
}

async function getReport() {
  const chatsvc = fakeChatsvc([], { getMessages: messagesFrom(messages()) });
  const sharepoint = {
    downloadFile: vi.fn().mockResolvedValue({ mimeType: 'application/pdf', data: PDF, size: PDF.length }),
  } as unknown as SharePointClient;
  const harness = await createTestHarness((server) => registerAttachmentTools(server, { chatsvc, sharepoint }));
  try {
    return await harness.callTool('teams_get_attachment', { conversation: GROUP, message_id: String(T0), ref: NAME });
  } finally {
    await harness.close();
  }
}

beforeEach(() => {
  signal.current = undefined;
  convertAttachment.mockReset();
  vi.stubEnv('MCP_BLOB_BASE_URL', '');
  vi.stubEnv('MCP_DATA_DIR', '');
});
afterEach(() => vi.unstubAllEnvs());

describe('teams_get_attachment hands the conversion its limits and cancellation', () => {
  it('passes the explicit output limits and the call signal to convertAttachment', async () => {
    const controller = new AbortController();
    signal.current = controller.signal;
    convertAttachment.mockResolvedValue(undefined);

    await getReport();

    expect(convertAttachment).toHaveBeenCalledTimes(1);
    expect(convertAttachment.mock.calls[0][0]).toMatchObject({
      fileName: NAME,
      mode: 'auto',
      limits: CONVERT_LIMITS,
      signal: controller.signal,
    });
    // The caps the README documents, pinned by value so they cannot drift.
    expect(CONVERT_LIMITS).toEqual({ maxTextBytes: 200 * 1024, maxImageBytes: 8 * 1024 * 1024 });
    expect(convertAttachment.mock.calls[0][0]).toMatchObject({
      limits: { maxTextBytes: 204800, maxImageBytes: 8388608 },
    });
  });

  it('passes no signal when the call has none', async () => {
    convertAttachment.mockResolvedValue(undefined);
    await getReport();
    expect(convertAttachment.mock.calls[0][0]).not.toHaveProperty('signal');
  });

  it('re-throws a failure of a cancelled call instead of degrading to plain delivery', async () => {
    const controller = new AbortController();
    signal.current = controller.signal;
    convertAttachment.mockImplementation(async () => {
      controller.abort(new Error('the caller went away'));
      throw new Error('the caller went away');
    });

    const result = await getReport();

    expect(result.isError).toBe(true);
    const text = JSON.stringify(result.content);
    expect(text).toContain('the caller went away');
    expect(text).not.toContain('conversion_error');
    expect((result.content as { type: string }[]).some((b) => b.type === 'resource')).toBe(false);
  });

  it('degrades a failure of a live call to plain delivery, naming the reason', async () => {
    signal.current = new AbortController().signal;
    convertAttachment.mockRejectedValue(new Error('corrupt xref'));

    const result = await getReport();

    expect(result.isError).toBeFalsy();
    expect((result.content as { type: string }[])[0].type).toBe('resource');
    expect(JSON.stringify(result.content)).toContain('corrupt xref');
  });
});
