import { afterEach, beforeEach, describe, it, expect, vi } from 'vitest';
import { existsSync, mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createTestHarness } from '@chrischall/mcp-utils/test';
import { registerAttachmentTools } from '../../src/tools/attachment.js';
import type { ChatServiceClient, RawMessage } from '../../src/chatsvc.js';
import type { SharePointClient } from '../../src/sharepoint.js';
import { fakeChatsvc, messagesFrom, rawMsg } from './fake-chatsvc.js';
import { makeDeck, makePdf, makeWorkbook } from './convert-fixtures.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
const GROUP = '19:group0001@thread.v2';
const TARGET = T0 + 2 * MIN;
const IMAGE_REF = '0-eus-d1-00000000000000000000000000000abc';
const IMAGE_URL = `https://us-api.asm.skype.com/v1/objects/${IMAGE_REF}/views/imgo`;
const PNG_BYTES = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const BLOB_BASE = 'https://host.example/b/reg123';

const PDF_MIME = 'application/pdf';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

const FILE_URL = (name: string) => `https://example.sharepoint.com/sites/demo/Shared%20Documents/${encodeURIComponent(name)}`;
const OPEN_URL = (name: string) => `https://example.sharepoint.com/:x:/r/sites/demo/${encodeURIComponent(name)}`;

/** A message carrying one inline image and one shared file named `name`. */
function messagesWith(name: string): RawMessage[] {
  const files = JSON.stringify([{ fileName: name, fileType: name.split('.').pop(), openUrl: OPEN_URL(name), objectUrl: FILE_URL(name) }]);
  return [
    rawMsg(T0 + MIN),
    {
      ...rawMsg(TARGET, { text: `<p>see attached</p><img src="${IMAGE_URL}" alt="image">` }),
      messagetype: 'RichText/Html',
      properties: { files },
    },
    rawMsg(T0 + 3 * MIN),
  ];
}

function chatsvcFor(name: string): ChatServiceClient {
  return fakeChatsvc([], {
    getMessages: messagesFrom(messagesWith(name)),
    fetchImage: vi.fn().mockResolvedValue({ mimeType: 'image/png', data: PNG_BYTES }),
  });
}

function sharepointServing(mimeType: string, data: Buffer): SharePointClient {
  return { downloadFile: vi.fn().mockResolvedValue({ mimeType, data, size: data.length }) } as unknown as SharePointClient;
}

type Block = { type: string; text?: string; data?: string; mimeType?: string; resource?: { uri: string; blob?: string } };

async function getFile(name: string, mimeType: string, data: Buffer, extra: Record<string, unknown> = {}) {
  const chatsvc = chatsvcFor(name);
  const sharepoint = sharepointServing(mimeType, data);
  const harness = await createTestHarness((server) => registerAttachmentTools(server, { chatsvc, sharepoint }));
  try {
    const result = await harness.callTool('teams_get_attachment', {
      conversation: GROUP,
      message_id: String(TARGET),
      ref: name,
      ...extra,
    });
    return { result, content: (result.content ?? []) as Block[] };
  } finally {
    await harness.close();
  }
}

/** The caption: the last text block, its fields under `data` when it carried its own `note`. */
function captionOf(content: Block[]): Record<string, unknown> {
  const c = JSON.parse(content.filter((b) => b.type === 'text').at(-1)!.text!) as Record<string, unknown> & {
    data?: Record<string, unknown>;
  };
  return c.data ?? c;
}

/** Nothing is ever inlined as an embedded resource or blob when converted content is returned. */
function expectNoOriginalInline(content: Block[]): void {
  expect(content.some((b) => b.type === 'resource')).toBe(false);
}

let outDir: string;
beforeEach(() => {
  outDir = mkdtempSync(join(tmpdir(), 'teams-convert-'));
  vi.stubEnv('TEAMS_OUTPUT_DIR', outDir);
  vi.stubEnv('MCP_BLOB_BASE_URL', '');
  vi.stubEnv('MCP_BLOB_SIGNING_KEY', '');
  vi.stubEnv('MCP_DATA_DIR', '');
});
afterEach(() => {
  vi.unstubAllEnvs();
  vi.unstubAllGlobals();
  rmSync(outDir, { recursive: true, force: true });
});

function stubBlobStore(): { url: string; bytes: number }[] {
  vi.stubEnv('MCP_BLOB_BASE_URL', BLOB_BASE);
  vi.stubEnv('MCP_BLOB_SIGNING_KEY', 'k'.repeat(32));
  const puts: { url: string; bytes: number }[] = [];
  vi.stubGlobal(
    'fetch',
    vi.fn(async (url: string | URL | Request, init?: RequestInit) => {
      puts.push({ url: String(url), bytes: (init?.body as Uint8Array).byteLength });
      return new Response(null, { status: 201 });
    }),
  );
  return puts;
}

describe('teams_get_attachment converts PDF, Excel and PowerPoint files', () => {
  it('reads a PDF as Markdown inside the untrusted envelope, and saves the original locally', async () => {
    const pdf = makePdf(['Hello from the quarterly report']);
    const { result, content } = await getFile('Report.pdf', PDF_MIME, pdf);

    expect(result.isError).toBeFalsy();
    expectNoOriginalInline(content);
    // The envelope opens the result, before any third-party text, and counts the blocks it fences.
    const head = JSON.parse(content[0].text!) as Record<string, unknown>;
    expect(head.untrusted_content).toBe(true);
    expect(head).toMatchObject({ converted_from: 'Report.pdf', format: 'pdf' });
    const fenced = content.slice(1, 1 + (head.blocks as number));
    expect(fenced.length).toBe(head.blocks);
    const markdown = fenced.find((b) => b.type === 'text' && b.text!.includes('text/markdown'));
    expect(markdown?.text).toContain('## Page 1');
    expect(markdown?.text).toContain('Hello from the quarterly report');
    // The caption follows the fenced blocks, and is the last block.
    expect(content.length).toBe(2 + (head.blocks as number));

    const caption = captionOf(content);
    expect(caption).toMatchObject({ name: 'Report.pdf', content_available: true, via: 'disk', bytes: pdf.length });
    expect(caption.converted).toMatchObject({ format: 'pdf', count: 1 });
    expect((caption.converted as { parts: string[] }).parts.length).toBeGreaterThan(0);
    expect(Array.isArray((caption.converted as { notes: unknown }).notes)).toBe(true);
    // The original is still handed over as a file the user can open.
    expect(existsSync(caption.path as string)).toBe(true);
    expect(readFileSync(caption.path as string)).toEqual(pdf);
  });

  it('links the original instead when hosted, never inlining it', async () => {
    const puts = stubBlobStore();
    const pdf = makePdf(['Hosted text']);
    const { result, content } = await getFile('Report.pdf', PDF_MIME, pdf);

    expect(result.isError).toBeFalsy();
    expectNoOriginalInline(content);
    expect(content.some((b) => b.text?.includes('Hosted text'))).toBe(true);
    const caption = captionOf(content);
    expect(caption.via).toBe('link');
    expect(caption.downloadUrl).toMatch(new RegExp(`^${BLOB_BASE}/attachments/`));
    expect(puts).toEqual([expect.objectContaining({ bytes: pdf.length })]);
    expect(readdirSync(outDir)).toHaveLength(0);
  });

  it('still returns the converted content when the original has no route back (hosted, no blob store)', async () => {
    vi.stubEnv('MCP_DATA_DIR', '/data');
    const { result, content } = await getFile('Report.pdf', PDF_MIME, makePdf(['No route text']));

    expect(result.isError).toBeFalsy();
    expectNoOriginalInline(content);
    expect(content.some((b) => b.text?.includes('No route text'))).toBe(true);
    const caption = captionOf(content);
    expect(caption.converted).toMatchObject({ format: 'pdf' });
    expect(caption.via).toBe('none');
    expect(caption.openUrl).toBe(OPEN_URL('Report.pdf'));
    expect(caption.original_unavailable).toMatch(/inline limit|no other way/);
  });

  it('reads a workbook as one CSV per sheet', async () => {
    const { result, content } = await getFile('Budget.xlsx', XLSX_MIME, await makeWorkbook());

    expect(result.isError).toBeFalsy();
    expectNoOriginalInline(content);
    const csv = content.filter((b) => b.type === 'text' && b.text!.includes('text/csv'));
    expect(csv).toHaveLength(2);
    expect(csv[0].text).toContain('coffee,3');
    const caption = captionOf(content);
    expect(caption.converted).toMatchObject({ format: 'xlsx', count: 2 });
    expect((caption.converted as { parts: string[] }).parts).toHaveLength(2);
  });

  it('reads a deck as per-slide Markdown', async () => {
    const { result, content } = await getFile('Plan.pptx', PPTX_MIME, makeDeck());

    expect(result.isError).toBeFalsy();
    expectNoOriginalInline(content);
    expect(content.some((b) => b.type === 'text' && b.text!.includes('Quarterly plan'))).toBe(true);
    expect(captionOf(content).converted).toMatchObject({ format: 'pptx', count: 1 });
  });

  it('detects a PDF served as application/octet-stream by its name', async () => {
    const { content } = await getFile('Report.pdf', 'application/octet-stream', makePdf(['Sniffed by name']));
    expect(content.some((b) => b.text?.includes('Sniffed by name'))).toBe(true);
    expect(captionOf(content).converted).toMatchObject({ format: 'pdf' });
  });

  it("convert: 'images' renders PDF pages as images", async () => {
    const { result, content } = await getFile('Report.pdf', PDF_MIME, makePdf(['Rendered page']), { convert: 'images' });
    expect(result.isError).toBeFalsy();
    expect(content.some((b) => b.type === 'image' && /^image\/(png|jpeg)$/.test(b.mimeType ?? ''))).toBe(true);
    expectNoOriginalInline(content);
  });

  it("convert: 'auto' adds page images for a scan (a page with no text), or says why it cannot", async () => {
    const { result, content } = await getFile('Scan.pdf', PDF_MIME, makePdf(['']));
    expect(result.isError).toBeFalsy();
    expectNoOriginalInline(content);
    const notes = (captionOf(content).converted as { notes: string[] }).notes;
    expect(notes.some((n) => /little or no text layer/.test(n))).toBe(true);
    const pageImages = content.filter((b) => b.type === 'image' && /^image\/(png|jpeg)$/.test(b.mimeType ?? ''));
    // Page images need @napi-rs/canvas; without it the degrade is a note, never silence.
    expect(pageImages.length > 0 || notes.some((n) => /page images unavailable/.test(n))).toBe(true);
  });

  it("convert: 'text' never renders page images, even for a scan", async () => {
    const { content } = await getFile('Scan.pdf', PDF_MIME, makePdf(['']), { convert: 'text' });
    expect(content.some((b) => b.type === 'image')).toBe(false);
    expect(captionOf(content).converted).toMatchObject({ format: 'pdf' });
  });

  it("convert: 'none' keeps the old behaviour: the original inline, no conversion", async () => {
    const pdf = makePdf(['Untouched']);
    const { content } = await getFile('Report.pdf', PDF_MIME, pdf, { convert: 'none' });
    expect(content[0].type).toBe('resource');
    expect(content[0].resource?.blob).toBe(pdf.toString('base64'));
    expect(captionOf(content)).not.toHaveProperty('converted');
  });

  it('save: true skips conversion — the user asked for the file itself', async () => {
    const pdf = makePdf(['Saved, not read']);
    const { content } = await getFile('Report.pdf', PDF_MIME, pdf, { save: true });
    expect(content.map((b) => b.type)).toEqual(['text']);
    const caption = captionOf(content);
    expect(caption).not.toHaveProperty('converted');
    expect(caption.via).toBe('disk');
    expect(readFileSync(caption.path as string)).toEqual(pdf);
  });

  it('falls back to the old behaviour, naming the reason, when the file will not convert', async () => {
    const corrupt = Buffer.from('%PDF-1.4\nthis is not really a pdf\n');
    const { result, content } = await getFile('Broken.pdf', PDF_MIME, corrupt);

    expect(result.isError).toBeFalsy();
    expect(content[0].type).toBe('resource');
    expect(content[0].resource?.blob).toBe(corrupt.toString('base64'));
    const caption = captionOf(content);
    expect(caption).not.toHaveProperty('converted');
    expect(typeof caption.conversion_error).toBe('string');
    expect((caption.conversion_error as string).length).toBeGreaterThan(0);
  });

  it('leaves a file whose bytes are not the format it claims as the old behaviour', async () => {
    const notPdf = Buffer.from('plain text pretending');
    const { content } = await getFile('Report.pdf', PDF_MIME, notPdf);
    expect(content[0].type).toBe('resource');
    expect(captionOf(content)).not.toHaveProperty('converted');
  });

  it('leaves a non-convertible file unchanged', async () => {
    const bytes = Buffer.from('PK\u0003\u0004 pretend docx');
    const { content } = await getFile('Notes.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', bytes);
    expect(content[0].type).toBe('resource');
    const caption = captionOf(content);
    expect(caption).not.toHaveProperty('converted');
    expect(caption).not.toHaveProperty('conversion_error');
  });

  it('leaves inline images unaffected', async () => {
    const chatsvc = chatsvcFor('Report.pdf');
    const sharepoint = sharepointServing(PDF_MIME, makePdf(['x']));
    const harness = await createTestHarness((server) => registerAttachmentTools(server, { chatsvc, sharepoint }));
    try {
      const result = await harness.callTool('teams_get_attachment', {
        conversation: GROUP,
        message_id: String(TARGET),
        ref: IMAGE_REF,
        convert: 'auto',
      });
      const content = result.content as Block[];
      expect(content[0]).toEqual({ type: 'image', data: PNG_BYTES.toString('base64'), mimeType: 'image/png' });
      expect(captionOf(content)).not.toHaveProperty('converted');
    } finally {
      await harness.close();
    }
  });

  it('documents the convert input', async () => {
    const chatsvc = chatsvcFor('Report.pdf');
    const sharepoint = sharepointServing(PDF_MIME, makePdf(['x']));
    const harness = await createTestHarness((server) => registerAttachmentTools(server, { chatsvc, sharepoint }));
    try {
      const tools = await harness.client.listTools();
      const tool = tools.tools.find((t) => t.name === 'teams_get_attachment')!;
      const convert = (tool.inputSchema.properties as Record<string, { enum?: string[]; description?: string }>).convert;
      expect(convert.enum).toEqual(['auto', 'text', 'images', 'none']);
      expect(convert.description).toMatch(/CSV/);
      expect(convert.description).toMatch(/slide/i);
      expect(tool.description).toMatch(/PDF/);
    } finally {
      await harness.close();
    }
  });
});
