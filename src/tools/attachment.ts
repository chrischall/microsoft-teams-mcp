/**
 * `teams_get_attachment` — fetch one attachment of one message, from the
 * Teams chat-service API.
 *
 * The message is re-read (newest first, stopping once paging is past it), so
 * the attachment is located in what the service says now rather than trusted
 * from a ref the caller supplies: a ref only ever selects among the URLs the
 * message itself carries, and `fetchImage` still vets the host before the
 * bearer goes anywhere.
 *
 * Two kinds, verified live:
 * - an inline image (Teams' own image store) downloads with the chat-service
 *   token and comes back as MCP image content;
 * - a shared file lives in SharePoint, downloaded through `SharePointClient`
 *   with a separate SharePoint token that every SharePoint host of the tenant
 *   accepts (see `sharepoint.ts`). A file the account cannot open — or one
 *   fetched before a token could be captured — comes back as metadata and its
 *   open link, degrading gracefully rather than failing the tool.
 *
 * `ref` is matched exactly, else ignoring case, else (a file) ignoring case and
 * extension, else ignoring differences in spacing (Unicode and zero-width
 * spaces); a loose match naming more than one attachment is refused. `save`
 * asks for a file in the download folder rather than the contents.
 *
 * Either kind is handed back through `deliverFile` (mcp-utils): inline up to
 * 10 MiB (an image as MCP image content, anything else as an embedded
 * resource); past that a signed download link from the blob store when hosted,
 * else a fresh file in the download folder on a local install. Inlining up to
 * the 25 MiB download cap, as this tool once did, put ~33 MiB of base64 in one
 * result — over mcp-host's 14 MiB result ceiling, which swaps it for a generic
 * "result too large" error.
 *
 * A shared PDF, workbook (.xlsx/.xlsm) or deck (.pptx) is also CONVERTED
 * (`convert`, mcp-utils' `/convert`): hosted clients render only images, so a
 * PDF handed over as a link was a file the model could not read. The converted
 * Markdown / CSV / images come back fenced by the untrusted envelope (a header
 * block counting the blocks that follow — block boundaries cannot be forged by
 * the file's text), then the caption; the original is still handed over as a
 * link (hosted) or a saved file (local), but never inlined beside its
 * conversion. A conversion that fails falls back to the plain delivery above,
 * with the reason in the caption's `conversion_error`. `save` skips it.
 */
import type { CallToolResult, McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import {
  McpToolError,
  currentCallSignal,
  errorKindOf,
  errorStatusOf,
  expandPath,
  readEnvVar,
  toolAnnotations,
} from '@chrischall/mcp-utils';
import { deliverFile, isHostedEnv, type DeliveredFile } from '@chrischall/mcp-utils/export';
import {
  convertAttachment,
  convertedToContent,
  detectConvertible,
  type ConvertResult,
} from '@chrischall/mcp-utils/convert';
import { homedir } from 'node:os';
import { join } from 'node:path';
import type { RawMessage } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import type { FileAttachment, NormalisedMessage } from '../model.js';
import { normaliseMessage } from '../model.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';
import { nameConversations, type ConversationName } from './titles.js';
import { workMeta } from './hints.js';
import { EXPORT_OUTPUT_ENV, EXPORT_SERVER_NAME } from './export.js';

/**
 * Raw messages read looking for the one asked for (20 pages of 200). The
 * `stopAt` boundary — paging past the message's arrival time — is the real
 * bound; this only caps an id that is not an arrival time.
 */
const MAX_RAW_MESSAGES = 4000;

/**
 * Largest shared file whose contents are downloaded. Past this the bounded read
 * in `SharePointClient` aborts and the file degrades to metadata + its open
 * link. What is downloaded is not necessarily inlined: `deliverFile` inlines
 * only up to 10 MiB and hands anything larger over as a link or a saved file.
 */
export const MAX_FILE_BYTES = 25 * 1024 * 1024;

/**
 * Output caps for a conversion: the library defaults, spelled out because they
 * are what keeps a converted result (text + base64 images + caption) under
 * mcp-host's 14 MiB result ceiling.
 */
export const CONVERT_LIMITS = { maxTextBytes: 200 * 1024, maxImageBytes: 8 * 1024 * 1024 } as const;

const CONVERT_MODES = ['auto', 'text', 'images', 'none'] as const;
type ConvertOption = (typeof CONVERT_MODES)[number];

const CONVERT_DESCRIPTION =
  'How to read a shared PDF, Excel workbook (.xlsx/.xlsm) or PowerPoint deck (.pptx) so its contents can be read ' +
  'here. "auto" (default): a PDF as Markdown, plus page images when it has little text (a scan); a workbook as one ' +
  'CSV per sheet; a deck as per-slide Markdown (titles, text, tables, speaker notes) plus its embedded pictures. ' +
  '"text": a PDF as Markdown only, never page images. "images": a PDF as page images (workbooks and decks convert ' +
  'as "auto"). "none": no conversion — the file is handed over as it is. The original is still returned as a ' +
  'download link (hosted) or a saved file (local). Ignored with save: true and for inline images and other formats.';

const FILE_HINT =
  'A shared file is downloaded from SharePoint and returned as its contents (an image as an image, ' +
  'anything else as an embedded resource; over 10 MiB, or with save, as a download link or a saved file ' +
  'on a local install). The first download may need your Teams tab reloaded to capture a SharePoint ' +
  'token; a file your account cannot open comes back as metadata and its link.';

export function registerAttachmentTools(
  server: McpServer,
  { chatsvc, sharepoint }: Pick<TeamsDeps, 'chatsvc' | 'sharepoint'>,
): void {
  server.registerTool(
    'teams_get_attachment',
    {
      title: 'Get a Teams message attachment',
      description:
        'Fetch one attachment of a Teams chat or channel message, as listed by teams_list_messages / ' +
        "teams_list_unread_messages. ref names it: an inline image's ref (attachments.images[].ref) or a shared " +
        "file's name (attachments.files[].name), matched ignoring case and spacing (a Unicode space such as the " +
        'one macOS puts before AM/PM matches a plain one) and, for a file, with or without its extension. ' +
        'An inline image is returned as the image itself. A shared file is downloaded from SharePoint and ' +
        'returned as its contents: an image as the image, anything else as an embedded resource carrying the ' +
        'bytes. Over 10 MiB it comes back instead as a signed, expiring download link (downloadUrl) on a hosted ' +
        `install, or saved to a fresh file (path) under ${EXPORT_OUTPUT_ENV}, else ` +
        `~/Downloads/${EXPORT_SERVER_NAME}, on a local one; the caption's via says which. save: true does the same ` +
        'for any size — use it when the user wants the file itself rather than its contents read here. ' +
        'A shared PDF, Excel workbook or PowerPoint deck is also read for you (see convert): its Markdown, CSV ' +
        'or images come back with the original as a link or saved file, never inlined beside them; the ' +
        "caption's converted lists what came back, and a file that will not convert is handed over as above " +
        'with the reason in conversion_error. ' +
        'The first download may need your Teams tab reloaded to capture a short-lived SharePoint token, which ' +
        'then serves every SharePoint host; a file your account cannot open, or one that could not be fetched, ' +
        'comes back as metadata and its open link instead of failing. The caption names the conversation beside ' +
        'its id: title (as teams_list_conversations names it — an untitled group chat by its participants, a 1:1 ' +
        'by the other person) and participants. Does not mark anything read.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      _meta: workMeta('other', { aliases: ['attachment', 'image', 'file', 'screenshot'] }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe('The conversation id the message is in (from teams_list_messages or teams_list_unread_messages).'),
        message_id: z.string().min(1).describe('The message id, exactly as a message listing returned it.'),
        ref: z
          .string()
          .min(1)
          .describe(
            "An inline image's ref (attachments.images[].ref) or a shared file's name (attachments.files[].name). " +
              'Case and spacing are ignored, and a file may be named without its extension.',
          ),
        save: z
          .boolean()
          .optional()
          .describe(
            `Save the attachment to a fresh file under ${EXPORT_OUTPUT_ENV} (else ~/Downloads/${EXPORT_SERVER_NAME}) ` +
              'and return its path instead of its contents; a hosted install returns a download link instead. Default false.',
          ),
        convert: z.enum(CONVERT_MODES).optional().describe(CONVERT_DESCRIPTION),
      },
    },
    async ({ conversation, message_id, ref, save = false, convert = 'auto' }) => {
      try {
        // A message id IS its arrival ms, and pages run newest first: the first
        // message older than it means it was not there.
        const target = Number(message_id);
        const stopAt = Number.isFinite(target) && target > 0 ? (m: RawMessage) => Number(m.id) < target : undefined;
        const fetched = await chatsvc.getMessages(conversation, { stopAt, max: MAX_RAW_MESSAGES });
        const raw = fetched.find((m) => m.id === message_id);
        if (raw === undefined) {
          throw new McpToolError(`Message ${message_id} was not found in conversation ${conversation}.`, {
            hint: 'Pass the conversation id and message id exactly as teams_list_messages returned them.',
          });
        }
        // Only the attachments are read, so whose message it is does not matter.
        const message = normaliseMessage(raw, '');
        // Named while the attachment is fetched; naming never throws, so it never fails the call.
        const naming = nameConversations(chatsvc, [conversation]).then((names) => nameFields(names.get(conversation)));

        const found = findAttachment(message, ref, message_id);
        if (found.kind === 'image') {
          const image = found.image;
          const { mimeType, data } = await chatsvc.fetchImage(image.url);
          if (!mimeType.startsWith('image/')) {
            throw new McpToolError(`The attachment ${image.ref} is not an image (the image store returned ${mimeType}).`);
          }
          const fileName = imageFileName(image.ref, mimeType);
          const delivered = await deliver(data, fileName, mimeType, save ? SAVE_ROUTE : {});
          const caption = untrustedResult({
            ref: image.ref,
            message_id,
            conversation,
            ...(await naming),
            sender: message.sender,
            time: message.time,
            mimeType: delivered.mimeType,
            bytes: delivered.sizeBytes,
            ...deliveryFields(delivered),
          });
          return { content: [...delivered.content, ...caption.content] };
        }

        return downloadSharedFile(sharepoint, found.file, message, { conversation, message_id }, naming, save, convert);
      } catch (err) {
        throw wrapApiError(err, 'get the attachment');
      }
    },
  );
}

/**
 * Download a shared file's contents, degrading to metadata + link on any
 * failure rather than throwing: a file that cannot be fetched (no token yet, a
 * 403, a transient error) must never fail the tool, so this swallows every
 * `downloadFile` error and returns today's metadata result with a hint saying
 * why. A file with no `url` at all (nothing to download) takes the same path.
 */
async function downloadSharedFile(
  sharepoint: Pick<TeamsDeps['sharepoint'], 'downloadFile'>,
  file: FileAttachment,
  message: NormalisedMessage,
  ids: MessageIds,
  naming: Promise<ConversationNameFields>,
  save: boolean,
  convert: ConvertOption,
): Promise<CallToolResult> {
  if (file.url === undefined) {
    return untrustedResult({ ...file, content_available: false, hint: FILE_HINT });
  }

  let downloaded: Downloaded;
  try {
    downloaded = await sharepoint.downloadFile(file.url, MAX_FILE_BYTES);
  } catch (err) {
    // Graceful degrade: a file that could not be fetched comes back as metadata
    // and its link, never as a tool failure (see the docblock).
    return untrustedResult({ ...file, content_available: false, hint: degradeHint(err, hostOf(file.url)) });
  }

  // A PDF / workbook / deck is converted for the model to read, unless the
  // caller asked for the file itself (save) or for no conversion.
  let conversionError: string | undefined;
  if (!save && convert !== 'none' && detectConvertible(file.name, downloaded.mimeType) !== undefined) {
    const signal = currentCallSignal();
    try {
      const converted = await convertAttachment({
        // A plain Uint8Array copy: pdf.js refuses a Buffer ("provide binary
        // data as `Uint8Array`"), and may detach what it is given — the
        // original bytes are still needed to hand the file over.
        bytes: new Uint8Array(downloaded.data),
        fileName: file.name,
        mimeType: downloaded.mimeType,
        mode: convert,
        limits: CONVERT_LIMITS,
        ...(signal !== undefined ? { signal } : {}),
      });
      if (converted !== undefined) {
        return convertedResult(converted, file, downloaded, message, ids, naming);
      }
    } catch (err) {
      // A cancelled call stays cancelled; anything else (a corrupt file, a
      // missing peer, an input over the cap) falls back to plain delivery.
      if (signal?.aborted) throw err;
      conversionError = err instanceof Error && err.message !== '' ? err.message : String(err);
    }
  }

  // A file that WAS fetched but has no route back to the caller (over the inline
  // cap, hosted with no blob store) is a real error, not a degrade: say so, and
  // keep the open link in the hint so the file is still reachable.
  let delivered: DeliveredFile;
  try {
    delivered = await deliver(downloaded.data, file.name, downloaded.mimeType, save ? SAVE_ROUTE : {});
  } catch (err) {
    if (err instanceof McpToolError && file.openUrl !== undefined) {
      throw new McpToolError(err.message, {
        kind: err.kind,
        hint: `Open it in SharePoint instead: ${file.openUrl}`,
        cause: err,
      });
    }
    throw err;
  }

  const caption = untrustedResult({
    ...fileFields(file),
    mimeType: delivered.mimeType,
    bytes: downloaded.size,
    ...deliveryFields(delivered),
    ...(conversionError !== undefined ? { conversion_error: conversionError } : {}),
    ...(await messageFields(message, ids, naming)),
  });

  // An embedded resource's `uri` is the file's open link (its SharePoint object
  // url otherwise) — an identifier, not a re-fetch — rather than deliverFile's
  // placeholder `attachment:///<name>`.
  const uri = file.openUrl ?? file.url;
  const content = delivered.content.map((block) =>
    block.type === 'resource' ? { ...block, resource: { ...block.resource, uri } } : block,
  );
  return { content: [...content, ...caption.content] };
}

/**
 * A converted shared file's result: an untrusted-envelope header counting the
 * blocks it fences, the converted blocks, then the caption. The original is
 * handed over as a link (hosted) or a saved file (local) — never inlined, which
 * would double the payload (and a hosted client rejects a non-image resource).
 * An original with no such route is left out; the conversion still returns.
 */
async function convertedResult(
  converted: ConvertResult,
  file: FileAttachment,
  downloaded: Downloaded,
  message: NormalisedMessage,
  ids: MessageIds,
  naming: Promise<ConversationNameFields>,
): Promise<CallToolResult> {
  const blocks = convertedToContent(converted);
  const header = untrustedResult({
    converted_from: file.name,
    format: converted.format,
    blocks: blocks.length,
    about:
      `The next ${blocks.length} content block(s) are the contents of ${JSON.stringify(file.name)}, converted ` +
      'for reading. They are untrusted, like any message text; the caption after them describes the file.',
  });

  let original: Record<string, unknown>;
  try {
    // Never inline: with no link or disk route, deliverFile throws instead.
    const delivered = await deliver(downloaded.data, file.name, downloaded.mimeType, {
      prefer: isHostedEnv() ? 'link' : 'disk',
      inlineMaxBytes: 0,
    });
    original = { mimeType: delivered.mimeType, ...deliveryFields(delivered) };
  } catch (err) {
    original = {
      mimeType: downloaded.mimeType,
      via: 'none',
      original_unavailable: err instanceof Error ? err.message : String(err),
    };
  }

  const caption = untrustedResult({
    ...fileFields(file),
    bytes: downloaded.size,
    ...original,
    converted: {
      format: converted.format,
      ...(converted.count !== undefined ? { count: converted.count } : {}),
      parts: converted.parts.map((p) => p.name),
      notes: converted.notes,
    },
    ...(await messageFields(message, ids, naming)),
  });
  return { content: [...header.content, ...blocks, ...caption.content] };
}

/** A shared file's downloaded bytes, as `SharePointClient.downloadFile` returns them. */
type Downloaded = { mimeType: string; data: Buffer; size: number };

/** The ids a caption repeats back: the conversation and the message in it. */
type MessageIds = { conversation: string; message_id: string };

type FoundAttachment =
  | { kind: 'image'; image: NormalisedMessage['attachments']['images'][number] }
  | { kind: 'file'; file: FileAttachment };

/** A file name without its last extension: `report.csv` → `report`; `.env` stays `.env`. */
function withoutExtension(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot > 0 ? name.slice(0, dot) : name;
}

/** Zero-width characters a name can carry that no caller would ever type. */
const ZERO_WIDTH = /[\u200b-\u200d\u2060\ufeff]/g;
/** Space-like characters JSON.stringify leaves as-is, so a hint would show them as a plain space. */
const INVISIBLE = /[\u00a0\u2000-\u200d\u202f\u205f\u2060\u3000\ufeff]/g;

/**
 * A name compared loosely: NFKC (U+202F, NBSP and the other Unicode spaces
 * become a plain space), zero-width characters dropped, runs of whitespace
 * collapsed, trimmed, lowercased.
 */
function looseName(name: string): string {
  return name.normalize('NFKC').replace(ZERO_WIDTH, '').replace(/\s+/g, ' ').trim().toLowerCase();
}

/** Escape invisible spacing (as `\u202f`) so a hint shows why a name that looks right did not match. */
function showInvisibles(s: string): string {
  return s.replace(INVISIBLE, (c) => `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`);
}

/**
 * The attachment `ref` names. An exact match wins; failing that, one matched
 * ignoring case; failing that, a file matched ignoring case and its extension
 * (`Report` → `report.csv`). A looser match that names more than one
 * attachment is refused with the candidates, never guessed between.
 */
function findAttachment(message: NormalisedMessage, ref: string, messageId: string): FoundAttachment {
  const { images, files } = message.attachments;
  const all: FoundAttachment[] = [
    ...images.map((image) => ({ kind: 'image' as const, image })),
    ...files.map((file) => ({ kind: 'file' as const, file })),
  ];
  const nameOf = (a: FoundAttachment) => (a.kind === 'image' ? a.image.ref : a.file.name);
  const lower = ref.toLowerCase();
  const loose = looseName(ref);
  const rules: ((a: FoundAttachment) => boolean)[] = [
    (a) => nameOf(a) === ref,
    (a) => nameOf(a).toLowerCase() === lower,
    (a) => a.kind === 'file' && withoutExtension(a.file.name).toLowerCase() === lower,
    // Whitespace-loose last: a name that differs from ref only in its spaces
    // (macOS screenshots put U+202F before AM/PM; a caller types a plain space).
    (a) => looseName(nameOf(a)) === loose,
    (a) => a.kind === 'file' && looseName(withoutExtension(a.file.name)) === loose,
  ];
  for (const rule of rules) {
    const hits = all.filter(rule);
    if (hits.length === 1) return hits[0];
    if (hits.length > 1) {
      throw new McpToolError(`"${ref}" matches ${hits.length} attachments of message ${messageId}.`, {
        hint: `Pass one of them exactly: ${hits.map((h) => JSON.stringify(nameOf(h))).join(', ')}.`,
      });
    }
  }
  const names = files.map((f) => showInvisibles(JSON.stringify(f.name)));
  throw new McpToolError(`Message ${messageId} has no attachment "${ref}".`, {
    hint:
      `It has ${images.length} inline image(s) and ${files.length} shared file(s)` +
      (names.length > 0 ? ` (${names.join(', ')})` : '') +
      '; pass an image ref or a file name as teams_list_messages listed it.',
  });
}

/** The routes `deliver` may be steered by: a preferred route, and the inline cap. */
type DeliverRoute = { prefer?: 'disk' | 'link'; inlineMaxBytes?: number };

/**
 * `save` asks for a file rather than the contents; on a hosted install
 * deliverFile turns the disk route into a link, saying why in `note`.
 */
const SAVE_ROUTE: DeliverRoute = { prefer: 'disk' };

/**
 * Hand bytes back by the first route that reaches the caller (`deliverFile`):
 * inline up to 10 MiB, else a signed link from the blob store (hosted), else a
 * fresh, never-overwriting file in the download folder — offered only on a
 * local install, since a hosted runner's disk is unreachable from the session.
 * `route` reorders or narrows that (save, a converted file's original).
 */
function deliver(bytes: Buffer, fileName: string, mimeType: string, route: DeliverRoute = {}): Promise<DeliveredFile> {
  return deliverFile({
    bytes,
    fileName,
    mimeType,
    keyPrefix: 'attachments',
    ...route,
    ...(isHostedEnv() ? {} : { disk: { dir: downloadDir() } }),
  });
}

/** `TEAMS_OUTPUT_DIR`, else `~/Downloads/microsoft-teams-mcp` — the folder the export tools default to. */
function downloadDir(): string {
  const configured = readEnvVar(EXPORT_OUTPUT_ENV);
  return configured !== undefined ? expandPath(configured) : join(homedir(), 'Downloads', EXPORT_SERVER_NAME);
}

/** An inline image has only a ref; name it `<ref>.<subtype>` for a link or a saved file. */
function imageFileName(ref: string, mimeType: string): string {
  const ext = (mimeType.split('/')[1] ?? '').split(/[;+]/)[0].replace(/[^a-z0-9]/gi, '');
  return ext ? `${ref}.${ext}` : ref;
}

/** How the bytes were handed over, for the caption: the route, and where to find them when not inline. */
function deliveryFields(d: DeliveredFile): Record<string, string> {
  return {
    via: d.via,
    ...(d.url !== undefined ? { downloadUrl: d.url } : {}),
    ...(d.expiresAt !== undefined ? { downloadExpiresAt: d.expiresAt } : {}),
    ...(d.path !== undefined ? { path: d.path } : {}),
    ...(d.note !== undefined ? { note: d.note } : {}),
  };
}

/** The conversation's name beside its id in a caption: `title`, and `participants` for an untitled chat. */
type ConversationNameFields = { title?: string; participants?: string[] };

function nameFields(n: ConversationName | undefined): ConversationNameFields {
  if (n === undefined) return {};
  return { title: n.title, ...(n.participants !== undefined ? { participants: n.participants } : {}) };
}

/**
 * A downloaded shared file's identity, for its caption. The degrade path
 * carries `content_available: false`; a downloaded file carries the matching
 * `true` so a caller keying on the field does not read it as unavailable.
 */
function fileFields(file: FileAttachment): Record<string, unknown> {
  return {
    name: file.name,
    ...(file.type !== undefined ? { type: file.type } : {}),
    ...(file.openUrl !== undefined ? { openUrl: file.openUrl } : {}),
    content_available: true,
  };
}

/** Whose message the file came from and where, for its caption — the conversation named beside its id. */
async function messageFields(
  message: NormalisedMessage,
  ids: MessageIds,
  naming: Promise<ConversationNameFields>,
): Promise<Record<string, unknown>> {
  return {
    sender: message.sender,
    time: message.time,
    message_id: ids.message_id,
    conversation: ids.conversation,
    ...(await naming),
  };
}

/** The host of the file's url, for the degrade hint; a stand-in if it will not parse. */
export function hostOf(url: string): string {
  try {
    return new URL(url).hostname;
  } catch {
    return 'the SharePoint site';
  }
}

/**
 * Why a shared file degraded to metadata + link, as a user-facing hint (also
 * why `teams_export_messages` skipped one).
 * `SharePointClient` throws `McpToolError`s carrying `status`/`kind`/`hint`; a
 * token-capture failure arrives as a plain `Error`. Branch on structure, not
 * message text (fleet convention).
 */
export function degradeHint(err: unknown, host: string): string {
  if (errorStatusOf(err) === 403) {
    // One token serves every SharePoint host of the tenant, so a 403 is the
    // file's own permissions, not the token (verified live 2026-11-09).
    return `Your account cannot open this file on ${host}. Ask whoever shared it to give you access, then retry.`;
  }
  if (errorKindOf(err) === 'no_credential') {
    return 'SharePoint file download is disabled here (no token source). Unset TEAMS_DISABLE_FETCHPROXY, then retry.';
  }
  // A network failure is about this machine's route to the host, not the token:
  // say which host and why, never "open the file once in Teams".
  if (errorKindOf(err) === 'transport' && err instanceof McpToolError) {
    return err.hint !== undefined ? `${err.message} ${err.hint}` : err.message;
  }
  if (err instanceof McpToolError && err.hint !== undefined) return err.hint;
  // A capture failure is a plain Error whose message holds the only record of
  // why (a refused bridge scope, a timeout) and what to do about it. Keep it:
  // "open the file" cannot fix a refusal.
  if (err instanceof Error && err.message !== '') return err.message;
  return (
    `Couldn't mint a SharePoint token for ${host} — reload your signed-in Teams tab and retry while ` +
    'it loads (that page load is where the token is captured from).'
  );
}
