/**
 * `teams_export_messages` with `attachments: true` — the inline images and
 * shared files of the messages being exported, written into the export beside
 * the conversation files.
 *
 * Each attachment goes to `attachments/<conversation short id>/<message id>-<name>`
 * (an inline image is named `image-<ref>.<ext>`), and the message's record lists
 * it with that path relative to the export root, or with `skipped: <reason>`.
 * The fetching is `teams_get_attachment`'s: an inline image through
 * `chatsvc.fetchImage` (which vets the host before the bearer goes anywhere), a
 * shared file through `sharepoint.downloadFile` capped at the same 25 MB, and a
 * shared file that cannot be fetched is explained with the same
 * {@link degradeHint}. One attachment failing never fails the export; at most
 * {@link ATTACHMENT_CONCURRENCY} downloads run at once.
 */
import { currentCallSignal, mapWithConcurrency, throwIfCancelled } from '@chrischall/mcp-utils';
import { sanitizeExportFileName, type ExportSink } from '@chrischall/mcp-utils/export';
import type { ChatServiceClient } from '../chatsvc.js';
import type { FileAttachment, ImageAttachment, NormalisedMessage } from '../model.js';
import type { SharePointClient } from '../sharepoint.js';
import { degradeHint, hostOf, MAX_FILE_BYTES } from './attachment.js';
import { wrapApiError } from './errors.js';

/** Downloads in flight at once. */
export const ATTACHMENT_CONCURRENCY = 4;

/** What an exported attachment adds to its listing: where it was written, or why it was not. */
export interface ExportedAttachmentFields {
  /** Relative to the export root, `/`-separated. */
  path?: string;
  bytes?: number;
  mimeType?: string;
  skipped?: string;
}

export type ExportedFile = FileAttachment & ExportedAttachmentFields;
export type ExportedImage = ImageAttachment & ExportedAttachmentFields;

/** A message as an export writes it: its attachments carry their export paths (or skip reasons). */
export type ExportMessage = Omit<NormalisedMessage, 'attachments'> & {
  attachments: { files: ExportedFile[]; images: ExportedImage[] };
};

/** One attachment written into the export, for the manifest. */
export interface WrittenAttachment {
  name: string;
  bytes: number;
  conversation: string;
  message_id: string;
}

export interface AttachmentOutcome {
  messages: ExportMessage[];
  written: WrittenAttachment[];
  skipped: number;
}

const IMAGE_EXT: Record<string, string> = { 'image/jpeg': 'jpg', 'image/svg+xml': 'svg' };

/** A file extension for an image MIME type: `image/png` → `png`, `image/jpeg` → `jpg`. */
function imageExt(mimeType: string): string {
  const base = mimeType.split(';')[0].trim().toLowerCase();
  return IMAGE_EXT[base] ?? (base.replace(/^image\//, '').replace(/[^a-z0-9]/g, '') || 'img');
}

/** Why a fetch failed, in one line: the error and its hint. */
function reason(err: unknown, what: string): string {
  const e = wrapApiError(err, what);
  return e.hint !== undefined ? `${e.message} ${e.hint}` : e.message;
}

/**
 * Fetch every attachment of `messages` (one conversation's) into `sink`, and
 * return the messages with each attachment's export path or skip reason.
 * Rejects only when the call is cancelled.
 */
export async function exportAttachments(
  deps: { chatsvc: Pick<ChatServiceClient, 'fetchImage'>; sharepoint: Pick<SharePointClient, 'downloadFile'> },
  sink: Pick<ExportSink, 'addFile'>,
  convShort: string,
  conversationId: string,
  messages: readonly NormalisedMessage[],
): Promise<AttachmentOutcome> {
  const out: ExportMessage[] = messages.map((m) => ({
    ...m,
    attachments: { files: m.attachments.files.map((f) => ({ ...f })), images: m.attachments.images.map((i) => ({ ...i })) },
  }));
  const used = new Set<string>();
  /** A unique relative path for `name` under this conversation's folder. */
  const pathFor = (messageId: string, name: string): string => {
    const stem = sanitizeExportFileName(`${messageId}-${name}`);
    let candidate = stem;
    for (let n = 2; used.has(candidate); n++) {
      const dot = stem.lastIndexOf('.');
      candidate = dot > 0 ? `${stem.slice(0, dot)}-${n}${stem.slice(dot)}` : `${stem}-${n}`;
    }
    used.add(candidate);
    return `attachments/${convShort}/${candidate}`;
  };

  type Job = { message: ExportMessage } & ({ image: ExportedImage } | { file: ExportedFile });
  const jobs: Job[] = [];
  for (const message of out) {
    for (const image of message.attachments.images) jobs.push({ message, image });
    for (const file of message.attachments.files) jobs.push({ message, file });
  }

  const written: WrittenAttachment[] = [];
  let skipped = 0;
  await mapWithConcurrency(jobs, ATTACHMENT_CONCURRENCY, async (job) => {
    throwIfCancelled();
    const target: ExportedAttachmentFields = 'image' in job ? job.image : job.file;
    try {
      let mimeType: string;
      let data: Uint8Array;
      let name: string;
      if ('image' in job) {
        ({ mimeType, data } = await deps.chatsvc.fetchImage(job.image.url));
        if (!mimeType.startsWith('image/')) {
          target.skipped = `Not an image (the image store returned ${mimeType}).`;
          skipped++;
          return;
        }
        name = `image-${job.image.ref}.${imageExt(mimeType)}`;
      } else {
        const url = job.file.url;
        if (url === undefined) {
          target.skipped = 'No download link in the message; open it in Teams with openUrl.';
          skipped++;
          return;
        }
        try {
          ({ mimeType, data } = await deps.sharepoint.downloadFile(url, MAX_FILE_BYTES));
        } catch (err) {
          if (currentCallSignal()?.aborted) throw err;
          target.skipped = degradeHint(err, hostOf(url));
          skipped++;
          return;
        }
        name = job.file.name;
      }
      const rel = pathFor(job.message.id, name);
      const rec = await sink.addFile(rel, data, mimeType);
      target.path = rec.name;
      target.bytes = rec.bytes;
      target.mimeType = mimeType;
      written.push({ name: rec.name, bytes: rec.bytes, conversation: conversationId, message_id: job.message.id });
    } catch (err) {
      if (currentCallSignal()?.aborted) throw err;
      target.skipped = reason(err, 'fetch the attachment');
      skipped++;
    }
  });
  return { messages: out, written, skipped };
}
