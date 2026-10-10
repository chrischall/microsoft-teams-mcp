/**
 * Normalisation of raw chat-service records into what the tools return.
 *
 * Pure functions only — no I/O, no token — so every rule about what counts as
 * unread, what a conversation is called, and how a message body reads as text
 * is tested in isolation from the client. Shapes are the ones recorded in the
 * "Verified live facts" section of `docs/plans/2026-10-10-chatsvc-api.md`.
 */
import type { RawConversation, RawLastMessage, RawMessage } from './chatsvc.js';

export type ConversationKind = 'channel' | 'meeting' | 'oneOnOne' | 'group';

export interface ConversationSummary {
  id: string;
  kind: ConversationKind;
  title: string;
  /** A channel's team name, when the team's own thread is in the listing. */
  team?: string;
  unread: boolean;
  /**
   * Set when `unread` rests on a last message that is a system event or
   * deleted: a real unread message may or may not sit behind it, so the read
   * tools confirm by fetching.
   */
  unread_uncertain?: true;
  /** ISO time of the last message; absent for a conversation with none. */
  lastActivity?: string;
  lastMessagePreview?: string;
  lastSender?: string;
}

export interface FileAttachment {
  name: string;
  type?: string;
  /** Opens the file in SharePoint/Office for the user. */
  openUrl?: string;
  /** The file's own SharePoint URL. The chat-service token cannot download it. */
  url?: string;
}

export interface ImageAttachment {
  /** The image-store object id — what `teams_get_attachment` takes as `ref`. */
  ref: string;
  url: string;
}

export interface NormalisedMessage {
  id: string;
  time: string;
  sender: string;
  fromMe: boolean;
  text: string;
  subject?: string;
  /** In a channel: the root post this replies to (absent on the root itself). */
  replyTo?: string;
  edited: boolean;
  deleted?: true;
  attachments: { files: FileAttachment[]; images: ImageAttachment[] };
}

/** The message fields the unread test reads — a full message or a listing's `lastMessage`. */
interface MessageLike {
  id?: string;
  from?: string;
  messagetype?: string;
  properties?: RawMessage['properties'];
}

/** Message types a person wrote. Everything else (`ThreadActivity/*`, `Event/Call`, …) is a system event. */
const USER_MESSAGE_TYPES = new Set(['Text', 'RichText/Html']);

/** Longest last-message preview a conversation summary carries. */
const PREVIEW_CHARS = 200;

/** The preview of a last message that is a system event (a call, a member added, …). */
const SYSTEM_EVENT_PREVIEW = '(system event)';

/** A channel's team name when the team's own thread is listed but names nothing. */
const UNNAMED_TEAM = '(team)';

/** Inline images live on Skype's AMS object store; only those are attachment refs. */
const IMAGE_HOST_SUFFIX = '.asm.skype.com';
const IMAGE_OBJECT_PATH = /^\/v1\/objects\/([^/]+)\/views\//;

export function conversationKind(c: RawConversation): ConversationKind {
  const id = c.id;
  if (id.endsWith('@thread.tacv2')) return 'channel';
  const tp = c.threadProperties ?? {};
  if (
    id.startsWith('19:meeting_') ||
    /meeting/i.test(tp.productThreadType ?? '') ||
    /meeting/i.test(tp.threadType ?? '')
  ) {
    return 'meeting';
  }
  if (id.endsWith('@unq.gbl.spaces')) return 'oneOnOne';
  return 'group';
}

/**
 * The arrival time (epoch ms) of the last message the user has read: the first
 * segment of `consumptionhorizon` (`"<lastReadArrivalMs>;<clientMs>;<messageId>"`).
 * 0 — everything unread — when absent or unparseable.
 */
export function readHorizon(c: RawConversation): number {
  const raw = c.properties?.consumptionhorizon;
  if (typeof raw !== 'string') return 0;
  const n = Number(raw.split(';')[0]);
  return Number.isFinite(n) && n > 0 ? n : 0;
}

/** The sender's MRI (`8:orgid:<oid>`): `from` is a URL ending in it. */
export function senderMri(from: string | undefined): string | undefined {
  if (typeof from !== 'string' || from === '') return undefined;
  const mri = from.slice(from.lastIndexOf('/') + 1);
  return mri === '' ? undefined : mri;
}

/** True for a message a person wrote, false for a system event. */
export function isUserMessage(m: { messagetype?: string }): boolean {
  return USER_MESSAGE_TYPES.has(m.messagetype ?? '');
}

function isDeleted(m: { properties?: RawMessage['properties'] }): boolean {
  const t = m.properties?.deletetime;
  return t !== undefined && t !== null && t !== '' && t !== 0;
}

/**
 * Unread = arrived after the read horizon (a message id IS its arrival ms),
 * written by someone other than the user, a real message, and not deleted.
 */
export function isUnread(m: MessageLike, horizon: number, meMri: string): boolean {
  const arrival = Number(m.id);
  if (!Number.isFinite(arrival) || arrival <= horizon) return false;
  if (senderMri(m.from) === meMri) return false;
  return isUserMessage(m) && !isDeleted(m);
}

function toIso(value: string | undefined): string | undefined {
  if (typeof value !== 'string' || value === '') return undefined;
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? undefined : d.toISOString();
}

function nonEmpty(s: string | undefined): string | undefined {
  const t = s?.trim();
  return t ? t : undefined;
}

/** A team's thread: its own `spaceId`, or a `space` thread type. */
function isTeamThread(c: RawConversation): boolean {
  const tp = c.threadProperties ?? {};
  return tp.spaceId === c.id || tp.threadType === 'space';
}

/**
 * A team's name from its own thread: the topic, else the space/topic-thread
 * topic Teams sometimes carries instead, else {@link UNNAMED_TEAM}.
 */
function teamName(team: RawConversation): string {
  const tp = team.threadProperties ?? {};
  return nonEmpty(tp.topic) ?? nonEmpty(tp.spaceThreadTopic) ?? nonEmpty(tp.topicThreadTopic) ?? UNNAMED_TEAM;
}

/**
 * Whether the listing's `lastMessage` makes the conversation unread. A real,
 * live last message decides it outright. A system event or a deleted message
 * hides what is behind it, so one past the horizon (and not the user's own)
 * counts as unread but uncertain — the read tools confirm by fetching.
 */
function lastMessageUnread(
  last: RawLastMessage | undefined,
  horizon: number,
  meMri: string,
): { unread: boolean; uncertain: boolean } {
  if (last === undefined) return { unread: false, uncertain: false };
  if (isUserMessage(last) && !isDeleted(last)) return { unread: isUnread(last, horizon, meMri), uncertain: false };
  const arrival = Number(last.id);
  const unread = Number.isFinite(arrival) && arrival > horizon && senderMri(last.from) !== meMri;
  return { unread, uncertain: unread };
}

export function conversationSummary(
  c: RawConversation,
  meMri: string,
  all: readonly RawConversation[] = [],
): ConversationSummary {
  const kind = conversationKind(c);
  const tp = c.threadProperties ?? {};
  const last: RawLastMessage | undefined = c.lastMessage;
  const systemEvent = last !== undefined && !isUserMessage(last);
  const lastFromMe = last !== undefined && senderMri(last.from) === meMri;
  // A system event's display name is whoever triggered it, not a sender.
  const lastSender = systemEvent ? undefined : nonEmpty(last?.imdisplayname);
  const ownTeam = kind === 'channel' && isTeamThread(c);

  let title = ownTeam ? 'General' : nonEmpty(tp.topic);
  if (title === undefined) {
    // An untopiced chat is named after the other person — learnable only when
    // they sent the last message.
    title = kind === 'channel' ? 'General' : (!lastFromMe ? lastSender : undefined) ?? '(untitled chat)';
  }

  let team: string | undefined;
  if (kind === 'channel') {
    const teamThread = ownTeam ? c : tp.spaceId ? all.find((x) => x.id === tp.spaceId) : undefined;
    if (teamThread !== undefined) team = teamName(teamThread);
  }

  const lastActivity = last ? (toIso(last.originalarrivaltime) ?? toIso(last.composetime)) : undefined;
  const preview = systemEvent ? SYSTEM_EVENT_PREVIEW : last?.content !== undefined ? htmlToText(last.content) : undefined;
  const { unread, uncertain } = lastMessageUnread(last, readHorizon(c), meMri);

  return {
    id: c.id,
    kind,
    title,
    ...(team !== undefined ? { team } : {}),
    unread,
    ...(uncertain ? { unread_uncertain: true as const } : {}),
    ...(lastActivity !== undefined ? { lastActivity } : {}),
    ...(preview ? { lastMessagePreview: truncate(preview, PREVIEW_CHARS) } : {}),
    ...(lastSender !== undefined ? { lastSender } : {}),
  };
}

function truncate(s: string, max: number): string {
  return s.length <= max ? s : `${s.slice(0, max - 1)}…`;
}

/*
 * Placeholders that survive tag-stripping, so blockquote depth can be applied
 * after the text is decoded. Control characters never come out of entity
 * decoding (it drops them), so a message cannot forge one.
 */
const QUOTE_OPEN = '\u0001';
const QUOTE_CLOSE = '\u0002';

const NAMED_ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: ' ',
};

function decodeEntities(s: string): string {
  // One pass, so `&amp;lt;` becomes `&lt;` and not `<`.
  return s.replace(/&(?:#(\d+)|#x([0-9a-f]+)|([a-z]+));/gi, (whole, dec, hex, name) => {
    if (name !== undefined) return NAMED_ENTITIES[name.toLowerCase()] ?? whole;
    const code = dec !== undefined ? Number(dec) : parseInt(hex, 16);
    if (!Number.isFinite(code) || code > 0x10ffff) return whole;
    // Drop control characters (keeping tab/newline): they carry no text and
    // would collide with the quote placeholders.
    if (code < 0x20 && code !== 0x09 && code !== 0x0a) return '';
    return String.fromCodePoint(code);
  });
}

function attr(tag: string, name: string): string | undefined {
  const m = new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`, 'i').exec(tag);
  return m ? (m[1] ?? m[2]) : undefined;
}

const stripTags = (s: string): string => s.replace(/<\/?[a-z!][^>]*>/gi, '');

/**
 * A Teams message body as readable plain text. Links keep their target as
 * `text (url)`, inline images become `[image]` (emoji become their alt text),
 * and a quoted reply's lines are prefixed `> `. No HTML library: Teams bodies
 * are a small, regular subset.
 */
export function htmlToText(html: string): string {
  let s = html.replace(/[\u0001\u0002]/g, '');
  s = s.replace(/<(script|style)\b[\s\S]*?<\/\1\s*>/gi, '');
  s = s.replace(/<blockquote\b[^>]*>/gi, `\n${QUOTE_OPEN}\n`).replace(/<\/blockquote\s*>/gi, `\n${QUOTE_CLOSE}\n`);
  s = s.replace(/<img\b[^>]*>/gi, (tag) => {
    const alt = attr(tag, 'alt');
    return /schema\.skype\.com\/Emoji/i.test(attr(tag, 'itemtype') ?? '') && alt ? alt : '[image]';
  });
  s = s.replace(/<a\b([^>]*)>([\s\S]*?)<\/a\s*>/gi, (_whole, attrs: string, inner: string) => {
    const href = attr(` ${attrs}`, 'href');
    const label = stripTags(inner).trim();
    if (!href) return inner;
    if (label === '' || label === href) return href;
    return `${inner} (${href})`;
  });
  // Block elements start and end a line (a quote's author <strong> runs
  // straight into its <p> otherwise); the blank lines this makes collapse below.
  s = s.replace(/<br\s*\/?>/gi, '\n').replace(/<\/?(p|div|li|tr|h[1-6])\b[^>]*>/gi, '\n');
  s = decodeEntities(stripTags(s));

  const out: string[] = [];
  let depth = 0;
  for (const raw of s.split(/\r?\n/)) {
    const line = raw.replace(/[ \t ]+/g, ' ').trim();
    if (line === QUOTE_OPEN) depth++;
    else if (line === QUOTE_CLOSE) depth = Math.max(0, depth - 1);
    else if (line !== '') out.push(`${'> '.repeat(depth)}${line}`);
  }
  return out.join('\n');
}

function fileAttachments(raw: unknown): FileAttachment[] {
  let parsed: unknown = raw;
  if (typeof raw === 'string') {
    try {
      parsed = JSON.parse(raw);
    } catch {
      return [];
    }
  }
  if (!Array.isArray(parsed)) return [];
  const out: FileAttachment[] = [];
  for (const f of parsed) {
    if (f === null || typeof f !== 'object') continue;
    const o = f as Record<string, unknown>;
    const info = (o.fileInfo && typeof o.fileInfo === 'object' ? o.fileInfo : {}) as Record<string, unknown>;
    const str = (v: unknown): string | undefined => (typeof v === 'string' && v !== '' ? v : undefined);
    const name = str(o.fileName) ?? str(o.title) ?? '(file)';
    const type = str(o.fileType) ?? str(o.type);
    const openUrl = str(o.openUrl) ?? str(info.shareUrl);
    const url = str(o.objectUrl) ?? str(info.fileUrl);
    out.push({
      name,
      ...(type !== undefined ? { type } : {}),
      ...(openUrl !== undefined ? { openUrl } : {}),
      ...(url !== undefined ? { url } : {}),
    });
  }
  return out;
}

function imageAttachments(html: string): ImageAttachment[] {
  const out: ImageAttachment[] = [];
  const seen = new Set<string>();
  for (const [tag] of html.matchAll(/<img\b[^>]*>/gi)) {
    const src = attr(tag, 'src');
    if (!src) continue;
    let u: URL;
    try {
      u = new URL(decodeEntities(src));
    } catch {
      continue;
    }
    if (u.protocol !== 'https:' || !u.hostname.endsWith(IMAGE_HOST_SUFFIX)) continue;
    const ref = IMAGE_OBJECT_PATH.exec(u.pathname)?.[1];
    if (!ref || seen.has(ref)) continue;
    seen.add(ref);
    out.push({ ref, url: u.toString() });
  }
  return out;
}

export function normaliseMessage(m: RawMessage, meMri: string): NormalisedMessage {
  const props = m.properties ?? {};
  const mri = senderMri(m.from);
  const arrival = Number(m.id);
  const time =
    toIso(m.originalarrivaltime) ??
    toIso(m.composetime) ??
    (Number.isFinite(arrival) ? new Date(arrival).toISOString() : '');
  const subject = typeof props.subject === 'string' ? nonEmpty(props.subject) : undefined;
  const root = /;messageid=(\d+)/.exec(m.conversationLink ?? '')?.[1];
  const content = m.content ?? '';

  return {
    id: m.id,
    time,
    sender: nonEmpty(m.imdisplayname) ?? mri ?? 'unknown',
    fromMe: mri !== undefined && mri === meMri,
    text: htmlToText(content),
    ...(subject !== undefined ? { subject } : {}),
    ...(root !== undefined && root !== m.id ? { replyTo: root } : {}),
    edited: props.edittime !== undefined && props.edittime !== null && props.edittime !== '',
    ...(isDeleted(m) ? { deleted: true as const } : {}),
    attachments: { files: fileAttachments(props.files), images: imageAttachments(content) },
  };
}

/**
 * Resolve a user's `conversation` argument: an exact id wins outright, else a
 * case-insensitive substring of the title or team name.
 */
export function matchConversations<T extends Pick<ConversationSummary, 'id' | 'title' | 'team'>>(
  summaries: readonly T[],
  query: string,
): T[] {
  const exact = summaries.find((s) => s.id === query);
  if (exact) return [exact];
  const q = query.trim().toLowerCase();
  if (q === '') return [];
  return summaries.filter((s) => s.title.toLowerCase().includes(q) || (s.team?.toLowerCase().includes(q) ?? false));
}
