/**
 * Normalisation of raw chat-service records into what the tools return.
 *
 * Pure functions only — no I/O, no token — so every rule about what counts as
 * unread, what a conversation is called, and how a message body reads as text
 * is tested in isolation from the client. Shapes are the ones recorded in the
 * "Verified live facts" section of `docs/plans/2026-10-10-chatsvc-api.md`.
 */
import type { RawConversation, RawLastMessage, RawMessage } from './chatsvc.js';

export type ConversationKind = 'channel' | 'meeting' | 'oneOnOne' | 'group' | 'self';

/**
 * The user's own notes chat ("notes to self") — the one `48:*` system
 * conversation that is a real chat. Every other `48:*` id (calendar,
 * notifications, …) is not one and is never listed.
 */
export const SELF_CHAT = '48:notes';
const SELF_CHAT_TITLE = 'Notes (self-chat)';

/** Names that resolve to {@link SELF_CHAT} outright, ahead of any substring match. */
export const SELF_CHAT_ALIASES = new Set(['notes', 'self']);

export interface ConversationSummary {
  id: string;
  kind: ConversationKind;
  title: string;
  /** A channel's team name, when the team's own thread is in the listing. */
  team?: string;
  /**
   * A HINT from the listing alone: there is activity past the read horizon
   * (or no horizon at all) that is not plainly the user's own. The listing
   * cannot say whether that activity is a real message — `lastimreceivedtime`
   * moves for call recordings and other system events too, and `lastMessage`
   * can be stale — so `unread` is only exact once `unread_confirmed` is true,
   * i.e. after the messages past the horizon were read (see
   * `tools/confirm.ts`).
   */
  unread: boolean;
  /** True once a read of the messages past the horizon settled `unread`. */
  unread_confirmed: boolean;
  /**
   * A channel the user has hidden in Teams (`memberProperties.interest` is
   * `NotInterested`). Teams shows it nowhere, unread or not.
   */
  hidden?: true;
  /**
   * Muted in Teams (`properties.alerts` is `"false"`; live 2026-10-12: 8 of
   * 1,435 conversations, meeting chats and a 1:1, every other one `"true"` or
   * absent). Teams sends no notification for it; absent otherwise.
   */
  muted?: true;
  /** Marked a favorite in Teams (`properties.favorite`); absent otherwise. */
  favorite?: true;
  /** Pinned in Teams (`properties.ispinned`); absent otherwise. */
  pinned?: true;
  /** ISO time of the last message; absent for a conversation with none. */
  lastActivity?: string;
  lastMessagePreview?: string;
  lastSender?: string;
  /**
   * An untitled group chat's distinct recent senders other than the user,
   * most recent first — its title names the first three (see
   * `tools/titles.ts`) — or an untitled 1:1's other person. Absent until the
   * listing or a read of its messages found someone.
   */
  participants?: string[];
  /** With a query: why this conversation matched it (see {@link matchConversations}). */
  match_reason?: MatchReason;
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
  /** Reactions with at least one reactor; absent when there are none. */
  reactions?: Reaction[];
}

export interface Reaction {
  /** `like`, `heart`, … or a unicode-style key such as `1f389_partypopper`. */
  key: string;
  count: number;
  /** Whether the signed-in user is one of the reactors. */
  byMe: boolean;
}

/** The message fields the unread test reads — a full message or a listing's `lastMessage`. */
interface MessageLike {
  id?: string;
  from?: string;
  messagetype?: string;
  properties?: RawMessage['properties'];
}

/**
 * Message types a person wrote. Everything else is a system event that Teams
 * never counts as unread: `RichText/Media_*` (call recordings, transcripts),
 * `Event/*`, `ThreadActivity/*`, `Control/*`, ….
 */
const USER_MESSAGE_TYPES = new Set(['Text', 'RichText/Html']);

/** Longest last-message preview a conversation summary carries. */
const PREVIEW_CHARS = 200;

/** The preview of a last message that is a system event (a call, a member added, …). */
const SYSTEM_EVENT_PREVIEW = '(system event)';

/** An untopiced group chat's title when no one but the user is known to have written in it. */
export const UNTITLED_GROUP_TITLE = '(untitled group chat)';

/** An untopiced 1:1's title when the other person's name is not known. */
export const UNTITLED_ONE_ON_ONE_TITLE = '(1:1 chat)';

/** What an untopiced group chat's participant title starts with (see `tools/titles.ts`). */
export const GROUP_TITLE_PREFIX = 'Group chat: ';

/** A channel's team name when the team's own thread is listed but names nothing. */
const UNNAMED_TEAM = '(team)';

/** Inline images live on Skype's AMS object store; only those are attachment refs. */
const IMAGE_HOST_SUFFIX = '.asm.skype.com';
const IMAGE_OBJECT_PATH = /^\/v1\/objects\/([^/]+)\/views\//;

export function conversationKind(c: RawConversation): ConversationKind {
  const id = c.id;
  if (id === SELF_CHAT) return 'self';
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
 * `properties.lastimreceivedtime` as epoch ms. Despite the name it moves for
 * system events too (a call recording sets it to that message's id), so it
 * says only that SOMETHING arrived, not that a real message did. Seen as an ISO string; epoch ms (a number or a
 * numeric string) is read too. `undefined` when absent or unparseable.
 */
export function lastImReceivedMs(c: RawConversation): number | undefined {
  const raw: unknown = c.properties?.lastimreceivedtime;
  let ms: number;
  if (typeof raw === 'number') ms = raw;
  else if (typeof raw === 'string' && /^\s*\d+\s*$/.test(raw)) ms = Number(raw);
  else if (typeof raw === 'string' && raw.trim() !== '') ms = Date.parse(raw);
  else return undefined;
  return Number.isFinite(ms) && ms > 0 ? ms : undefined;
}

/**
 * The conversation's newest activity of any kind (epoch ms): the later of the
 * last message's arrival (its id) and `lastimreceivedtime`, since either can
 * lag the other. `undefined` when the listing carries neither. Nothing real
 * can be unread after this, so it bounds how recent an unread message can be.
 */
export function newestActivityMs(c: RawConversation): number | undefined {
  const lastArrival = c.lastMessage !== undefined ? Number(c.lastMessage.id) : NaN;
  const signals = [lastArrival, lastImReceivedMs(c)].filter((n): n is number => n !== undefined && Number.isFinite(n) && n > 0);
  return signals.length === 0 ? undefined : Math.max(...signals);
}

/**
 * Whether the conversation is an unread CANDIDATE: its newest activity — the
 * later of `lastimreceivedtime` and the last message's arrival, of any type —
 * is past the read horizon (0 for a conversation never opened), and is not
 * the user's own doing.
 *
 * Only a candidate. Measured live (2026-10-10): `lastimreceivedtime` equals
 * the last message's id even when that is a call recording, and `lastMessage`
 * can be stale (older than system events after it), so neither can say a real
 * message is waiting. The tools confirm a candidate by reading the messages
 * past the horizon. A false here is reliable: nothing at all came after the
 * horizon, or the newest thing is the user's own message.
 */
export function isUnreadCandidate(c: RawConversation, meMri: string): boolean {
  const horizon = readHorizon(c);
  const last = c.lastMessage;
  const lastArrival = last !== undefined ? Number(last.id) : NaN;
  const newest = newestActivityMs(c);
  if (newest === undefined || newest <= horizon) return false;
  // The newest thing is the user's own doing: they were there for it.
  const lastIsMine = last !== undefined && senderMri(last.from) === meMri;
  return !(lastIsMine && lastArrival >= newest);
}

/**
 * The other person's MRI in a 1:1: its id is `19:<oid>_<oid>@unq.gbl.spaces`,
 * the two people's AAD object ids, so the one that is not the user's is
 * theirs (`8:orgid:<oid>`). `undefined` when the id is not shaped so or does
 * not name the user — then the listing cannot say who the other person is.
 */
export function oneOnOnePeerMri(id: string, meMri: string): string | undefined {
  const m = /^19:([^_@]+)_([^_@]+)@unq\.gbl\.spaces$/.exec(id);
  const prefix = '8:orgid:';
  if (m === null || !meMri.startsWith(prefix)) return undefined;
  const mine = meMri.slice(prefix.length).toLowerCase();
  const [a, b] = [m[1], m[2]];
  if (a.toLowerCase() === mine && b.toLowerCase() !== mine) return `${prefix}${b}`;
  if (b.toLowerCase() === mine && a.toLowerCase() !== mine) return `${prefix}${a}`;
  return undefined;
}

/** True when the sender's MRI is `mri`, ignoring case (oids are GUIDs). */
export function isFromMri(from: string | undefined, mri: string): boolean {
  return senderMri(from)?.toLowerCase() === mri.toLowerCase();
}

/** A `properties` flag Teams sends as `"true"` / `"True"` (or, defensively, a boolean). */
function flag(c: RawConversation, name: string): boolean {
  const v: unknown = c.properties?.[name];
  return v === true || (typeof v === 'string' && v.trim().toLowerCase() === 'true');
}

/** A `properties` flag Teams sends as `"false"` / `"False"` (or a boolean) — absent is not false. */
function flagOff(c: RawConversation, name: string): boolean {
  const v: unknown = c.properties?.[name];
  return v === false || (typeof v === 'string' && v.trim().toLowerCase() === 'false');
}

/** A channel the user has hidden in Teams. */
function isHiddenChannel(c: RawConversation): boolean {
  return c.memberProperties?.interest === 'NotInterested';
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

  let title = kind === 'self' ? SELF_CHAT_TITLE : ownTeam ? 'General' : nonEmpty(tp.topic);
  let participants: string[] | undefined;
  if (title === undefined && kind === 'oneOnOne') {
    // A 1:1 is named after the other person, whom its id names. The listing
    // knows their name only when they sent the last message; otherwise
    // `tools/titles.ts` finds it in their recent messages.
    const peer = oneOnOnePeerMri(c.id, meMri);
    const fromPeer = peer !== undefined ? last !== undefined && isFromMri(last.from, peer) : !lastFromMe;
    const name = fromPeer ? lastSender : undefined;
    title = name ?? UNTITLED_ONE_ON_ONE_TITLE;
    if (name !== undefined) participants = [name];
  } else if (title === undefined) {
    // An untopiced group chat's last sender is only a stand-in: `tools/titles.ts`
    // names it after its recent participants instead.
    const untitled = kind === 'group' ? UNTITLED_GROUP_TITLE : '(untitled chat)';
    title = kind === 'channel' ? 'General' : (!lastFromMe ? lastSender : undefined) ?? untitled;
  }

  let team: string | undefined;
  if (kind === 'channel') {
    const teamThread = ownTeam ? c : tp.spaceId ? all.find((x) => x.id === tp.spaceId) : undefined;
    if (teamThread !== undefined) team = teamName(teamThread);
  }

  const lastActivity = last ? (toIso(last.originalarrivaltime) ?? toIso(last.composetime)) : undefined;
  const preview = systemEvent ? SYSTEM_EVENT_PREVIEW : last?.content !== undefined ? htmlToText(last.content) : undefined;
  const unread = isUnreadCandidate(c, meMri);
  const hidden = kind === 'channel' && isHiddenChannel(c);

  return {
    id: c.id,
    kind,
    title,
    ...(team !== undefined ? { team } : {}),
    unread,
    unread_confirmed: false,
    ...(hidden ? { hidden: true as const } : {}),
    ...(flagOff(c, 'alerts') ? { muted: true as const } : {}),
    ...(flag(c, 'favorite') ? { favorite: true as const } : {}),
    ...(flag(c, 'ispinned') ? { pinned: true as const } : {}),
    ...(lastActivity !== undefined ? { lastActivity } : {}),
    ...(preview ? { lastMessagePreview: truncate(preview, PREVIEW_CHARS) } : {}),
    ...(lastSender !== undefined ? { lastSender } : {}),
    ...(participants !== undefined ? { participants } : {}),
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

/**
 * `properties.emotions`: a JSON string (or, sometimes, an already-parsed
 * array) of `{ key, users: [{ mri, time }] }`. Keys nobody holds any more are
 * dropped; anything unreadable yields no reactions rather than an error.
 */
function reactionsOf(raw: unknown, meMri: string): Reaction[] {
  let parsed: unknown = raw;
  if (typeof raw === 'string') {
    try {
      parsed = JSON.parse(raw);
    } catch {
      return [];
    }
  }
  if (!Array.isArray(parsed)) return [];
  const out: Reaction[] = [];
  for (const e of parsed) {
    if (e === null || typeof e !== 'object') continue;
    const { key, users } = e as { key?: unknown; users?: unknown };
    if (typeof key !== 'string' || key === '' || !Array.isArray(users) || users.length === 0) continue;
    const byMe = users.some((u) => (u as { mri?: unknown } | null)?.mri === meMri);
    out.push({ key, count: users.length, byMe });
  }
  return out;
}

/**
 * Messages oldest first, by arrival (a message id IS its arrival ms).
 *
 * The service pages newest first but does not strictly order within a page:
 * live, a quote-reply sent 6s after its original came back ahead of it, so a
 * plain reverse() listed the reply first. Ids that are not numbers sort by
 * originalarrivaltime instead, and keep their page order when that is missing.
 */
export function oldestFirst<T extends Pick<RawMessage, 'id' | 'originalarrivaltime'>>(messages: readonly T[]): T[] {
  const at = (m: T): number => {
    const id = Number(m.id);
    if (Number.isFinite(id)) return id;
    const t = Date.parse(m.originalarrivaltime ?? '');
    return Number.isNaN(t) ? Number.NaN : t;
  };
  return messages
    .map((m, i) => ({ m, i, t: at(m) }))
    .reverse()
    .sort((a, b) => (Number.isNaN(a.t) || Number.isNaN(b.t) ? 0 : a.t - b.t))
    .map(({ m }) => m);
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
  const reactions = reactionsOf(props.emotions, meMri);

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
    ...(reactions.length > 0 ? { reactions } : {}),
  };
}

/**
 * Every conversation's summary, with the self-chat added when the listing did
 * not carry it — so `48:notes` resolves by id or name whether or not Teams
 * listed it.
 */
export function summariseAll(raw: readonly RawConversation[], meMri: string): ConversationSummary[] {
  const out = raw.map((c) => conversationSummary(c, meMri, raw));
  if (!out.some((s) => s.id === SELF_CHAT)) out.push(conversationSummary({ id: SELF_CHAT }, meMri));
  return out;
}

/** Why a conversation matched a query, in the order the reasons are tried. */
export type MatchReason = 'id' | 'title' | 'team' | 'participant' | 'message';

export interface MatchOptions {
  /**
   * Also match the listing's last-message preview when nothing else does
   * (`match_reason: "message"`). For the list tools only: a name given to a
   * tool that writes must not resolve to a chat because of what was said in it.
   */
  messages?: boolean;
}

type Matchable = Pick<ConversationSummary, 'id' | 'title' | 'team' | 'participants' | 'lastMessagePreview'>;

/**
 * Resolve a user's `conversation` argument: an exact id wins outright, then
 * `notes` / `self` name the self-chat, else a case-insensitive substring of
 * the title, team name, or one of an untitled group chat's participants (and,
 * with `messages`, the last-message preview). A title this server made up —
 * "Group chat: …", "(untitled group chat)", or an untitled 1:1's other person
 * — is not itself matched, so "group" or "chat" does not hit every untitled
 * group chat: such a chat matches by its participants instead. Each match is a
 * copy carrying its {@link MatchReason} — the first that applies, in order.
 */
export function matchConversations<T extends Matchable>(
  summaries: readonly T[],
  query: string,
  { messages = false }: MatchOptions = {},
): (T & { match_reason: MatchReason })[] {
  const exact = summaries.find((s) => s.id === query);
  if (exact) return [{ ...exact, match_reason: 'id' }];
  const q = query.trim().toLowerCase();
  if (q === '') return [];
  if (SELF_CHAT_ALIASES.has(q)) {
    const self = summaries.find((s) => s.id === SELF_CHAT);
    if (self) return [{ ...self, match_reason: 'title' }];
  }
  return summaries.flatMap((s) => {
    const reason = matchReason(s, q, messages);
    return reason === undefined ? [] : [{ ...s, match_reason: reason }];
  });
}

/** The first reason `s` matches the lower-cased query `q`, or `undefined`. */
function matchReason(s: Matchable, q: string, messages: boolean): MatchReason | undefined {
  if (hasRealTitle(s) && s.title.toLowerCase().includes(q)) return 'title';
  if (s.team?.toLowerCase().includes(q)) return 'team';
  if (s.participants?.some((p) => p.toLowerCase().includes(q))) return 'participant';
  if (messages && s.lastMessagePreview !== undefined && s.lastMessagePreview !== SYSTEM_EVENT_PREVIEW) {
    if (s.lastMessagePreview.toLowerCase().includes(q)) return 'message';
  }
  return undefined;
}

/**
 * False for a title made from participants (an untitled group chat's, or an
 * untitled 1:1's other person — only such chats carry `participants`), or an
 * untitled-group / untitled-1:1 stand-in.
 */
function hasRealTitle(s: Pick<ConversationSummary, 'title' | 'participants'>): boolean {
  if (s.title === UNTITLED_GROUP_TITLE || s.title === UNTITLED_ONE_ON_ONE_TITLE) return false;
  return s.participants === undefined;
}
