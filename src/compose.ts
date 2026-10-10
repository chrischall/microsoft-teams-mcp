/**
 * Building an outgoing Teams message — pure functions, no I/O.
 *
 * Every shape here is the one captured from Teams' own requests (see "Verified
 * write shapes" in `docs/plans/2026-10-10-chatsvc-api.md`), except the mention
 * markup, which follows the documented Skype format and is NOT yet captured
 * live. The caller's text is always escaped: the only markup a message can
 * carry is what these builders emit (paragraph, line breaks, mention spans,
 * emoji spans, the quote block), so a caller cannot smuggle HTML into a post.
 */
import { randomInt } from 'node:crypto';
import { McpToolError } from '@chrischall/mcp-utils';

const HTML_ESCAPES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
};

export function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => HTML_ESCAPES[c]);
}

/**
 * Teams' short legacy emoji keys this module can render, with the title Teams
 * shows and the character used as the `alt` fallback. The keys are the
 * observed ones; the titles and characters are best-effort (they only affect
 * hover text and copy-paste, never which emoji renders).
 */
const NAMED_EMOJI: Record<string, { title: string; char: string }> = {
  like: { title: 'Like', char: '👍' },
  heart: { title: 'Heart', char: '❤️' },
  laugh: { title: 'Laugh', char: '😆' },
  surprised: { title: 'Surprised', char: '😮' },
  sad: { title: 'Sad', char: '😢' },
  angry: { title: 'Angry', char: '😡' },
  smile: { title: 'Smile', char: '🙂' },
  wink: { title: 'Wink', char: '😉' },
  cool: { title: 'Cool', char: '😎' },
  hearteyes: { title: 'Heart eyes', char: '😍' },
  fire: { title: 'Fire', char: '🔥' },
  party: { title: 'Party', char: '🎉' },
};

/**
 * A unicode-style key: lowercase hex code point(s), `_`, the character's name
 * squashed to `[a-z0-9]` (`1f389_partypopper`). Multi-code-point keys are
 * assumed `-`-joined — unverified, so such a key at worst renders as Teams'
 * missing-emoji image.
 */
const UNICODE_KEY = /^([0-9a-f]{4,6}(?:-[0-9a-f]{4,6})*)_([a-z0-9]+)$/;

/**
 * The captured inline-emoji span for `key`, or `undefined` when the key is
 * neither a known short name nor unicode-style — the caller then leaves the
 * `:key:` as plain text (so "10:30:45" is never mangled).
 */
export function emojiSpan(key: string): string | undefined {
  let title: string;
  let char: string;
  const named = NAMED_EMOJI[key];
  if (named !== undefined) {
    ({ title, char } = named);
  } else {
    const m = UNICODE_KEY.exec(key);
    if (m === null) return undefined;
    const points = m[1].split('-').map((h) => parseInt(h, 16));
    if (points.some((p) => p > 0x10ffff)) return undefined;
    char = String.fromCodePoint(...points);
    title = m[2];
  }
  // `key` is [a-z0-9_-] by construction here, so it is safe in a URL and attribute.
  const t = escapeHtml(title);
  return (
    `<span contenteditable="false" title="${t}" type="(${key})" class="animated-emoticon-20-${key}" itemscope>` +
    `<img itemscope itemtype="http://schema.skype.com/Emoji" itemid="${key}" ` +
    `src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/${key}/default/20_f.png" ` +
    `title="${t}" alt="${escapeHtml(char)}" style="width:20px;height:20px;"></span>`
  );
}

/**
 * Plain text → escaped HTML, newlines as `<br>`, known `:key:` shortcodes as
 * emoji spans. A `:key:` that cannot be rendered gives its closing colon back,
 * so in "a:b:smile:" the ":b:" miss still leaves ":smile:" to render.
 */
function renderPlain(text: string): string {
  let out = '';
  let at = 0;
  const shortcode = /:([a-z0-9_-]+):/g;
  for (let m = shortcode.exec(text); m !== null; m = shortcode.exec(text)) {
    const span = emojiSpan(m[1]);
    if (span === undefined) {
      shortcode.lastIndex = m.index + m[0].length - 1;
      continue;
    }
    out += escapeLines(text.slice(at, m.index)) + span;
    at = shortcode.lastIndex;
  }
  return out + escapeLines(text.slice(at));
}

const escapeLines = (s: string): string => escapeHtml(s).replace(/\r?\n/g, '<br>');

/** Someone to mention: `text` is how they are named in the message, `mri` who they are. */
export interface MentionTarget {
  text: string;
  mri: string;
}

/** A placed mention, as `properties.mentions` lists it. */
export interface PlacedMention {
  itemid: string;
  mri: string;
  displayName: string;
}

const isWordChar = (c: string | undefined): boolean => c !== undefined && /[\p{L}\p{N}_]/u.test(c);

/**
 * Where `name` first occurs in `text` outside `taken`: `@name` (exactly, then
 * ignoring case) wins, else the bare name as a whole word, ignoring case.
 * Either way the name must END at a word boundary, so "@Sam" is never found
 * inside "@Samantha".
 */
function locate(text: string, name: string, taken: readonly [number, number][]): [number, number] | undefined {
  const free = (s: number, e: number) => taken.every(([ts, te]) => e <= ts || s >= te);
  const find = (needle: string, ci: boolean, wordStart: boolean): [number, number] | undefined => {
    const hay = ci ? text.toLowerCase() : text;
    const n = ci ? needle.toLowerCase() : needle;
    for (let i = hay.indexOf(n); i >= 0; i = hay.indexOf(n, i + 1)) {
      const end = i + n.length;
      if (isWordChar(text[end]) || (wordStart && isWordChar(text[i - 1]))) continue;
      if (free(i, end)) return [i, end];
    }
    return undefined;
  };
  return find(`@${name}`, false, false) ?? find(`@${name}`, true, false) ?? find(name, true, true);
}

/**
 * The message body as Teams HTML: one paragraph of escaped text, with each
 * mention's first free occurrence (`@name`, else `name`) replaced by a mention
 * span. Mentions are numbered 0, 1, … in order of appearance in the text.
 */
export function renderMessageHtml(
  text: string,
  mentions: readonly MentionTarget[],
): { html: string; mentions: PlacedMention[] } {
  const placed: { start: number; end: number; mri: string }[] = [];
  for (const m of mentions) {
    const range = locate(
      text,
      m.text,
      placed.map((p) => [p.start, p.end]),
    );
    if (range === undefined) {
      throw new McpToolError(`The mention "${m.text}" does not appear in the text.`, {
        hint: 'name_in_text must be written in text (as @name or name) once per mention.',
      });
    }
    placed.push({ start: range[0], end: range[1], mri: m.mri });
  }
  placed.sort((a, b) => a.start - b.start);

  let html = '';
  let at = 0;
  const out: PlacedMention[] = [];
  placed.forEach((p, n) => {
    const shown = text.slice(p.start, p.end).replace(/^@/, '');
    html += renderPlain(text.slice(at, p.start));
    html += `<span itemscope itemtype="http://schema.skype.com/Mention" itemid="${n}">${escapeHtml(shown)}</span>`;
    out.push({ itemid: String(n), mri: p.mri, displayName: shown });
    at = p.end;
  });
  html += renderPlain(text.slice(at));
  return { html: `<p>${html}</p>`, mentions: out };
}

/** `properties.mentions` — a JSON STRING, like every `properties` value. */
export function mentionsProperty(mentions: readonly PlacedMention[]): string {
  return JSON.stringify(
    mentions.map((m) => ({
      '@type': 'http://schema.skype.com/Mention',
      itemid: m.itemid,
      mri: m.mri,
      mentionType: 'person',
      displayName: m.displayName,
    })),
  );
}

/** A message being quoted: its id (= arrival ms), sender, and plain text. */
export interface QuotedMessage {
  id: string;
  senderMri: string;
  senderName: string;
  text: string;
}

/** Longest quoted-message preview Teams' own quote block carries. */
const QUOTE_PREVIEW_CHARS = 200;

/**
 * `s` cut to at most `max` UTF-16 units, ending in `ellipsis` when cut. Never
 * splits a surrogate pair, so an emoji at the cut is dropped whole rather than
 * leaving a lone half (invalid UTF-16 that renders as �).
 */
export function clipText(s: string, max: number, ellipsis = ''): string {
  if (s.length <= max) return s;
  let end = Math.max(0, max - ellipsis.length);
  const last = s.charCodeAt(end - 1);
  if (last >= 0xd800 && last <= 0xdbff) end -= 1;
  return s.slice(0, end) + ellipsis;
}

/** The captured quote-reply block, placed before the reply's own paragraph. */
export function quoteHtml(q: QuotedMessage): string {
  const preview = clipText(q.text, QUOTE_PREVIEW_CHARS, '…');
  const id = escapeHtml(q.id);
  return (
    `<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="${id}">` +
    `<strong itemprop="mri" itemid="${escapeHtml(q.senderMri)}">${escapeHtml(q.senderName)}</strong>` +
    `<span itemprop="time" itemid="${id}"></span>` +
    `<p itemprop="preview">${escapeHtml(preview)}</p></blockquote>`
  );
}

/** `properties.qtdMsgs` — the quoted message, as a JSON string. */
export function qtdMsgsProperty(q: QuotedMessage): string {
  return JSON.stringify([{ messageId: q.id, sender: q.senderMri, time: Number(q.id) }]);
}

/**
 * A fresh `clientmessageid`: a random 19-digit number, as Teams generates.
 * It is how the read-back finds the message this send created.
 */
export function newClientMessageId(): string {
  let id = String(randomInt(1, 10));
  for (let i = 0; i < 18; i++) id += String(randomInt(0, 10));
  return id;
}

export interface SendBodyInput {
  /** The conversation posted to (a channel thread is `<channelId>;messageid=<rootId>`). */
  conversationId: string;
  meMri: string;
  meName: string;
  /** Message HTML, from {@link renderMessageHtml} (and {@link quoteHtml}). */
  content: string;
  clientmessageid: string;
  now: Date;
  /** From {@link mentionsProperty}; `"[]"` when absent. */
  mentions?: string;
  /** From {@link qtdMsgsProperty}; omitted when absent. */
  qtdMsgs?: string;
}

export type SendBody = ReturnType<typeof buildSendBody>;

/**
 * The captured `POST …/messages` body. `properties` values are JSON STRINGS,
 * and `conversationLink` really is the literal `blah/<id>` Teams sends.
 */
export function buildSendBody(i: SendBodyInput) {
  const time = i.now.toISOString();
  return {
    id: '-1',
    type: 'Message',
    conversationid: i.conversationId,
    conversationLink: `blah/${i.conversationId}`,
    from: i.meMri,
    fromUserId: i.meMri,
    composetime: time,
    originalarrivaltime: time,
    content: i.content,
    messagetype: 'RichText/Html',
    contenttype: 'Text',
    imdisplayname: i.meName,
    clientmessageid: i.clientmessageid,
    callId: '',
    state: 0,
    version: '0',
    amsreferences: [] as string[],
    properties: {
      importance: '',
      subject: '',
      title: '',
      cards: '[]',
      links: '[]',
      mentions: i.mentions ?? '[]',
      onbehalfof: null,
      files: '[]',
      policyViolation: null,
      formatVariant: 'TEAMS',
      ...(i.qtdMsgs !== undefined ? { qtdMsgs: i.qtdMsgs } : {}),
    } as Record<string, string | null>,
    crossPostChannels: [] as string[],
  };
}
