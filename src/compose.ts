/**
 * Building an outgoing Teams message — pure functions, no I/O.
 *
 * Every shape here is the one captured from Teams' own requests (see "Verified
 * write shapes" in `docs/plans/2026-10-10-chatsvc-api.md`), except the mention
 * markup, which follows the documented Skype format and is NOT yet captured
 * live. The caller's text is always escaped: the only markup a message can
 * carry is what these builders emit (paragraph, line breaks, mention spans,
 * emoji spans, the quote block) and, for `markdown` text, what the Markdown
 * renderer in `@chrischall/mcp-utils/rich` emits, so a caller cannot smuggle
 * HTML into a post.
 */
import { randomInt } from 'node:crypto';
import { McpToolError } from '@chrischall/mcp-utils';
import { escapeHtml, renderMarkdown } from '@chrischall/mcp-utils/rich';

/** `s` safe as HTML text or inside a quoted attribute (`'` → `&#39;`). */
export { escapeHtml };

/**
 * Teams' short legacy emoji keys, rendered as Teams' own animated emoji span,
 * with the title Teams shows and the character used as the `alt` fallback.
 * The keys are the observed ones, and each has an image on Teams' emoji CDN
 * (all 12 checked 2026-11-11); the titles and characters only affect hover
 * text and copy-paste.
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
 * Common shortcode names (the GitHub/Slack spellings) sent as the plain
 * character, which Teams always shows — unlike an image span, whose image
 * Teams may not have.
 */
const PLAIN_EMOJI: Record<string, string> = {
  rocket: '🚀',
  tada: '🎉',
  thumbsup: '👍',
  thumbsdown: '👎',
  clap: '👏',
  raised_hands: '🙌',
  pray: '🙏',
  muscle: '💪',
  wave: '👋',
  eyes: '👀',
  thinking: '🤔',
  joy: '😂',
  sweat_smile: '😅',
  grin: '😁',
  slightly_smiling_face: '🙂',
  upside_down: '🙃',
  shrug: '🤷',
  facepalm: '🤦',
  white_check_mark: '✅',
  heavy_check_mark: '✔️',
  x: '❌',
  warning: '⚠️',
  bangbang: '‼️',
  question: '❓',
  bulb: '💡',
  star: '⭐',
  sparkles: '✨',
  zap: '⚡',
  boom: '💥',
  '100': '💯',
  robot: '🤖',
  brain: '🧠',
  coffee: '☕',
  chart_with_upwards_trend: '📈',
  chart_with_downwards_trend: '📉',
  memo: '📝',
  calendar: '📅',
  lock: '🔒',
  bug: '🐛',
  wrench: '🔧',
  hammer: '🔨',
  package: '📦',
  ship: '🚢',
  hourglass: '⌛',
  stopwatch: '⏱️',
  point_right: '👉',
  point_up: '☝️',
  ok_hand: '👌',
  crossed_fingers: '🤞',
  see_no_evil: '🙈',
  skull: '💀',
  partying_face: '🥳',
  exploding_head: '🤯',
  mind_blown: '🤯',
};

/**
 * A unicode-style key: lowercase hex code point(s), `_`, the character's name
 * squashed to `[a-z0-9]` (`1f389_partypopper`); a multi-code-point key is
 * `-`-joined.
 */
const UNICODE_KEY = /^([0-9a-f]{4,6}(?:-[0-9a-f]{4,6})*)_([a-z0-9]+)$/;

/**
 * The HTML for the `:key:` shortcode, or `undefined` when the key is none
 * this module knows — the caller then leaves the `:key:` as plain text (so
 * "10:30:45" is never mangled).
 *
 * A built-in name ({@link NAMED_EMOJI}) is Teams' animated emoji span. A
 * common name ({@link PLAIN_EMOJI}) or a unicode-style key is the plain
 * character: Teams' emoji CDN has no image for many code points (live
 * 2026-11-11, `1f680_rocket` is a 404 and showed as a broken image), and a
 * plain character always shows.
 */
export function emojiHtml(key: string): string | undefined {
  const named = NAMED_EMOJI[key];
  if (named !== undefined) {
    const t = escapeHtml(named.title);
    // `key` is one of the names above, so it is safe in a URL and attribute.
    return (
      `<span contenteditable="false" title="${t}" type="(${key})" class="animated-emoticon-20-${key}" itemscope>` +
      `<img itemscope itemtype="http://schema.skype.com/Emoji" itemid="${key}" ` +
      `src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/${key}/default/20_f.png" ` +
      `title="${t}" alt="${escapeHtml(named.char)}" style="width:20px;height:20px;"></span>`
    );
  }
  const plain = Object.hasOwn(PLAIN_EMOJI, key) ? PLAIN_EMOJI[key] : undefined;
  if (plain !== undefined) return plain;
  const m = UNICODE_KEY.exec(key);
  if (m === null) return undefined;
  const points = m[1].split('-').map((h) => parseInt(h, 16));
  if (points.some((p) => p > 0x10ffff)) return undefined;
  return escapeHtml(String.fromCodePoint(...points));
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
    const span = emojiHtml(m[1]);
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
 * How a message's text is written: plain text, or Markdown — the subset
 * `renderMarkdown` in `@chrischall/mcp-utils/rich` reads: paragraphs (a blank
 * line between them; a single newline is a line break), **bold**, *italic*,
 * ~~strike~~, `inline code`, fenced code blocks, nested bulleted and numbered
 * lists, [links](https://…) with their text, and :key: emoji. Anything else —
 * raw HTML, headings, quotes, images — is escaped and shown as written.
 */
export const MESSAGE_FORMATS = ['text', 'markdown'] as const;
export type MessageFormat = (typeof MESSAGE_FORMATS)[number];

/** Each mention's first free occurrence in `text`, in order of appearance. */
function placeMentions(text: string, mentions: readonly MentionTarget[]): { start: number; end: number; mri: string }[] {
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
  return placed.sort((a, b) => a.start - b.start);
}

const mentionSpan = (n: number, shown: string): string =>
  `<span itemscope itemtype="http://schema.skype.com/Mention" itemid="${n}">${escapeHtml(shown)}</span>`;

/*
 * Private-use placeholders for a mention (MENTION_OPEN n MENTION_CLOSE),
 * carried through the Markdown renderer as ordinary text. They sit above the
 * renderer's own reserved range (U+E000–U+E00F, which it strips), and the
 * caller's text has U+E010–U+E01F stripped first, so it cannot forge one.
 */
const MENTION_OPEN = '\ue010';
const MENTION_CLOSE = '\ue011';
const PLACEHOLDER_CHARS = /[\ue010-\ue01f]/g;

/**
 * The message body as Teams HTML, with each mention's first free occurrence
 * (`@name`, else `name`) replaced by a mention span. Mentions are numbered
 * 0, 1, … in order of appearance in the text.
 *
 * `text` format: one paragraph of escaped text, newlines as `<br>`.
 * `markdown` format: the subset described at {@link MESSAGE_FORMATS}.
 */
export function renderMessageHtml(
  text: string,
  mentions: readonly MentionTarget[],
  format: MessageFormat = 'text',
): { html: string; mentions: PlacedMention[] } {
  const placed = placeMentions(text, mentions);
  const out: PlacedMention[] = placed.map((p, n) => ({
    itemid: String(n),
    mri: p.mri,
    displayName: text.slice(p.start, p.end).replace(/^@/, ''),
  }));

  if (format === 'markdown') {
    let src = '';
    let at = 0;
    const clean = (s: string) => s.replace(PLACEHOLDER_CHARS, '');
    placed.forEach((p, n) => {
      src += clean(text.slice(at, p.start)) + `${MENTION_OPEN}${n}${MENTION_CLOSE}`;
      at = p.end;
    });
    src += clean(text.slice(at));
    const html = renderMarkdown(src, { shortcode: emojiHtml }).replace(
      new RegExp(`${MENTION_OPEN}(\\d+)${MENTION_CLOSE}`, 'g'),
      (_w, n: string) => mentionSpan(Number(n), out[Number(n)].displayName),
    );
    return { html, mentions: out };
  }

  let html = '';
  let at = 0;
  placed.forEach((p, n) => {
    html += renderPlain(text.slice(at, p.start));
    html += mentionSpan(n, out[n].displayName);
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
