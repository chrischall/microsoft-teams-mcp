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

/** How a message's text is written: plain text, or the Markdown subset {@link renderMarkdown} reads. */
export type MessageFormat = 'text' | 'markdown';
export const MESSAGE_FORMATS = ['text', 'markdown'] as const;

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
 * Private-use placeholders the Markdown renderer carries through untouched:
 * a mention (MENTION_OPEN n MENTION_CLOSE) and a piece already rendered to
 * HTML (HOLE_OPEN n HOLE_CLOSE). The caller's text has them stripped first,
 * so it cannot forge one.
 */
const MENTION_OPEN = '';
const MENTION_CLOSE = '';
const HOLE_OPEN = '';
const HOLE_CLOSE = '';
const PLACEHOLDER_CHARS = /[-]/g;

/**
 * The message body as Teams HTML, with each mention's first free occurrence
 * (`@name`, else `name`) replaced by a mention span. Mentions are numbered
 * 0, 1, … in order of appearance in the text.
 *
 * `text` format: one paragraph of escaped text, newlines as `<br>`.
 * `markdown` format: the subset {@link renderMarkdown} reads.
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
    const html = renderMarkdown(src).replace(
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

/* ----------------------------------------------------------------------------
 * Markdown → Teams HTML
 * ------------------------------------------------------------------------- */

/** Link targets a message may carry. */
const SAFE_HREF = /^(https?:\/\/|mailto:)/i;

/**
 * Inline Markdown → HTML: `code`, [text](url) (http, https, mailto only),
 * **bold** / __bold__, *italic* / _italic_, ~~strike~~, backslash escapes,
 * :key: emoji, newlines as `<br>`. Everything else is escaped text.
 */
function renderInline(src: string): string {
  const holes: string[] = [];
  const hole = (html: string) => `${HOLE_OPEN}${holes.push(html) - 1}${HOLE_CLOSE}`;
  const token =
    /(`+)([\s\S]*?[^`])\1(?!`)|\[([^\]\n]+)\]\(\s*([^\s()]+(?:\([^\s()]*\)[^\s()]*)*)\s*\)|\\([\\`*_{}[\]()#+\-.!~>|:])/g;
  let s = src.replace(token, (whole, _ticks, code: string, label: string, href: string, escaped: string) => {
    if (code !== undefined) return hole(`<code>${escapeHtml(code.replace(/^ (.*[^ ].*) $/s, '$1'))}</code>`);
    if (escaped !== undefined) return hole(escapeHtml(escaped));
    if (!SAFE_HREF.test(href)) return whole;
    return hole(`<a href="${escapeHtml(href)}">${renderInline(label)}</a>`);
  });
  // Emoji before escaping and emphasis, so a key's "_" is never read as italic.
  s = s.replace(/:([a-z0-9_-]+):/g, (whole, key: string) => {
    const span = emojiSpan(key);
    return span === undefined ? whole : hole(span);
  });
  s = escapeHtml(s)
    .replace(/\*\*(?=\S)([\s\S]*?\S)\*\*/g, '<strong>$1</strong>')
    .replace(/(^|[^\w])__(?=\S)([\s\S]*?\S)__(?!\w)/g, '$1<strong>$2</strong>')
    .replace(/~~(?=\S)([\s\S]*?\S)~~/g, '<s>$1</s>')
    .replace(/(^|[^\w*])\*(?=[^\s*])([\s\S]*?[^\s*])\*(?![\w*])/g, '$1<em>$2</em>')
    .replace(/(^|[^\w])_(?=[^\s_])([\s\S]*?[^\s_])_(?!\w)/g, '$1<em>$2</em>')
    .replace(/\n/g, '<br>');
  const fill = new RegExp(`${HOLE_OPEN}(\\d+)${HOLE_CLOSE}`, 'g');
  return s.replace(fill, (_w, n: string) => holes[Number(n)]);
}

interface ListNode {
  indent: number;
  ordered: boolean;
  start: number;
  items: { lines: string[]; children: ListNode[] }[];
}

const LIST_ITEM = /^( *)([-*+]|\d{1,9}[.)])[ \t]+(.*)$/;

function renderList(list: ListNode): string {
  const tag = list.ordered ? 'ol' : 'ul';
  const start = list.ordered && list.start !== 1 ? ` start="${list.start}"` : '';
  const items = list.items.map((i) => `<li>${renderInline(i.lines.join('\n'))}${i.children.map(renderList).join('')}</li>`);
  return `<${tag}${start}>${items.join('')}</${tag}>`;
}

/**
 * Consecutive list lines (items, and indented lines continuing one) as one
 * list, nested by indentation. Returns the HTML and the lines used.
 */
function parseList(lines: readonly string[], from: number): { html: string; next: number } {
  const roots: ListNode[] = [];
  const stack: ListNode[] = [];
  let i = from;
  for (; i < lines.length; i++) {
    const line = lines[i];
    const m = LIST_ITEM.exec(line);
    if (m === null) {
      // An indented line continues the last item; anything else ends the list.
      const last = stack.at(-1)?.items.at(-1);
      if (line.trim() !== '' && /^\s/.test(line) && last !== undefined) {
        last.lines.push(line.trim());
        continue;
      }
      break;
    }
    const indent = m[1].length;
    const ordered = /\d/.test(m[2]);
    while (stack.length > 0 && indent < stack[stack.length - 1].indent) stack.pop();
    let top = stack.at(-1);
    if (top !== undefined && indent >= top.indent + 2 && top.items.length > 0) {
      const nested: ListNode = { indent, ordered, start: ordered ? Number.parseInt(m[2], 10) : 1, items: [] };
      top.items[top.items.length - 1].children.push(nested);
      stack.push(nested);
      top = nested;
    } else if (top === undefined || top.ordered !== ordered) {
      const fresh: ListNode = { indent, ordered, start: ordered ? Number.parseInt(m[2], 10) : 1, items: [] };
      if (top !== undefined && stack.length > 1) {
        stack.pop();
        stack[stack.length - 1].items.at(-1)!.children.push(fresh);
      } else {
        roots.push(fresh);
        stack.length = 0;
      }
      stack.push(fresh);
      top = fresh;
    }
    top.items.push({ lines: [m[3]], children: [] });
  }
  return { html: roots.map(renderList).join(''), next: i };
}

/**
 * The Markdown subset a message may use, as Teams HTML: paragraphs (a blank
 * line between them; a single newline is a line break), **bold**, *italic*,
 * ~~strike~~, `inline code`, fenced code blocks, bulleted (`-`, `*`, `+`) and
 * numbered (`1.`) lists nested by indentation, [links](https://…) with their
 * text, and :key: emoji. Anything else — raw HTML, headings, quotes, images —
 * is escaped and shown as written.
 */
export function renderMarkdown(src: string): string {
  const lines = src.replace(/\r\n?/g, '\n').split('\n');
  const out: string[] = [];
  let para: string[] = [];
  const flush = () => {
    if (para.length > 0) out.push(`<p>${renderInline(para.join('\n'))}</p>`);
    para = [];
  };
  for (let i = 0; i < lines.length; ) {
    const line = lines[i];
    const fence = /^ {0,3}(`{3,}|~{3,})/.exec(line);
    if (fence !== null) {
      flush();
      const close = new RegExp(`^ {0,3}${fence[1][0] === '`' ? '`' : '~'}{${fence[1].length},}\\s*$`);
      const code: string[] = [];
      for (i++; i < lines.length && !close.test(lines[i]); i++) code.push(lines[i]);
      i++; // the closing fence (or past the end)
      out.push(`<pre><code>${escapeHtml(code.join('\n'))}</code></pre>`);
      continue;
    }
    if (LIST_ITEM.test(line)) {
      flush();
      const { html, next } = parseList(lines, i);
      out.push(html);
      i = next;
      continue;
    }
    if (line.trim() === '') flush();
    else para.push(line);
    i++;
  }
  flush();
  return out.join('');
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
