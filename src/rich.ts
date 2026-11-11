/**
 * The structure of a Teams message body: sanitised HTML, its mentions and
 * links, and a Markdown rendering. Pure functions, no I/O.
 *
 * Teams bodies are a small, regular subset of HTML, so a tiny tolerant parser
 * is enough. What it does not know it drops, keeping the text inside, so
 * nothing a sender wrote is lost and no markup that this module did not choose
 * gets through.
 */

const NAMED_ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: ' ',
};

export function decodeEntities(s: string): string {
  // One pass, so `&amp;lt;` becomes `&lt;` and not `<`.
  return s.replace(/&(?:#(\d+)|#x([0-9a-f]+)|([a-z]+));/gi, (whole, dec, hex, name) => {
    if (name !== undefined) return NAMED_ENTITIES[name.toLowerCase()] ?? whole;
    const code = dec !== undefined ? Number(dec) : parseInt(hex, 16);
    if (!Number.isFinite(code) || code > 0x10ffff) return whole;
    // Drop control characters (keeping tab/newline): they carry no text and
    // would collide with the placeholders the renderers use.
    if (code < 0x20 && code !== 0x09 && code !== 0x0a) return '';
    return String.fromCodePoint(code);
  });
}

/** `s` safe inside a double-quoted attribute or as element text. */
export function escapeAttr(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** A tag's attribute, as written (entities NOT decoded). */
export function attr(attrs: string, name: string): string | undefined {
  const m = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s"'>]+))`, 'i').exec(attrs);
  return m ? (m[1] ?? m[2] ?? m[3]) : undefined;
}

interface Element {
  tag: string;
  attrs: string;
  children: Node[];
}
/** A string is raw (still entity-encoded) text. */
type Node = string | Element;

const VOID = new Set(['br', 'img', 'hr', 'wbr', 'input', 'meta', 'link', 'col', 'source']);
const DROPPED_WITH_CONTENT = 'script|style|head|title|template|noscript|iframe|object|svg|math';
const BLOCK = new Set(['p', 'div', 'ul', 'ol', 'pre', 'blockquote', 'table', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']);
const TAG = /<(\/?)([a-z][a-z0-9]*)\b((?:[^>"']|"[^"]*"|'[^']*')*)>/gi;

const MENTION_TYPE = 'http://schema.skype.com/Mention';
const EMOJI_TYPE = /schema\.skype\.com\/Emoji/i;

/** A forgiving parse into a tree: unknown nesting is closed, stray closing tags are ignored. */
function parse(html: string): Element {
  const root: Element = { tag: '#root', attrs: '', children: [] };
  const stack: Element[] = [root];
  const top = () => stack[stack.length - 1];
  const src = html
    .replace(/[\u0001-\u0003]/g, '')
    .replace(/<!--[\s\S]*?(?:-->|$)/g, '')
    .replace(new RegExp(`<(${DROPPED_WITH_CONTENT})\\b[\\s\\S]*?(?:<\\/\\1\\s*>|$)`, 'gi'), '');
  let at = 0;
  for (const m of src.matchAll(TAG)) {
    if (m.index > at) top().children.push(src.slice(at, m.index));
    at = m.index + m[0].length;
    const tag = m[2].toLowerCase();
    if (m[1] === '/') {
      const i = stack.map((e) => e.tag).lastIndexOf(tag);
      if (i > 0) stack.length = i;
      continue;
    }
    // Implicit closes HTML itself makes: a new item, cell or row ends the
    // last, and a block ends a paragraph.
    if ((tag === 'li' && top().tag === 'li') || (BLOCK.has(tag) && top().tag === 'p')) stack.pop();
    if ((tag === 'td' || tag === 'th' || tag === 'tr') && (top().tag === 'td' || top().tag === 'th')) stack.pop();
    if (tag === 'tr' && top().tag === 'tr') stack.pop();
    const el: Element = { tag, attrs: m[3], children: [] };
    top().children.push(el);
    if (!VOID.has(tag) && !m[3].trimEnd().endsWith('/')) stack.push(el);
  }
  if (at < src.length) top().children.push(src.slice(at));
  return root;
}

/** A link target this module passes on: http(s) or mailto, decoded; `undefined` otherwise. */
function safeHref(attrs: string): string | undefined {
  const raw = attr(attrs, 'href');
  if (raw === undefined) return undefined;
  const href = decodeEntities(raw).trim();
  return /^(https?:\/\/|mailto:)/i.test(href) ? href : undefined;
}

const isMention = (el: Element): boolean => el.tag === 'span' && attr(el.attrs, 'itemtype') === MENTION_TYPE;

/** An inline image as text: an emoji's character, else `[image]` (its ref is in the attachments). */
function imageText(el: Element): string {
  const alt = attr(el.attrs, 'alt');
  return EMOJI_TYPE.test(attr(el.attrs, 'itemtype') ?? '') && alt ? decodeEntities(alt) : '[image]';
}

/** Tags kept as they are, with no attributes. */
const PLAIN_TAGS = new Set([
  'p', 'div', 'br', 'b', 'strong', 'i', 'em', 'u', 's', 'strike', 'del', 'ul', 'li', 'code', 'pre', 'blockquote',
  'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'thead', 'tbody', 'tr', 'td', 'th',
]);

function serialise(nodes: readonly Node[]): string {
  let out = '';
  for (const n of nodes) {
    if (typeof n === 'string') {
      out += n.replace(/\r\n?/g, '\n').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      continue;
    }
    const inner = () => serialise(n.children);
    if (n.tag === 'img') out += escapeAttr(imageText(n));
    else if (n.tag === 'br') out += '<br>';
    else if (PLAIN_TAGS.has(n.tag)) out += `<${n.tag}>${inner()}</${n.tag}>`;
    else if (n.tag === 'ol') {
      const start = attr(n.attrs, 'start');
      out += `<ol${start !== undefined && /^\d+$/.test(start) ? ` start="${start}"` : ''}>${inner()}</ol>`;
    } else if (n.tag === 'a') {
      const href = safeHref(n.attrs);
      out += href !== undefined ? `<a href="${escapeAttr(href)}">${inner()}</a>` : inner();
    } else if (isMention(n)) {
      const id = attr(n.attrs, 'itemid');
      const itemid = id !== undefined && /^\d+$/.test(id) ? ` itemid="${id}"` : '';
      out += `<span itemtype="${MENTION_TYPE}"${itemid}>${inner()}</span>`;
    } else out += inner();
  }
  return out;
}

/**
 * A message body reduced to the formatting Teams shows: paragraphs, line
 * breaks, bold/italic/underline/strike, lists, links (http, https, mailto),
 * inline code and code blocks, quotes, headings, tables and @mention spans.
 * Every attribute but a link's href, a list's start and a mention's id is
 * dropped; scripts, styles and comments go with their content; any other tag
 * goes but its text stays. Emoji become their character, other images `[image]`.
 */
export function sanitiseHtml(html: string): string {
  return serialise(parse(html).children);
}

/** Any tag in sanitised HTML beyond paragraphs and line breaks. */
const FORMATTING = /<(?!\/?(?:p|div|br)>)[a-z]/i;

/** True when `html` carries any formatting beyond paragraphs and line breaks. */
export function isRichHtml(html: string): boolean {
  return FORMATTING.test(sanitiseHtml(html));
}

/** Someone @mentioned in a message. */
export interface MessageMention {
  mri: string;
  displayName: string;
}

/** A link in a message, with the text it was shown as. */
export interface MessageLink {
  href: string;
  text: string;
}

/**
 * `properties.mentions` (a JSON string, or an already-parsed array) as
 * `{ mri, displayName }`; entries without an MRI are skipped and anything
 * unreadable yields none.
 */
export function mentionsOf(raw: unknown): MessageMention[] {
  let parsed: unknown = raw;
  if (typeof raw === 'string') {
    try {
      parsed = JSON.parse(raw);
    } catch {
      return [];
    }
  }
  if (!Array.isArray(parsed)) return [];
  const out: MessageMention[] = [];
  for (const e of parsed) {
    if (e === null || typeof e !== 'object') continue;
    const { mri, displayName } = e as { mri?: unknown; displayName?: unknown };
    if (typeof mri !== 'string' || mri === '') continue;
    out.push({ mri, displayName: typeof displayName === 'string' ? displayName : '' });
  }
  return out;
}

/** An element's text, entities decoded and whitespace collapsed. */
function textOf(nodes: readonly Node[]): string {
  const raw = (ns: readonly Node[]): string =>
    ns.map((n) => (typeof n === 'string' ? decodeEntities(n) : n.tag === 'img' ? imageText(n) : n.tag === 'br' ? ' ' : raw(n.children))).join('');
  return raw(nodes).replace(/\s+/g, ' ').trim();
}

function walk(nodes: readonly Node[], visit: (el: Element) => void): void {
  for (const n of nodes) {
    if (typeof n === 'string') continue;
    visit(n);
    walk(n.children, visit);
  }
}

/** Every http(s)/mailto link in a body with its shown text, in order, each pair once. */
export function linksOf(html: string): MessageLink[] {
  return links(parse(html));
}

function links(root: Element): MessageLink[] {
  const out: MessageLink[] = [];
  const seen = new Set<string>();
  walk(root.children, (el) => {
    if (el.tag !== 'a') return;
    const href = safeHref(el.attrs);
    if (href === undefined) return;
    const text = textOf(el.children);
    const key = `${href}\n${text}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ href, text });
  });
  return out;
}

/**
 * What a read shows of a body's structure, from one parse: its sanitised HTML
 * when it carries formatting (`undefined` for a plain body, whose text says
 * everything), and its links.
 */
export function bodyStructure(html: string): { contentHtml?: string; links: MessageLink[] } {
  const root = parse(html);
  const sanitised = serialise(root.children);
  return { ...(FORMATTING.test(sanitised) ? { contentHtml: sanitised } : {}), links: links(root) };
}

/* ----------------------------------------------------------------------------
 * Plain text
 * ------------------------------------------------------------------------- */

/*
 * Line markers that survive until the lines are laid out, so a quote's depth
 * applies to every line inside it. `parse` strips them from the input, so a
 * message cannot forge one.
 */
const QUOTE_OPEN = '\u0001';
const QUOTE_CLOSE = '\u0002';

/** Elements that start and end a line of text. */
const TEXT_BLOCK = new Set(['p', 'div', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']);

/** A link as text: `label (url)`, or the address alone when that is all it shows. */
function linkText(el: Element, label: string): string {
  const href = attr(el.attrs, 'href');
  if (!href) return label;
  const url = decodeEntities(href);
  const shown = label.trim();
  // Teams shows a pasted link as its own, often clipped, address: give it once.
  if (shown === '' || shown === url || /^(https?:\/\/|www\.)/i.test(shown)) return url;
  return `${label} (${url})`;
}

/** One table row on one line, its cells joined by ` | ` (line breaks in a cell become spaces). */
function rowText(tr: Element): string {
  const cells = tr.children.filter((n): n is Element => typeof n !== 'string' && (n.tag === 'td' || n.tag === 'th'));
  if (cells.length === 0) return toText(tr.children);
  return cells.map((c) => toText(c.children).replace(/[\n\u0001\u0002]+/g, ' ').trim()).join(' | ');
}

function toText(nodes: readonly Node[]): string {
  let out = '';
  for (const n of nodes) {
    if (typeof n === 'string') {
      out += decodeEntities(n);
      continue;
    }
    if (n.tag === 'br') out += '\n';
    else if (n.tag === 'img') out += imageText(n);
    else if (n.tag === 'a') out += linkText(n, toText(n.children));
    else if (n.tag === 'tr') out += `\n${rowText(n)}\n`;
    else if (n.tag === 'blockquote') out += `\n${QUOTE_OPEN}\n${toText(n.children)}\n${QUOTE_CLOSE}\n`;
    else if (n.tag === 'ol' || n.tag === 'ul') {
      const start = Number(attr(n.attrs, 'start'));
      let i = n.tag === 'ol' && Number.isInteger(start) && start >= 0 ? start : 1;
      for (const c of n.children) {
        if (typeof c !== 'string' && c.tag === 'li') {
          out += `\n${n.tag === 'ol' ? `${i++}. ` : '- '}${toText(c.children)}\n`;
        } else out += toText([c]);
      }
    } else if (TEXT_BLOCK.has(n.tag)) out += `\n${toText(n.children)}\n`;
    else out += toText(n.children);
  }
  return out;
}

/**
 * A Teams message body as readable plain text. Links keep their target as
 * `text (url)`, inline images become `[image]` (emoji become their character),
 * a numbered list's items `1. `, `2. ` … and a bulleted one's `- `, a table
 * row one line of ` | `-separated cells, and a quoted reply's lines are
 * prefixed `> `. Never truncated: a long body comes back whole.
 */
export function htmlToText(html: string): string {
  const out: string[] = [];
  let depth = 0;
  for (const raw of toText(parse(html).children).split(/\r?\n/)) {
    const line = raw.replace(/[ \t\u00a0]+/g, ' ').trim();
    if (line === QUOTE_OPEN) depth++;
    else if (line === QUOTE_CLOSE) depth = Math.max(0, depth - 1);
    else if (line !== '') out.push(`${'> '.repeat(depth)}${line}`);
  }
  return out.join('\n');
}

/* ----------------------------------------------------------------------------
 * Markdown
 * ------------------------------------------------------------------------- */

/** Code blocks wait here while whitespace is tidied, so their indentation survives. */
const CODE_HOLE = '\u0003';
const block = (s: string): string => `\n\n${s}\n\n`;

function wrap(marker: string, inner: string): string {
  const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(inner)!;
  return m[2] === '' ? inner : `${m[1]}${marker}${m[2]}${marker}${m[3]}`;
}

/** Text with `<br>` as newlines and whitespace kept, for code. */
function codeText(nodes: readonly Node[]): string {
  return nodes
    .map((n) => (typeof n === 'string' ? decodeEntities(n) : n.tag === 'br' ? '\n' : codeText(n.children)))
    .join('');
}

function inlineCode(text: string): string {
  const longest = Math.max(0, ...[...text.matchAll(/`+/g)].map((m) => m[0].length));
  const fence = '`'.repeat(longest + 1);
  const pad = text.startsWith('`') || text.endsWith('`') ? ' ' : '';
  return `${fence}${pad}${text}${pad}${fence}`;
}

const tidyInner = (s: string): string => s.replace(/\n{3,}/g, '\n\n').trim();

function list(el: Element, codes: string[]): string {
  const ordered = el.tag === 'ol';
  const startAttr = Number(attr(el.attrs, 'start'));
  const start = ordered && Number.isInteger(startAttr) && startAttr >= 0 ? startAttr : 1;
  const items = el.children.filter((n): n is Element => typeof n !== 'string' && n.tag === 'li');
  return items
    .map((li, i) => {
      const marker = ordered ? `${start + i}. ` : '- ';
      const lines = md(li.children, codes).replace(/\n{2,}/g, '\n').trim().split('\n');
      return lines.map((l, n) => (n === 0 ? marker + l : l === '' ? '' : ' '.repeat(marker.length) + l)).join('\n');
    })
    .join('\n');
}

function table(el: Element, codes: string[]): string {
  const rows: Element[] = [];
  walk(el.children, (e) => {
    if (e.tag === 'tr') rows.push(e);
  });
  const lines = rows.map((tr) => {
    const cells = tr.children
      .filter((n): n is Element => typeof n !== 'string' && (n.tag === 'td' || n.tag === 'th'))
      .map((c) => md(c.children, codes).replace(/\s*\n\s*/g, ' ').trim().replace(/\|/g, '\\|'));
    return { text: `| ${cells.join(' | ')} |`, width: cells.length };
  });
  if (lines.length === 0) return '';
  const sep = `| ${Array.from({ length: lines[0].width }, () => '---').join(' | ')} |`;
  return [lines[0].text, sep, ...lines.slice(1).map((l) => l.text)].join('\n');
}

function md(nodes: readonly Node[], codes: string[]): string {
  let out = '';
  for (const n of nodes) {
    if (typeof n === 'string') {
      out += decodeEntities(n).replace(/[\s ]+/g, ' ');
      continue;
    }
    const inner = () => md(n.children, codes);
    switch (n.tag) {
      case 'br':
        out += '\n';
        break;
      case 'p':
      case 'div':
        out += block(inner().trim());
        break;
      case 'h1':
      case 'h2':
      case 'h3':
      case 'h4':
      case 'h5':
      case 'h6':
        out += block(`${'#'.repeat(Number(n.tag[1]))} ${inner().trim()}`);
        break;
      case 'b':
      case 'strong':
        out += wrap('**', inner());
        break;
      case 'i':
      case 'em':
        out += wrap('_', inner());
        break;
      case 's':
      case 'strike':
      case 'del':
        out += wrap('~~', inner());
        break;
      case 'code':
        out += inlineCode(codeText(n.children));
        break;
      case 'pre': {
        codes.push('```\n' + codeText(n.children).replace(/\n+$/, '') + '\n```');
        out += block(`${CODE_HOLE}${codes.length - 1}${CODE_HOLE}`);
        break;
      }
      case 'a': {
        const label = inner().trim();
        const href = safeHref(n.attrs);
        if (href === undefined) out += label;
        // Teams shows a pasted link as its own (often clipped) address: the address alone says it.
        else if (label === '' || label === href || /^(https?:\/\/|www\.)/i.test(label)) out += href;
        else out += `[${label}](${href.replace(/[()\s]/g, (c) => encodeURIComponent(c))})`;
        break;
      }
      case 'blockquote':
        out += block(
          tidyInner(inner())
            .split('\n')
            .map((l) => (l === '' ? '>' : `> ${l}`))
            .join('\n'),
        );
        break;
      case 'ul':
      case 'ol':
        out += block(list(n, codes));
        break;
      case 'table':
        out += block(table(n, codes));
        break;
      case 'img':
        out += imageText(n);
        break;
      case 'span':
        out += isMention(n) ? `@${inner().trim()}` : inner();
        break;
      default:
        out += inner();
    }
  }
  return out;
}

/**
 * A message body as Markdown: bold, italic, strike, inline code, fenced code
 * blocks, numbered and bulleted lists (nested ones indented), links with
 * their text, quotes, headings, tables, and mentions as `@name`.
 */
export function htmlToMarkdown(html: string): string {
  const codes: string[] = [];
  const lines = md(parse(html).children, codes).split('\n');
  const tidy: string[] = [];
  for (const raw of lines) {
    const line = raw.trimEnd();
    // A line opening a block loses the space its source indentation left; list nesting is never there.
    tidy.push(tidy.length === 0 || tidy[tidy.length - 1] === '' ? line.trimStart() : line);
  }
  return tidy
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
    .replace(new RegExp(`${CODE_HOLE}(\\d+)${CODE_HOLE}`, 'g'), (_w, i: string) => codes[Number(i)]);
}
