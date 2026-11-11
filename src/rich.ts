/**
 * The structure of a Teams message body: sanitised HTML, its mentions and
 * links, plain text and a Markdown rendering. Pure functions, no I/O.
 *
 * The parsing and rendering live in `@chrischall/mcp-utils/rich`; this module
 * is the Teams layer over it: {@link TEAMS} keeps @mention spans (shown as
 * `@name` in Markdown) and shows an emoji image as its character and any
 * other inline image as `[image]` (its ref is in the attachments). Mentions'
 * MRIs come from `properties.mentions` ({@link mentionsOf}).
 */
import * as rich from '@chrischall/mcp-utils/rich';
import type { HtmlReadOptions, RichLink } from '@chrischall/mcp-utils/rich';

export { attr, decodeEntities } from '@chrischall/mcp-utils/rich';

const MENTION_TYPE = 'http://schema.skype.com/Mention';
const EMOJI_TYPE = /schema\.skype\.com\/Emoji/i;

/** How Teams' own elements read: mention spans kept, emoji as their character, other images `[image]`. */
const TEAMS: HtmlReadOptions = {
  element: (el, output, inner) => {
    if (el.tag === 'img') {
      const alt = rich.attr(el.attrs, 'alt');
      const text = EMOJI_TYPE.test(rich.attr(el.attrs, 'itemtype') ?? '') && alt ? rich.decodeEntities(alt) : '[image]';
      return output === 'html' ? rich.escapeHtml(text) : text;
    }
    if (el.tag === 'span' && rich.attr(el.attrs, 'itemtype') === MENTION_TYPE) {
      if (output === 'markdown') return `@${inner.trim()}`;
      if (output === 'html') {
        const id = rich.attr(el.attrs, 'itemid');
        const itemid = id !== undefined && /^\d+$/.test(id) ? ` itemid="${id}"` : '';
        return `<span itemtype="${MENTION_TYPE}"${itemid}>${inner}</span>`;
      }
    }
    return undefined;
  },
};

/**
 * A message body reduced to the formatting Teams shows: paragraphs, line
 * breaks, bold/italic/underline/strike, lists, links (http, https, mailto),
 * inline code and code blocks, quotes, headings, tables and @mention spans.
 * Every attribute but a link's href, a list's start and a mention's id is
 * dropped; scripts, styles, comments and hidden content go with their
 * content; any other tag goes but its text stays. Emoji become their
 * character, other images `[image]`.
 */
export function sanitiseHtml(html: string): string {
  return rich.sanitiseHtml(html, TEAMS);
}

/** True when `html` carries any formatting beyond paragraphs and line breaks. */
export function isRichHtml(html: string): boolean {
  return rich.isRichHtml(html, TEAMS);
}

/** A link in a message, with the text it was shown as. */
export type MessageLink = RichLink;

/** Every http(s)/mailto link in a body with its shown text, in order, each pair once. */
export function linksOf(html: string): MessageLink[] {
  return rich.linksOf(html, TEAMS);
}

/**
 * What a read shows of a body's structure, from one parse: its sanitised HTML
 * when it carries formatting (`undefined` for a plain body, whose text says
 * everything), and its links.
 */
export function bodyStructure(html: string): { contentHtml?: string; links: MessageLink[] } {
  return rich.bodyStructure(html, TEAMS);
}

/**
 * A Teams message body as readable plain text. Links keep their target as
 * `text (url)`, inline images become `[image]` (emoji become their character),
 * a numbered list's items `1. `, `2. ` … and a bulleted one's `- `, a table
 * row one line of ` | `-separated cells, and a quoted reply's lines are
 * prefixed `> `. Never truncated: a long body comes back whole.
 */
export function htmlToText(html: string): string {
  return rich.htmlToText(html, TEAMS);
}

/**
 * A message body as Markdown: bold, italic, strike, inline code, fenced code
 * blocks, numbered and bulleted lists (nested ones indented), links with
 * their text, quotes, headings, tables, and mentions as `@name`.
 */
export function htmlToMarkdown(html: string): string {
  return rich.htmlToMarkdown(html, TEAMS);
}

/** Someone @mentioned in a message. */
export interface MessageMention {
  mri: string;
  displayName: string;
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
