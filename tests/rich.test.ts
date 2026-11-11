import { describe, expect, it } from 'vitest';
import { htmlToMarkdown, isRichHtml, linksOf, mentionsOf, sanitiseHtml } from '../src/rich.js';

const MENTION = (n: number, name: string) =>
  `<span itemscope="" itemtype="http://schema.skype.com/Mention" itemid="${n}">${name}</span>`;

describe('sanitiseHtml', () => {
  it('keeps formatting tags and drops their attributes', () => {
    expect(sanitiseHtml('<p style="color:red" class="x"><b data-x="1">bold</b> <em>it</em></p>')).toBe(
      '<p><b>bold</b> <em>it</em></p>',
    );
  });

  it('keeps lists, code, quotes, headings and tables', () => {
    const html =
      '<ol start="3"><li>a</li><li>b</li></ol><ul><li>c</li></ul><pre class="language-js"><code>x &lt; 1</code></pre>' +
      '<blockquote><p>q</p></blockquote><h2>H</h2><table><tr><td>1</td></tr></table>';
    expect(sanitiseHtml(html)).toBe(
      '<ol start="3"><li>a</li><li>b</li></ol><ul><li>c</li></ul><pre><code>x &lt; 1</code></pre>' +
        '<blockquote><p>q</p></blockquote><h2>H</h2><table><tr><td>1</td></tr></table>',
    );
  });

  it('removes script and style elements with their content, and unknown tags but not their text', () => {
    expect(sanitiseHtml('<p>a<script>alert(1)</script><style>p{}</style><font face="x">b</font></p>')).toBe('<p>ab</p>');
  });

  it('keeps only http(s) and mailto link targets', () => {
    expect(sanitiseHtml('<a href="https://example.com/?a=1&amp;b=2" onclick="x()">ok</a>')).toBe(
      '<a href="https://example.com/?a=1&amp;b=2">ok</a>',
    );
    expect(sanitiseHtml('<a href="javascript:alert(1)">bad</a>')).toBe('bad');
    expect(sanitiseHtml('<a href="mailto:a@example.com">mail</a>')).toBe('<a href="mailto:a@example.com">mail</a>');
  });

  it('keeps a mention span with its itemid, and drops other spans', () => {
    expect(sanitiseHtml(`<p>hi ${MENTION(0, 'Alex Example')} <span style="x">there</span></p>`)).toBe(
      '<p>hi <span itemtype="http://schema.skype.com/Mention" itemid="0">Alex Example</span> there</p>',
    );
  });

  it('turns emoji into their character and other images into [image]', () => {
    const emoji = '<img itemtype="http://schema.skype.com/Emoji" alt="🙂" src="https://statics.example/e.png">';
    expect(sanitiseHtml(`<p>${emoji}<img src="https://us-api.asm.skype.com/v1/objects/x/views/imgo"></p>`)).toBe(
      '<p>🙂[image]</p>',
    );
  });

  it('writes CRLF line endings as LF', () => {
    expect(sanitiseHtml('<p>a</p>\r\n<p>b</p>')).toBe('<p>a</p>\n<p>b</p>');
  });

  it('drops comments and escapes stray angle brackets', () => {
    expect(sanitiseHtml('<p>a<!-- hidden -->b 1 < 2</p>')).toBe('<p>ab 1 &lt; 2</p>');
  });

  it('closes tags left open and ignores stray closing tags', () => {
    expect(sanitiseHtml('<p><b>x</p></i>y')).toBe('<p><b>x</b></p>y');
  });
});

describe('isRichHtml', () => {
  it('is false for plain paragraphs and line breaks', () => {
    expect(isRichHtml('<p>one<br>two</p><div>three</div>')).toBe(false);
  });
  it('is true for formatting, lists, links and mentions', () => {
    expect(isRichHtml('<p><b>x</b></p>')).toBe(true);
    expect(isRichHtml('<ol><li>a</li></ol>')).toBe(true);
    expect(isRichHtml('<a href="https://x.example">x</a>')).toBe(true);
    expect(isRichHtml(`<p>${MENTION(0, 'A')}</p>`)).toBe(true);
  });
});

describe('mentionsOf', () => {
  it('reads properties.mentions as a JSON string', () => {
    const raw = JSON.stringify([
      { '@type': 'http://schema.skype.com/Mention', itemid: '0', mri: '8:orgid:abc', mentionType: 'person', displayName: 'Alex Example' },
    ]);
    expect(mentionsOf(raw)).toEqual([{ mri: '8:orgid:abc', displayName: 'Alex Example' }]);
  });
  it('reads an already-parsed array and skips entries without an mri', () => {
    expect(mentionsOf([{ mri: '8:orgid:a', displayName: 'A' }, { displayName: 'no mri' }, null])).toEqual([
      { mri: '8:orgid:a', displayName: 'A' },
    ]);
  });
  it('is empty for anything unreadable', () => {
    expect(mentionsOf('not json')).toEqual([]);
    expect(mentionsOf(undefined)).toEqual([]);
    expect(mentionsOf('[]')).toEqual([]);
  });
});

describe('linksOf', () => {
  it('lists each link with its text, in order, once', () => {
    const html =
      '<p><a href="https://a.example/x?y=1&amp;z=2">the <b>doc</b></a> and <a href="https://b.example">https://b.example</a>' +
      ' and <a href="https://a.example/x?y=1&amp;z=2">the doc</a> <a href="javascript:x">no</a></p>';
    expect(linksOf(html)).toEqual([
      { href: 'https://a.example/x?y=1&z=2', text: 'the doc' },
      { href: 'https://b.example', text: 'https://b.example' },
    ]);
  });
});

describe('htmlToMarkdown', () => {
  it('renders bold, italic, strike and inline code', () => {
    expect(htmlToMarkdown('<p><b>bold</b> <i>it</i> <s>gone</s> <code>x()</code></p>')).toBe('**bold** _it_ ~~gone~~ `x()`');
  });

  it('numbers an ordered list from its start, and bullets an unordered one', () => {
    expect(htmlToMarkdown('<p>Steps:</p><ol><li>one</li><li>two</li></ol><ul><li>a</li><li>b</li></ul>')).toBe(
      'Steps:\n\n1. one\n2. two\n\n- a\n- b',
    );
    expect(htmlToMarkdown('<ol start="4"><li>four</li><li>five</li></ol>')).toBe('4. four\n5. five');
  });

  it('indents a nested list under its item', () => {
    expect(htmlToMarkdown('<ol><li>top<ul><li>inner</li></ul></li><li>next</li></ol>')).toBe('1. top\n   - inner\n2. next');
  });

  it('renders links with their text, and a bare link as its address', () => {
    expect(htmlToMarkdown('<p><a href="https://a.example/doc">the doc</a></p>')).toBe('[the doc](https://a.example/doc)');
    expect(htmlToMarkdown('<p><a href="https://a.example/very/long/path">https://a.example/very/lo…</a></p>')).toBe(
      'https://a.example/very/long/path',
    );
  });

  it('renders a code block fenced, its text kept verbatim', () => {
    expect(htmlToMarkdown('<pre><code>if (a &lt; b) {\n  go();\n}</code></pre>')).toBe('```\nif (a < b) {\n  go();\n}\n```');
  });

  it('renders mentions as @name, headings, quotes, and tables', () => {
    expect(htmlToMarkdown(`<p>hi ${MENTION(0, 'Alex Example')}</p>`)).toBe('hi @Alex Example');
    expect(htmlToMarkdown('<h2>Title</h2><p>x</p>')).toBe('## Title\n\nx');
    expect(htmlToMarkdown('<blockquote><p><b>Sam</b></p><p>said this</p></blockquote><p>reply</p>')).toBe(
      '> **Sam**\n>\n> said this\n\nreply',
    );
    expect(htmlToMarkdown('<table><tr><th>a</th><th>b</th></tr><tr><td>1</td><td>2</td></tr></table>')).toBe(
      '| a | b |\n| --- | --- |\n| 1 | 2 |',
    );
  });

  it('renders images as [image] (emoji as their character) and plain spans as their text', () => {
    expect(
      htmlToMarkdown(
        '<p><span style="x">see</span> <img src="https://us-api.asm.skype.com/v1/objects/x/views/imgo"> ' +
          '<img itemtype="http://schema.skype.com/Emoji" alt="🙂"></p>',
      ),
    ).toBe('see [image] 🙂');
  });

  it('keeps line breaks inside a paragraph and decodes entities', () => {
    expect(htmlToMarkdown('<p>a &amp; b<br>c&nbsp;d</p>')).toBe('a & b\nc d');
  });
});
