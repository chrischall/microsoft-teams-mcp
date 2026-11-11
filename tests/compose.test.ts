import { describe, expect, it } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import {
  buildSendBody,
  clipText,
  emojiHtml,
  escapeHtml,
  mentionsProperty,
  newClientMessageId,
  qtdMsgsProperty,
  quoteHtml,
  renderMessageHtml,
} from '../src/compose.js';
import { htmlToText } from '../src/model.js';

const ME = '8:orgid:00000000-0000-0000-0000-0000000000a1';
const OTHER = '8:orgid:00000000-0000-0000-0000-0000000000b2';
const THIRD = '8:orgid:00000000-0000-0000-0000-0000000000c3';

describe('escapeHtml', () => {
  it('escapes the five HTML-significant characters', () => {
    expect(escapeHtml(`<b>"Tom" & 'Jerry'</b>`)).toBe('&lt;b&gt;&quot;Tom&quot; &amp; &#39;Jerry&#39;&lt;/b&gt;');
  });
});

describe('emojiHtml', () => {
  it('renders a known named key as the captured emoji span', () => {
    expect(emojiHtml('smile')).toBe(
      '<span contenteditable="false" title="Smile" type="(smile)" class="animated-emoticon-20-smile" itemscope>' +
        '<img itemscope itemtype="http://schema.skype.com/Emoji" itemid="smile" ' +
        'src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/smile/default/20_f.png" ' +
        'title="Smile" alt="🙂" style="width:20px;height:20px;"></span>',
    );
  });

  it('renders a unicode-style key as its plain character, never an image Teams may not have', () => {
    // Live 2026-11-11: Teams' CDN has no image for 1f680_rocket, so a span showed a broken image.
    expect(emojiHtml('1f680_rocket')).toBe('🚀');
    expect(emojiHtml('1f389_partypopper')).toBe('🎉');
    expect(emojiHtml('1f468-200d-1f4bb_technologist')).toBe('👨\u200d💻');
  });

  it('renders a common name as its plain character', () => {
    expect(emojiHtml('rocket')).toBe('🚀');
    expect(emojiHtml('tada')).toBe('🎉');
    expect(emojiHtml('thumbsup')).toBe('👍');
    expect(emojiHtml('white_check_mark')).toBe('✅');
  });

  it('keeps the captured span for every built-in name (all 12 images exist on Teams\' CDN, checked 2026-11-11)', () => {
    for (const key of ['like', 'heart', 'laugh', 'surprised', 'sad', 'angry', 'smile', 'wink', 'cool', 'hearteyes', 'fire', 'party']) {
      expect(emojiHtml(key), key).toContain(`itemid="${key}"`);
    }
  });

  it('returns undefined for a key it cannot render', () => {
    expect(emojiHtml('notakey')).toBeUndefined();
    expect(emojiHtml('30')).toBeUndefined();
    expect(emojiHtml('110000_toohigh')).toBeUndefined();
  });
});

describe('renderMessageHtml', () => {
  it('escapes plain text, turns newlines into <br> and wraps it in one paragraph', () => {
    expect(renderMessageHtml('a < b & c\nnext line\r\nlast', []).html).toBe(
      '<p>a &lt; b &amp; c<br>next line<br>last</p>',
    );
  });

  it('never lets the caller inject markup', () => {
    const { html } = renderMessageHtml('<img src=x onerror=alert(1)>', []);
    expect(html).not.toContain('<img');
    expect(htmlToText(html)).toBe('<img src=x onerror=alert(1)>');
  });

  it('turns a known :key: shortcode into an emoji span and leaves anything else as text', () => {
    const { html } = renderMessageHtml('done :smile: at 10:30:45 :nope:', []);
    expect(html).toContain('itemid="smile"');
    expect(html).toContain('at 10:30:45 :nope:');
    expect(htmlToText(html)).toBe('done 🙂 at 10:30:45 :nope:');
  });

  it('gives an unrenderable :key: its closing colon back, so the next shortcode still renders', () => {
    const { html } = renderMessageHtml('a:b:smile:', []);
    expect(html).toContain('itemid="smile"');
    expect(htmlToText(html)).toBe('a:b🙂');
    expect(htmlToText(renderMessageHtml(':smile:fire:', []).html)).toBe('🙂fire:');
    expect(htmlToText(renderMessageHtml('at 10:30:smile: ok', []).html)).toBe('at 10:30🙂 ok');
  });

  it('never matches @name inside a longer name: "@Sam" skips "@Samantha"', () => {
    const { html, mentions } = renderMessageHtml('@Samantha and @Sam', [{ text: 'Sam', mri: OTHER }]);
    expect(html).toBe(
      '<p>@Samantha and <span itemscope itemtype="http://schema.skype.com/Mention" itemid="0">Sam</span></p>',
    );
    expect(mentions).toEqual([{ itemid: '0', mri: OTHER, displayName: 'Sam' }]);
  });

  it('applies the trailing boundary to the case-insensitive @name search too', () => {
    const { html } = renderMessageHtml('@samantha, @sam!', [{ text: 'Sam', mri: OTHER }]);
    expect(html).toContain('@samantha, <span');
    expect(html).toContain('itemid="0">sam</span>!');
  });

  it('replaces the first @name with a mention span, numbering mentions in order of appearance', () => {
    const { html, mentions } = renderMessageHtml('hi @Bea and @Al, Al again', [
      { text: 'Bea', mri: OTHER },
      { text: 'Al', mri: THIRD },
    ]);
    expect(html).toBe(
      '<p>hi <span itemscope itemtype="http://schema.skype.com/Mention" itemid="0">Bea</span> and ' +
        '<span itemscope itemtype="http://schema.skype.com/Mention" itemid="1">Al</span>, Al again</p>',
    );
    expect(mentions).toEqual([
      { itemid: '0', mri: OTHER, displayName: 'Bea' },
      { itemid: '1', mri: THIRD, displayName: 'Al' },
    ]);
  });

  it('numbers by position in the text, not by argument order', () => {
    const { mentions } = renderMessageHtml('Al then Bea', [
      { text: 'Bea', mri: OTHER },
      { text: 'Al', mri: THIRD },
    ]);
    expect(mentions.map((m) => [m.displayName, m.itemid])).toEqual([
      ['Al', '0'],
      ['Bea', '1'],
    ]);
  });

  it('falls back to the bare name, case-insensitively, when there is no @name', () => {
    const { html } = renderMessageHtml('thanks bea!', [{ text: 'Bea', mri: OTHER }]);
    expect(html).toContain('itemid="0">bea</span>!');
  });

  it('escapes a mention name', () => {
    const { html } = renderMessageHtml('@A&B', [{ text: 'A&B', mri: OTHER }]);
    expect(html).toContain('itemid="0">A&amp;B</span>');
  });

  it('throws when a mention name is not in the text', () => {
    expect(() => renderMessageHtml('hello', [{ text: 'Bea', mri: OTHER }])).toThrow(McpToolError);
    expect(() => renderMessageHtml('hello', [{ text: 'Bea', mri: OTHER }])).toThrow(/"Bea" does not appear/);
  });

  it('gives two mentions of one name two separate occurrences', () => {
    const { mentions, html } = renderMessageHtml('@Al and @Al', [
      { text: 'Al', mri: THIRD },
      { text: 'Al', mri: OTHER },
    ]);
    expect(mentions.map((m) => m.mri)).toEqual([THIRD, OTHER]);
    expect(html.match(/schema\.skype\.com\/Mention/g)).toHaveLength(2);
  });
});

describe('renderMessageHtml (markdown)', () => {
  const md = (text: string, mentions: { text: string; mri: string }[] = []) => renderMessageHtml(text, mentions, 'markdown');
  const MENTION = (n: number, name: string) =>
    `<span itemscope itemtype="http://schema.skype.com/Mention" itemid="${n}">${name}</span>`;

  it('renders bold, italic and strike', () => {
    expect(md('**bold** and _it_ and *it2* and ~~gone~~').html).toBe(
      '<p><strong>bold</strong> and <em>it</em> and <em>it2</em> and <s>gone</s></p>',
    );
  });

  it('leaves underscores inside words and a lone asterisk alone', () => {
    expect(md('snake_case_name costs 2 * 3').html).toBe('<p>snake_case_name costs 2 * 3</p>');
  });

  it('keeps line breaks in a paragraph and splits paragraphs on a blank line', () => {
    expect(md('a\nb\n\nc').html).toBe('<p>a<br>b</p><p>c</p>');
  });

  it('renders numbered and bulleted lists, keeping a numbered list\'s start', () => {
    expect(md('Steps:\n1. one\n2. two\n\n- a\n* b').html).toBe(
      '<p>Steps:</p><ol><li>one</li><li>two</li></ol><ul><li>a</li><li>b</li></ul>',
    );
    expect(md('3. three\n4. four').html).toBe('<ol start="3"><li>three</li><li>four</li></ol>');
  });

  it('nests an indented list under its item', () => {
    expect(md('1. top\n   - inner\n   - more\n2. next').html).toBe(
      '<ol><li>top<ul><li>inner</li><li>more</li></ul></li><li>next</li></ol>',
    );
  });

  it('continues an item on an indented line, and switches list kind at a nested level', () => {
    expect(md('1. first\n   still first\n2. second').html).toBe('<ol><li>first<br>still first</li><li>second</li></ol>');
    expect(md('1. top\n   - bullet\n   1. numbered').html).toBe(
      '<ol><li>top<ul><li>bullet</li></ul><ol><li>numbered</li></ol></li></ol>',
    );
    expect(md('- a\n1. b').html).toBe('<ul><li>a</li></ul><ol><li>b</li></ol>');
  });

  it('renders a link with its text, escaping the address', () => {
    expect(md('see [the **doc**](https://example.com/a?b=1&c=2)').html).toBe(
      '<p>see <a href="https://example.com/a?b=1&amp;c=2">the <strong>doc</strong></a></p>',
    );
  });

  it('leaves a link to anything but http(s) or mailto as literal text', () => {
    expect(md('[x](javascript:alert(1))').html).toBe('<p>[x](javascript:alert(1))</p>');
  });

  it('renders inline code verbatim and escaped', () => {
    expect(md('run `a<b && **c**` now').html).toBe('<p>run <code>a&lt;b &amp;&amp; **c**</code> now</p>');
  });

  it('renders a fenced code block verbatim', () => {
    expect(md('before\n```js\nif (a < b) {\n  go(); // **x**\n}\n```\nafter').html).toBe(
      '<p>before</p><pre><code>if (a &lt; b) {\n  go(); // **x**\n}</code></pre><p>after</p>',
    );
  });

  it('escapes raw HTML and leaves unsupported markdown as text', () => {
    expect(md('<b>x</b> <img src=x onerror=alert(1)>').html).toBe(
      '<p>&lt;b&gt;x&lt;/b&gt; &lt;img src=x onerror=alert(1)&gt;</p>',
    );
    expect(md('# not a heading\n> not a quote\n![img](https://example.com/i.png)').html).toBe(
      '<p># not a heading<br>&gt; not a quote<br>!<a href="https://example.com/i.png">img</a></p>',
    );
  });

  it('honours backslash escapes', () => {
    expect(md('\\*not italic\\* and \\_nor this\\_').html).toBe('<p>*not italic* and _nor this_</p>');
  });

  it('renders :shortcodes: as emoji outside code only', () => {
    const { html } = md(':smile: **hi** `:smile:`');
    expect(html).toContain('itemid="smile"');
    expect(html).toContain('<strong>hi</strong>');
    expect(html).toContain('<code>:smile:</code>');
    expect(md(':1f389_partypopper: :rocket:_ship_').html).toBe('<p>🎉 🚀<em>ship</em></p>');
  });

  it('places mentions inside formatting and lists, numbered in order of appearance', () => {
    const { html, mentions } = md('**@Bea** see:\n1. ask @Al', [
      { text: 'Al', mri: THIRD },
      { text: 'Bea', mri: OTHER },
    ]);
    expect(html).toBe(`<p><strong>${MENTION(0, 'Bea')}</strong> see:</p><ol><li>ask ${MENTION(1, 'Al')}</li></ol>`);
    expect(mentions).toEqual([
      { itemid: '0', mri: OTHER, displayName: 'Bea' },
      { itemid: '1', mri: THIRD, displayName: 'Al' },
    ]);
  });

  it('cannot be tricked into a mention by private-use characters in the text', () => {
    expect(md('0 x').html).toBe('<p>0 x</p>');
  });

  it('reads back with the structure kept', () => {
    expect(htmlToText(md('**Plan**\n1. one\n2. two\n\nsee [doc](https://example.com)').html)).toBe(
      'Plan\n1. one\n2. two\nsee doc (https://example.com)',
    );
  });
});

describe('mentionsProperty', () => {
  it('is the documented JSON string', () => {
    expect(JSON.parse(mentionsProperty([{ itemid: '0', mri: OTHER, displayName: 'Bea' }]))).toEqual([
      {
        '@type': 'http://schema.skype.com/Mention',
        itemid: '0',
        mri: OTHER,
        mentionType: 'person',
        displayName: 'Bea',
      },
    ]);
  });
});

describe('quoteHtml / qtdMsgsProperty', () => {
  const quoted = { id: '1760000000000', senderMri: OTHER, senderName: 'Bea <B>', text: 'see <this> & that' };

  it('renders the captured Reply blockquote with an escaped preview', () => {
    expect(quoteHtml(quoted)).toBe(
      `<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="1760000000000">` +
        `<strong itemprop="mri" itemid="${OTHER}">Bea &lt;B&gt;</strong>` +
        `<span itemprop="time" itemid="1760000000000"></span>` +
        `<p itemprop="preview">see &lt;this&gt; &amp; that</p></blockquote>`,
    );
  });

  it('caps the preview at 200 characters', () => {
    const html = quoteHtml({ ...quoted, text: 'x'.repeat(500) });
    const preview = /<p itemprop="preview">([^<]*)<\/p>/.exec(html)![1];
    expect(preview.length).toBeLessThanOrEqual(200);
  });

  it('never splits a surrogate pair at the cut', () => {
    // 198 x's then an emoji straddling the 199th/200th unit: cut before it, whole.
    const html = quoteHtml({ ...quoted, text: `${'x'.repeat(198)}🎉🎉 tail` });
    const preview = /<p itemprop="preview">([^<]*)<\/p>/.exec(html)![1];
    expect(preview).toBe(`${'x'.repeat(198)}…`);
    expect(preview).not.toMatch(/[\ud800-\udbff](?![\udc00-\udfff])/);
  });

  it('lists the quoted message as a JSON string', () => {
    expect(JSON.parse(qtdMsgsProperty(quoted))).toEqual([
      { messageId: '1760000000000', sender: OTHER, time: 1760000000000 },
    ]);
  });
});

describe('newClientMessageId', () => {
  it('is a random 19-digit numeric string', () => {
    const a = newClientMessageId();
    expect(a).toMatch(/^[1-9]\d{18}$/);
    expect(newClientMessageId()).not.toBe(a);
  });
});

describe('buildSendBody', () => {
  it('builds the captured send shape, properties as JSON strings', () => {
    const now = new Date('2026-10-10T12:00:00.000Z');
    const body = buildSendBody({
      conversationId: '19:group0001@thread.v2',
      meMri: ME,
      meName: 'Me Myself',
      content: '<p>hi</p>',
      clientmessageid: '1234567890123456789',
      now,
    });
    expect(body).toEqual({
      id: '-1',
      type: 'Message',
      conversationid: '19:group0001@thread.v2',
      conversationLink: 'blah/19:group0001@thread.v2',
      from: ME,
      fromUserId: ME,
      composetime: now.toISOString(),
      originalarrivaltime: now.toISOString(),
      content: '<p>hi</p>',
      messagetype: 'RichText/Html',
      contenttype: 'Text',
      imdisplayname: 'Me Myself',
      clientmessageid: '1234567890123456789',
      callId: '',
      state: 0,
      version: '0',
      amsreferences: [],
      properties: {
        importance: '',
        subject: '',
        title: '',
        cards: '[]',
        links: '[]',
        mentions: '[]',
        onbehalfof: null,
        files: '[]',
        policyViolation: null,
        formatVariant: 'TEAMS',
      },
      crossPostChannels: [],
    });
  });

  it('carries mentions and a quote when given', () => {
    const body = buildSendBody({
      conversationId: 'c',
      meMri: ME,
      meName: '',
      content: '',
      clientmessageid: '1',
      now: new Date(0),
      mentions: '[{"x":1}]',
      qtdMsgs: '[{"y":2}]',
    });
    expect(body.properties.mentions).toBe('[{"x":1}]');
    expect(body.properties.qtdMsgs).toBe('[{"y":2}]');
  });
});

describe('clipText', () => {
  it('returns short text unchanged', () => {
    expect(clipText('hello', 200)).toBe('hello');
  });

  it('cuts at max units, with the ellipsis inside the limit', () => {
    expect(clipText('abcdef', 4)).toBe('abcd');
    expect(clipText('abcdef', 4, '…')).toBe('abc…');
  });

  it('drops a whole emoji rather than leaving half a surrogate pair', () => {
    expect(clipText('ab🎉cd', 3)).toBe('ab');
    expect(clipText('ab🎉cd', 4)).toBe('ab🎉');
    expect(clipText('ab🎉cd', 4, '…')).toBe('ab…');
  });
});
