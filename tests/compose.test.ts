import { describe, expect, it } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import {
  buildSendBody,
  emojiSpan,
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

describe('emojiSpan', () => {
  it('renders a known named key as the captured emoji span', () => {
    expect(emojiSpan('smile')).toBe(
      '<span contenteditable="false" title="Smile" type="(smile)" class="animated-emoticon-20-smile" itemscope>' +
        '<img itemscope itemtype="http://schema.skype.com/Emoji" itemid="smile" ' +
        'src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/smile/default/20_f.png" ' +
        'title="Smile" alt="🙂" style="width:20px;height:20px;"></span>',
    );
  });

  it('derives the character and title of a unicode-style key', () => {
    const span = emojiSpan('1f389_partypopper')!;
    expect(span).toContain('alt="🎉"');
    expect(span).toContain('title="partypopper"');
    expect(span).toContain('itemid="1f389_partypopper"');
  });

  it('returns undefined for a key it cannot render', () => {
    expect(emojiSpan('notakey')).toBeUndefined();
    expect(emojiSpan('30')).toBeUndefined();
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
