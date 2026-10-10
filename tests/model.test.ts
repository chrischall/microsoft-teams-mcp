import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import type { RawConversation, RawMessage } from '../src/chatsvc.js';
import {
  conversationKind,
  conversationSummary,
  htmlToText,
  isUnread,
  isUserMessage,
  matchConversations,
  normaliseMessage,
  readHorizon,
  senderMri,
  type ConversationSummary,
} from '../src/model.js';

/** A hand-written fixture from `tests/fixtures/` — invented names and text only. */
function fixture<T>(name: string): T {
  return JSON.parse(readFileSync(new URL(`./fixtures/${name}`, import.meta.url), 'utf8')) as T;
}

const ME = '8:orgid:00000000-0000-0000-0000-0000000000a1';
const CONVS = fixture<{ conversations: RawConversation[] }>('model-conversations.json').conversations;
const MSGS = fixture<{ messages: RawMessage[] }>('model-messages.json').messages;

function conv(id: string): RawConversation {
  const c = CONVS.find((x) => x.id === id);
  if (!c) throw new Error(`no fixture conversation ${id}`);
  return c;
}
function message(id: string): RawMessage {
  const m = MSGS.find((x) => x.id === id);
  if (!m) throw new Error(`no fixture message ${id}`);
  return m;
}

const GROUP = '19:group0001@thread.v2';
const GROUP_READ = '19:group0002@thread.v2';
const ONE_ON_ONE = '19:00000000-0000-0000-0000-0000000000a1_00000000-0000-0000-0000-0000000000b2@unq.gbl.spaces';
const TEAM = '19:team0001@thread.tacv2';
const CHANNEL = '19:channel0001@thread.tacv2';
const ORPHAN_CHANNEL = '19:channel0002@thread.tacv2';

describe('conversationKind', () => {
  it('classifies a @thread.tacv2 id as a channel', () => {
    expect(conversationKind(conv(CHANNEL))).toBe('channel');
  });

  it('classifies a 19:meeting_ id as a meeting', () => {
    expect(conversationKind(conv('19:meeting_ZXhhbXBsZQ@thread.v2'))).toBe('meeting');
  });

  it('classifies a meeting by productThreadType or threadType', () => {
    expect(conversationKind(conv('19:group0003@thread.v2'))).toBe('meeting');
    expect(conversationKind(conv('19:group0004@thread.v2'))).toBe('meeting');
  });

  it('classifies a @unq.gbl.spaces id as a 1:1', () => {
    expect(conversationKind(conv(ONE_ON_ONE))).toBe('oneOnOne');
  });

  it('classifies any other thread as a group chat', () => {
    expect(conversationKind(conv(GROUP))).toBe('group');
  });
});

describe('readHorizon', () => {
  it('reads the first consumptionhorizon segment as a number', () => {
    expect(readHorizon(conv(GROUP))).toBe(1760000000000);
  });

  it('is 0 when the horizon is absent', () => {
    expect(readHorizon(conv('19:group0003@thread.v2'))).toBe(0);
    expect(readHorizon({ id: 'x' })).toBe(0);
  });

  it('is 0 when the horizon is unparseable', () => {
    expect(readHorizon(conv(TEAM))).toBe(0);
  });
});

describe('senderMri', () => {
  it('takes the MRI off the end of the from URL', () => {
    expect(senderMri(message('1760000050000').from)).toBe(ME);
  });

  it('is undefined for a missing from', () => {
    expect(senderMri(undefined)).toBeUndefined();
  });
});

describe('isUserMessage', () => {
  it('accepts Text and RichText/Html', () => {
    expect(isUserMessage(message('1760000050000'))).toBe(true);
    expect(isUserMessage(message('1760000300000'))).toBe(true);
  });

  it('rejects system events', () => {
    expect(isUserMessage(message('1760000240000'))).toBe(false);
  });
});

describe('isUnread', () => {
  const horizon = 1760000000000;

  it('is true for someone else\'s message newer than the horizon', () => {
    expect(isUnread(message('1760000300000'), horizon, ME)).toBe(true);
  });

  it('is false at or below the horizon', () => {
    expect(isUnread(message('1760000300000'), 1760000300000, ME)).toBe(false);
  });

  it('is false for the user\'s own message', () => {
    expect(isUnread(message('1760000050000'), 0, ME)).toBe(false);
  });

  it('is false for a system message', () => {
    expect(isUnread(message('1760000240000'), horizon, ME)).toBe(false);
  });

  it('is false for a deleted message', () => {
    expect(isUnread(message('1760000250000'), horizon, ME)).toBe(false);
  });

  it('is false for a non-numeric id', () => {
    expect(isUnread({ ...message('1760000300000'), id: 'abc' }, horizon, ME)).toBe(false);
  });
});

describe('htmlToText', () => {
  it('turns <br>, </p> and </div> into newlines and strips other tags', () => {
    expect(htmlToText('<div>one<br>two<br/>three</div><p><b>four</b></p>')).toBe('one\ntwo\nthree\nfour');
  });

  it('decodes the named and numeric entities', () => {
    expect(htmlToText('a &amp; b &lt;c&gt; &quot;d&quot; &#39;e&#39;&nbsp;f &#8364; &#x41;')).toBe(
      'a & b <c> "d" \'e\' f € A',
    );
  });

  it('decodes once, so an escaped entity stays escaped', () => {
    expect(htmlToText('&amp;lt;b&amp;gt;')).toBe('&lt;b&gt;');
  });

  it('keeps link targets as "text (url)"', () => {
    expect(htmlToText('see <a href="https://example.com/a?x=1&amp;y=2">the doc</a>')).toBe(
      'see the doc (https://example.com/a?x=1&y=2)',
    );
  });

  it('does not repeat a link whose text is its url', () => {
    expect(htmlToText('<a href="https://example.com">https://example.com</a>')).toBe('https://example.com');
  });

  it('replaces an <img> with [image], but an emoji with its alt text', () => {
    expect(htmlToText('look <img src="https://us-api.asm.skype.com/v1/objects/x/views/imgo">')).toBe('look [image]');
    expect(
      htmlToText('ok <img itemtype="http://schema.skype.com/Emoji" alt="🙂" src="https://statics.example/e.png">'),
    ).toBe('ok 🙂');
  });

  it('collapses blank lines and trims', () => {
    expect(htmlToText('<p>one</p><p></p><p>&nbsp;</p>\n\n<p>two</p>  ')).toBe('one\ntwo');
  });

  it('renders a quoted reply as "> " lines before the answer', () => {
    expect(htmlToText(message('1760000300000').content ?? '')).toBe(
      '> Example Person\n> Original question\nHere is the answer',
    );
  });

  it('leaves plain text with a bare < alone', () => {
    expect(htmlToText('a < b and c > d')).toBe('a < b and c > d');
  });
});

describe('normaliseMessage', () => {
  it('normalises a quoted-reply, edited chat message', () => {
    expect(normaliseMessage(message('1760000300000'), ME)).toEqual({
      id: '1760000300000',
      time: '2026-10-09T12:05:00.000Z',
      sender: 'Another Person',
      fromMe: false,
      text: '> Example Person\n> Original question\nHere is the answer',
      edited: true,
      attachments: { files: [], images: [] },
    });
  });

  it('marks a deleted message', () => {
    const n = normaliseMessage(message('1760000250000'), ME);
    expect(n.deleted).toBe(true);
    expect(n.text).toBe('');
  });

  it('sets replyTo for a channel reply', () => {
    expect(normaliseMessage(message('1760000230000'), ME).replyTo).toBe('1760000100000');
  });

  it('extracts subject, files and inline images from a channel root post (no replyTo)', () => {
    const n = normaliseMessage(message('1760000100000'), ME);
    expect(n.replyTo).toBeUndefined();
    expect(n.subject).toBe('Weekly update');
    expect(n.text).toBe('Root post with a file and a picture [image]');
    expect(n.attachments).toEqual({
      files: [
        {
          name: 'example-notes.docx',
          type: 'docx',
          openUrl: 'https://example.sharepoint.com/sites/example/_layouts/15/Doc.aspx?sourcedoc=x',
          url: 'https://example.sharepoint.com/sites/example/Shared%20Documents/example-notes.docx',
        },
        {
          name: 'diagram.png',
          type: 'png',
          openUrl: 'https://example.sharepoint.com/:i:/s/share',
          url: 'https://example.sharepoint.com/diagram.png',
        },
      ],
      images: [
        { ref: '0-eus-d1-abc123', url: 'https://us-api.asm.skype.com/v1/objects/0-eus-d1-abc123/views/imgo' },
      ],
    });
  });

  it('normalises a 1:1 message from me, timed from composetime', () => {
    expect(normaliseMessage(message('1760000050000'), ME)).toEqual({
      id: '1760000050000',
      time: '2026-10-09T12:00:50.000Z',
      sender: 'Test User',
      fromMe: true,
      text: 'A note from me',
      edited: false,
      attachments: { files: [], images: [] },
    });
  });

  it('falls back to the id for time and the MRI for sender', () => {
    const n = normaliseMessage(
      { id: '1760000000000', messagetype: 'Text', from: 'https://x/contacts/8:orgid:zzz', content: 'hi' },
      ME,
    );
    expect(n.time).toBe(new Date(1760000000000).toISOString());
    expect(n.sender).toBe('8:orgid:zzz');
  });

  it('ignores a files property that is not valid JSON', () => {
    const n = normaliseMessage({ id: '1', messagetype: 'Text', properties: { files: 'not json' } }, ME);
    expect(n.attachments.files).toEqual([]);
  });

  it('only lists images hosted on the Teams image store, once each', () => {
    const n = normaliseMessage(
      {
        id: '1',
        messagetype: 'RichText/Html',
        content:
          '<img src="https://evil.example/v1/objects/bad/views/imgo">' +
          '<img src="https://eu-api.asm.skype.com/v1/objects/good/views/imgo">' +
          '<img src="https://eu-api.asm.skype.com/v1/objects/good/views/imgo">',
      },
      ME,
    );
    expect(n.attachments.images).toEqual([
      { ref: 'good', url: 'https://eu-api.asm.skype.com/v1/objects/good/views/imgo' },
    ]);
  });
});

describe('conversationSummary', () => {
  it('summarises an unread group chat with a topic', () => {
    expect(conversationSummary(conv(GROUP), ME, CONVS)).toEqual({
      id: GROUP,
      kind: 'group',
      title: 'Example project group',
      unread: true,
      lastActivity: '2026-10-09T12:01:40.000Z',
      lastMessagePreview: 'Example & message',
      lastSender: 'Example Person',
    });
  });

  it('titles an untopiced chat after the other participant and marks it read', () => {
    const s = conversationSummary(conv(GROUP_READ), ME, CONVS);
    expect(s.title).toBe('Another Person');
    expect(s.unread).toBe(false);
  });

  it('titles a chat whose last message is mine "(untitled chat)" and never unread', () => {
    const s = conversationSummary(conv(ONE_ON_ONE), ME, CONVS);
    expect(s.kind).toBe('oneOnOne');
    expect(s.title).toBe('(untitled chat)');
    expect(s.unread).toBe(false);
    expect(s.lastSender).toBe('Test User');
  });

  it('has no lastActivity/preview for a conversation without a last message', () => {
    const s = conversationSummary(conv('19:meeting_ZXhhbXBsZQ@thread.v2'), ME, CONVS);
    expect(s).toMatchObject({ kind: 'meeting', title: 'Example standup', unread: false });
    expect(s.lastActivity).toBeUndefined();
    expect(s.lastMessagePreview).toBeUndefined();
    expect(s.lastSender).toBeUndefined();
  });

  it('names a channel\'s team from the spaceId conversation', () => {
    const s = conversationSummary(conv(CHANNEL), ME, CONVS);
    expect(s).toMatchObject({ kind: 'channel', title: 'Design', team: 'Example Team' });
    // Only a system event since the horizon: not unread.
    expect(s.unread).toBe(false);
    expect(s.lastActivity).toBe('2026-10-09T12:08:20.000Z');
  });

  it('titles a topic-less channel "General" and omits an unknown team', () => {
    const s = conversationSummary(conv(ORPHAN_CHANNEL), ME, CONVS);
    expect(s.title).toBe('General');
    expect(s).not.toHaveProperty('team');
  });

  it('works without the full list (no team)', () => {
    expect(conversationSummary(conv(CHANNEL), ME)).not.toHaveProperty('team');
  });
});

describe('matchConversations', () => {
  const summaries: ConversationSummary[] = CONVS.map((c) => conversationSummary(c, ME, CONVS));

  it('lets an exact id win over substring matches', () => {
    expect(matchConversations(summaries, CHANNEL).map((s) => s.id)).toEqual([CHANNEL]);
  });

  it('matches title case-insensitively', () => {
    expect(matchConversations(summaries, 'PROJECT').map((s) => s.id)).toEqual([GROUP]);
  });

  it('matches team name', () => {
    expect(matchConversations(summaries, 'example team').map((s) => s.id)).toEqual(
      expect.arrayContaining([TEAM, CHANNEL]),
    );
  });

  it('returns nothing for no match', () => {
    expect(matchConversations(summaries, 'nothing like this')).toEqual([]);
  });
});
