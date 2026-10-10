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

  it('classifies 48:notes as the self-chat', () => {
    expect(conversationKind({ id: '48:notes' })).toBe('self');
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

  it.each(['RichText/Media_CallRecording', 'RichText/Media_CallTranscript', 'RichText/Media_Card', 'Event/Call', 'ThreadActivity/DeleteMember', 'Control/Typing'])(
    'rejects %s as a system message',
    (messagetype) => {
      expect(isUserMessage({ messagetype })).toBe(false);
    },
  );
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

  it('is false for a call recording or transcript', () => {
    const m = message('1760000300000');
    expect(isUnread({ ...m, messagetype: 'RichText/Media_CallRecording' }, horizon, ME)).toBe(false);
    expect(isUnread({ ...m, messagetype: 'RichText/Media_CallTranscript' }, horizon, ME)).toBe(false);
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

  it('surfaces reactions from an emotions JSON string, with a count and whether the user reacted', () => {
    const OTHER = '8:orgid:00000000-0000-0000-0000-0000000000b2';
    const emotions = JSON.stringify([
      { key: 'like', users: [{ mri: OTHER, time: 1 }, { mri: ME, time: 2 }] },
      { key: 'heart', users: [{ mri: OTHER, time: 3 }] },
      { key: 'laugh', users: [] },
    ]);
    const n = normaliseMessage({ id: '1', messagetype: 'Text', properties: { emotions } }, ME);
    expect(n.reactions).toEqual([
      { key: 'like', count: 2, byMe: true },
      { key: 'heart', count: 1, byMe: false },
    ]);
  });

  it('accepts emotions already parsed as an array', () => {
    const n = normaliseMessage(
      { id: '1', messagetype: 'Text', properties: { emotions: [{ key: '1f389_partypopper', users: [{ mri: ME }] }] } },
      ME,
    );
    expect(n.reactions).toEqual([{ key: '1f389_partypopper', count: 1, byMe: true }]);
  });

  it('omits reactions when there are none or emotions is unreadable', () => {
    expect(normaliseMessage({ id: '1', messagetype: 'Text' }, ME)).not.toHaveProperty('reactions');
    expect(
      normaliseMessage({ id: '1', messagetype: 'Text', properties: { emotions: 'not json' } }, ME),
    ).not.toHaveProperty('reactions');
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
    expect(s.lastActivity).toBe('2026-10-09T12:08:20.000Z');
  });

  it('never calls a conversation unread on a system event alone: with no lastimreceivedtime it is uncertain, not unread', () => {
    // A real message may sit between the horizon and the event; only a fetch can tell.
    const s = conversationSummary(conv(CHANNEL), ME, CONVS);
    expect(s.unread).toBe(false);
    expect(s.unread_uncertain).toBe(true);
  });

  it('previews a system-event last message as "(system event)" with no sender', () => {
    const s = conversationSummary(conv(CHANNEL), ME, CONVS);
    expect(s.lastMessagePreview).toBe('(system event)');
    expect(s).not.toHaveProperty('lastSender');
  });

  it('calls a conversation whose last message was deleted after the horizon uncertain, not unread', () => {
    const c: RawConversation = {
      ...conv(GROUP),
      lastMessage: { ...conv(GROUP).lastMessage!, properties: { deletetime: '1760000100001' } },
    };
    expect(conversationSummary(c, ME, CONVS)).toMatchObject({ unread: false, unread_uncertain: true });
  });

  describe('unread from the last real message (lastimreceivedtime)', () => {
    const HORIZON = 1_760_000_000_000;
    const withLast = (
      messagetype: string,
      lastim: string | number | undefined,
      opts: { fromMe?: boolean; deleted?: boolean } = {},
    ): RawConversation => {
      const base = conv(GROUP);
      return {
        ...base,
        properties: {
          ...base.properties,
          ...(lastim !== undefined ? { lastimreceivedtime: lastim as string } : {}),
        },
        lastMessage: {
          ...base.lastMessage!,
          messagetype,
          ...(opts.fromMe ? { from: `https://example.invalid/contacts/${ME}` } : {}),
          ...(opts.deleted ? { properties: { deletetime: '1760000100001' } } : {}),
        },
      };
    };
    const before = new Date(HORIZON - 60_000).toISOString();
    const after = new Date(HORIZON + 60_000).toISOString();

    it.each(['RichText/Media_CallRecording', 'RichText/Media_CallTranscript', 'Event/Call', 'ThreadActivity/AddMember', 'Control/Typing'])(
      'a %s last message after a read real message is read, and certain',
      (type) => {
        const s = conversationSummary(withLast(type, before), ME, CONVS);
        expect(s.unread).toBe(false);
        expect(s).not.toHaveProperty('unread_uncertain');
      },
    );

    it('a system last message behind a real message past the horizon is unread, and certain', () => {
      const s = conversationSummary(withLast('RichText/Media_CallRecording', after), ME, CONVS);
      expect(s.unread).toBe(true);
      expect(s).not.toHaveProperty('unread_uncertain');
    });

    it('reads lastimreceivedtime as epoch ms too, as a number or a numeric string', () => {
      expect(conversationSummary(withLast('Event/Call', HORIZON + 1), ME, CONVS).unread).toBe(true);
      expect(conversationSummary(withLast('Event/Call', String(HORIZON - 1)), ME, CONVS).unread).toBe(false);
    });

    it('ignores an unparseable lastimreceivedtime (falling back to uncertain)', () => {
      const s = conversationSummary(withLast('Event/Call', 'not a time'), ME, CONVS);
      expect(s).toMatchObject({ unread: false, unread_uncertain: true });
    });

    it('a deleted last message behind a read real message is read, and certain', () => {
      const s = conversationSummary(withLast('Text', before, { deleted: true }), ME, CONVS);
      expect(s.unread).toBe(false);
      expect(s).not.toHaveProperty('unread_uncertain');
    });

    it('a real last message of mine is never unread, whatever lastimreceivedtime says', () => {
      expect(conversationSummary(withLast('Text', after, { fromMe: true }), ME, CONVS).unread).toBe(false);
    });

    it('a real last message from someone else past the horizon is unread with or without lastimreceivedtime', () => {
      expect(conversationSummary(withLast('RichText/Html', after), ME, CONVS).unread).toBe(true);
      expect(conversationSummary(withLast('RichText/Html', undefined), ME, CONVS).unread).toBe(true);
    });
  });

  describe('hidden channels', () => {
    const interest = (c: RawConversation, value: string | undefined): RawConversation => ({
      ...c,
      memberProperties: value === undefined ? {} : { interest: value },
    });

    it('marks a channel the user is NotInterested in as hidden', () => {
      expect(conversationSummary(interest(conv(CHANNEL), 'NotInterested'), ME, CONVS).hidden).toBe(true);
    });

    it('leaves a visible channel (no interest, or another value) unmarked', () => {
      expect(conversationSummary(interest(conv(CHANNEL), undefined), ME, CONVS)).not.toHaveProperty('hidden');
      expect(conversationSummary(interest(conv(CHANNEL), 'Interested'), ME, CONVS)).not.toHaveProperty('hidden');
    });

    it('never marks a chat hidden', () => {
      expect(conversationSummary(interest(conv(GROUP), 'NotInterested'), ME, CONVS)).not.toHaveProperty('hidden');
    });
  });

  it('summarises 48:notes as the self-chat', () => {
    expect(conversationSummary({ id: '48:notes' }, ME)).toEqual({
      id: '48:notes',
      kind: 'self',
      title: 'Notes (self-chat)',
      unread: false,
    });
  });

  it('never calls a conversation unread when the last event past the horizon is the user\'s own', () => {
    const c: RawConversation = {
      ...conv(CHANNEL),
      lastMessage: { ...conv(CHANNEL).lastMessage!, from: `https://example.invalid/contacts/${ME}` },
    };
    const s = conversationSummary(c, ME, CONVS);
    expect(s.unread).toBe(false);
    expect(s).not.toHaveProperty('unread_uncertain');
  });

  it('is certain about a real last message: no unread_uncertain either way', () => {
    expect(conversationSummary(conv(GROUP), ME, CONVS)).not.toHaveProperty('unread_uncertain');
    expect(conversationSummary(conv(GROUP_READ), ME, CONVS)).not.toHaveProperty('unread_uncertain');
  });

  it('caps the last-message preview at 200 characters', () => {
    const c: RawConversation = {
      ...conv(GROUP),
      lastMessage: { ...conv(GROUP).lastMessage!, content: `<p>${'x'.repeat(500)}</p>` },
    };
    const preview = conversationSummary(c, ME, CONVS).lastMessagePreview!;
    expect(preview).toHaveLength(200);
    expect(preview.endsWith('…')).toBe(true);
  });

  it('lists a team\'s own thread as "General" in that team', () => {
    expect(conversationSummary(conv(TEAM), ME, CONVS)).toMatchObject({
      kind: 'channel',
      title: 'General',
      team: 'Example Team',
    });
  });

  it('names a team whose thread has no topic from spaceThreadTopic, then topicThreadTopic, else "(team)"', () => {
    const team = (tp: Record<string, string>): RawConversation => ({
      id: TEAM,
      threadProperties: { threadType: 'space', spaceId: TEAM, ...tp },
    });
    const teamOf = (t: RawConversation) => conversationSummary(conv(CHANNEL), ME, [t, conv(CHANNEL)]).team;

    expect(teamOf(team({ topic: ' ', spaceThreadTopic: 'Space Name', topicThreadTopic: 'Topic Name' }))).toBe('Space Name');
    expect(teamOf(team({ topicThreadTopic: 'Topic Name' }))).toBe('Topic Name');
    expect(teamOf(team({}))).toBe('(team)');
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

  describe('the self-chat', () => {
    const self = conversationSummary({ id: '48:notes' }, ME);
    const notesChat: ConversationSummary = { id: '19:notes0001@thread.v2', kind: 'group', title: 'Release notes', unread: false };
    const withSelf = [...summaries, notesChat, self];

    it.each(['48:notes', 'notes', 'NOTES', ' self ', 'Self'])('resolves %j to 48:notes alone', (q) => {
      expect(matchConversations(withSelf, q).map((s) => s.id)).toEqual(['48:notes']);
    });

    it('matches by substring like any other title otherwise', () => {
      expect(matchConversations(withSelf, 'self-chat').map((s) => s.id)).toEqual(['48:notes']);
      expect(matchConversations(withSelf, 'release').map((s) => s.id)).toEqual(['19:notes0001@thread.v2']);
    });

    it('treats "notes" as an ordinary substring when no self-chat is listed', () => {
      expect(matchConversations([...summaries, notesChat], 'notes').map((s) => s.id)).toEqual(['19:notes0001@thread.v2']);
    });
  });
});
