import { describe, it, expect } from 'vitest';
import { formatBody, formatConversation } from '../../src/tools/format.js';
import type { ConversationSummary, NormalisedMessage } from '../../src/model.js';

const msg = (id: number, over: Partial<NormalisedMessage> = {}): NormalisedMessage => ({
  id: String(id),
  time: new Date(id).toISOString(),
  sender: 'Example Person',
  fromMe: false,
  text: `message ${id}`,
  edited: false,
  attachments: { files: [], images: [] },
  ...over,
});

const T0 = 1_760_000_000_000;
const MIN = 60_000;

const SUMMARY: ConversationSummary = {
  id: '19:chan0001@thread.tacv2',
  kind: 'channel',
  title: 'Design',
  team: 'Platform Team',
  unread: true,
  unread_confirmed: true,
  favorite: true,
  lastActivity: new Date(T0).toISOString(),
  lastMessagePreview: 'a long preview',
  lastSender: 'Example Person',
};

describe('formatConversation', () => {
  it('keeps the whole summary in full', () => {
    expect(formatConversation(SUMMARY, 'full')).toEqual(SUMMARY);
  });

  it('slims it to id, kind, title and team in compact and summary', () => {
    for (const format of ['compact', 'summary'] as const) {
      expect(formatConversation(SUMMARY, format)).toEqual({
        id: SUMMARY.id,
        kind: 'channel',
        title: 'Design',
        team: 'Platform Team',
      });
    }
  });
});

describe('formatBody', () => {
  const chat = [
    msg(T0, {
      subject: 'Heads up',
      edited: true,
      reactions: [{ key: 'like', count: 1, byMe: false }],
      attachments: { files: [{ name: 'a.pdf' }], images: [] },
      time_local: '2025-10-09T04:53:20.000-04:00',
    }),
    msg(T0 + MIN, { deleted: true, fromMe: true }),
  ];

  it('returns messages untouched in full', () => {
    expect(formatBody('group', chat, 'full')).toEqual({ messages: chat });
  });

  it('keeps only id, time, time_local, sender, text and a non-empty subject in compact', () => {
    expect(formatBody('group', chat, 'compact')).toEqual({
      messages: [
        {
          id: String(T0),
          time: chat[0].time,
          time_local: '2025-10-09T04:53:20.000-04:00',
          sender: 'Example Person',
          text: `message ${T0}`,
          subject: 'Heads up',
        },
        { id: String(T0 + MIN), time: chat[1].time, sender: 'Example Person', text: `message ${T0 + MIN}` },
      ],
    });
  });

  it('groups channel posts into compact threads, keeping a missing root as a placeholder', () => {
    const posts = [
      msg(T0, { replyTo: String(T0 - MIN) }),
      msg(T0 + MIN),
      msg(T0 + 2 * MIN, { replyTo: String(T0 + MIN), reactions: [{ key: 'like', count: 2, byMe: true }] }),
    ];
    expect(formatBody('channel', posts, 'compact')).toEqual({
      threads: [
        {
          root: { id: String(T0 - MIN), missing: true },
          replies: [{ id: String(T0), time: posts[0].time, sender: 'Example Person', text: `message ${T0}` }],
        },
        {
          root: { id: String(T0 + MIN), time: posts[1].time, sender: 'Example Person', text: `message ${T0 + MIN}` },
          replies: [{ id: String(T0 + 2 * MIN), time: posts[2].time, sender: 'Example Person', text: `message ${T0 + 2 * MIN}` }],
        },
      ],
    });
  });

  it('counts a chat in summary: message_count and the newest time, no messages', () => {
    expect(formatBody('group', chat, 'summary')).toEqual({ message_count: 2, newest_time: chat[1].time });
  });

  it('adds newest_time_local in summary when the messages carry time_local', () => {
    const local = [msg(T0, { time_local: 'L0' }), msg(T0 + MIN, { time_local: 'L1' })];
    expect(formatBody('oneOnOne', local, 'summary')).toEqual({
      message_count: 2,
      newest_time: local[1].time,
      newest_time_local: 'L1',
    });
  });

  it('says only message_count 0 for an empty read in summary', () => {
    expect(formatBody('group', [], 'summary')).toEqual({ message_count: 0 });
    expect(formatBody('channel', [], 'summary')).toEqual({ message_count: 0, threads: [] });
  });

  it('summarises each channel thread: root id, subject, reply count and newest time', () => {
    const posts = [
      msg(T0, { replyTo: String(T0 - MIN) }),
      msg(T0 + MIN, { subject: 'Release plan' }),
      msg(T0 + 2 * MIN, { replyTo: String(T0 + MIN) }),
      msg(T0 + 3 * MIN, { replyTo: String(T0 + MIN) }),
    ];
    expect(formatBody('channel', posts, 'summary')).toEqual({
      message_count: 4,
      newest_time: posts[3].time,
      threads: [
        { root_id: String(T0 - MIN), root_missing: true, reply_count: 1, newest_time: posts[0].time },
        { root_id: String(T0 + MIN), subject: 'Release plan', reply_count: 2, newest_time: posts[3].time },
      ],
    });
  });

  describe('with roots fetched from outside the window', () => {
    const outside = msg(T0 - MIN, { subject: 'Old plan', text: 'old root' });
    const roots = new Map([[outside.id, outside]]);
    const posts = [msg(T0, { replyTo: outside.id }), msg(T0 + MIN, { replyTo: outside.id })];

    it('puts the fetched root on its thread, marked outside_window, in full', () => {
      expect(formatBody('channel', posts, 'full', roots)).toEqual({
        threads: [{ root: { ...outside, outside_window: true }, replies: posts }],
      });
    });

    it('keeps outside_window on the compact root', () => {
      expect(formatBody('channel', posts, 'compact', roots)).toEqual({
        threads: [
          {
            root: { id: outside.id, time: outside.time, sender: 'Example Person', text: 'old root', subject: 'Old plan', outside_window: true },
            replies: posts.map((p) => ({ id: p.id, time: p.time, sender: 'Example Person', text: p.text })),
          },
        ],
      });
    });

    it('gives the summary the root\'s subject and root_outside_window, timed by the replies', () => {
      expect(formatBody('channel', posts, 'summary', roots)).toEqual({
        message_count: 2,
        newest_time: posts[1].time,
        threads: [
          { root_id: outside.id, root_outside_window: true, subject: 'Old plan', reply_count: 2, newest_time: posts[1].time },
        ],
      });
    });

    it('ignores a fetched root for a chat, and one the window already holds', () => {
      const inWindow = msg(T0 - MIN, { text: 'in-window copy' });
      expect(formatBody('channel', [inWindow, ...posts], 'full', roots)).toEqual({
        threads: [{ root: inWindow, replies: posts }],
      });
      expect(formatBody('group', posts, 'full', roots)).toEqual({ messages: posts });
    });
  });
});
