import { describe, it, expect, vi } from 'vitest';
import type { ChatServiceClient, GetMessagesOptions, RawConversation, RawMessage } from '../../src/chatsvc.js';
import { summariseAll } from '../../src/model.js';
import { enrichTitles, groupTitleFrom, isUntitledGroup, TITLE_FETCH_MESSAGES, TITLE_PAGE_SIZE } from '../../src/tools/titles.js';
import { fakeChatsvc, ME, messagesFrom, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;

/** A message from a named person (MRI derived from the name), arriving at `at`. */
function from(name: string, at: number, opts: { type?: string } = {}): RawMessage {
  const mri = `8:orgid:${name.toLowerCase().replace(/[^a-z]/g, '')}`;
  return { ...rawMsg(at, { sender: name, ...(opts.type ? { type: opts.type } : {}) }), from: FROM(mri) };
}

const mine = (at: number): RawMessage => rawMsg(at, { fromMe: true });

/** 15 untitled group chats whose last message is from the same person, as seen live. */
const GROUPS = Array.from({ length: 15 }, (_, i) => `19:group${String(i).padStart(4, '0')}@thread.v2`);
const SAME_LAST_SENDER = GROUPS.map((id, i) => rawConv(id, { at: T0 + i * MIN, sender: 'King, Greg' }));
/** Each chat's other members: Greg, plus one invented person unique to that chat. */
const MEMBERS: Record<string, RawMessage[]> = Object.fromEntries(
  GROUPS.map((id, i) => [id, [from(`Person${String.fromCharCode(65 + i)}, Example`, T0 - 2 * MIN), mine(T0 - MIN), from('King, Greg', T0 + i * MIN)]]),
);

function perConversation(byId: Record<string, RawMessage[] | Error>) {
  return vi.fn(async (id: string, opts: GetMessagesOptions) => {
    const v = byId[id] ?? [];
    if (v instanceof Error) throw v;
    return messagesFrom(v)(id, opts);
  });
}

function chatsvcWith(convs: RawConversation[], byId: Record<string, RawMessage[] | Error>): ChatServiceClient {
  return fakeChatsvc(convs, { getMessages: perConversation(byId) });
}

describe('groupTitleFrom', () => {
  it('names up to three distinct other senders, most recent first, and counts the rest as +N', () => {
    const msgs = [
      from('Ann', T0 + 6),
      from('Bob', T0 + 5),
      from('Ann', T0 + 4),
      mine(T0 + 3),
      from('Cy', T0 + 2),
      from('Di', T0 + 1),
      from('Ed', T0),
    ];
    expect(groupTitleFrom(msgs, ME)).toEqual({
      title: 'Group chat: Ann; Bob; Cy +2',
      participants: ['Ann', 'Bob', 'Cy', 'Di', 'Ed'],
    });
  });

  it('counts two senders with the same display name once, even under different MRIs', () => {
    const twin = { ...from('Ann', T0 + 2), from: FROM('8:orgid:ann-other-account') };
    const msgs = [from('Ann', T0 + 3), twin, from('Bob', T0 + 1)];
    expect(groupTitleFrom(msgs, ME)).toEqual({ title: 'Group chat: Ann; Bob', participants: ['Ann', 'Bob'] });
  });

  it('leaves the user out, and system events (whose name is whoever triggered them)', () => {
    const msgs = [mine(T0 + 3), from('Sys Trigger', T0 + 2, { type: 'Event/Call' }), from('Ann', T0 + 1)];
    expect(groupTitleFrom(msgs, ME)).toEqual({ title: 'Group chat: Ann', participants: ['Ann'] });
  });

  it('is undefined when nobody but the user wrote', () => {
    expect(groupTitleFrom([mine(T0 + 1), mine(T0)], ME)).toBeUndefined();
  });
});

describe('isUntitledGroup', () => {
  it('is true only for a group chat with no topic', () => {
    expect(isUntitledGroup(rawConv('19:g@thread.v2'))).toBe(true);
    expect(isUntitledGroup(rawConv('19:g@thread.v2', { topic: 'Falcon' }))).toBe(false);
    expect(isUntitledGroup(rawConv('19:g@thread.v2', { topic: '  ' }))).toBe(true);
    expect(isUntitledGroup(rawConv('19:aaaa_bbbb@unq.gbl.spaces'))).toBe(false);
    expect(isUntitledGroup(rawConv('19:meeting_x@thread.v2'))).toBe(false);
    expect(isUntitledGroup(rawConv('19:c@thread.tacv2'))).toBe(false);
  });
});

describe('enrichTitles', () => {
  it('gives 15 untitled chats with the same last sender 15 distinct participant titles', async () => {
    const chatsvc = chatsvcWith(SAME_LAST_SENDER, MEMBERS);
    const before = summariseAll(SAME_LAST_SENDER, ME).filter((s) => s.kind === 'group');
    expect(new Set(before.map((s) => s.title))).toEqual(new Set(['King, Greg']));

    const after = await enrichTitles(chatsvc, SAME_LAST_SENDER, ME, before);
    expect(new Set(after.map((s) => s.title)).size).toBe(15);
    expect(after[0]).toMatchObject({ title: 'Group chat: King, Greg; PersonA, Example', participants: ['King, Greg', 'PersonA, Example'] });
    expect(chatsvc.getMessages).toHaveBeenCalledTimes(15);
    expect(vi.mocked(chatsvc.getMessages).mock.calls[0][1]).toMatchObject({ max: TITLE_FETCH_MESSAGES, pageSize: TITLE_PAGE_SIZE });
  });

  it('reads at most four chats at once', async () => {
    let inFlight = 0;
    let peak = 0;
    const getMessages = vi.fn(async (id: string, opts: GetMessagesOptions) => {
      inFlight++;
      peak = Math.max(peak, inFlight);
      await new Promise((r) => setTimeout(r, 2));
      inFlight--;
      return messagesFrom(MEMBERS[id])(id, opts);
    });
    const chatsvc = fakeChatsvc(SAME_LAST_SENDER, { getMessages });
    await enrichTitles(chatsvc, SAME_LAST_SENDER, ME, summariseAll(SAME_LAST_SENDER, ME));
    expect(peak).toBe(4);
  });

  it('leaves titled chats, 1:1s, meetings and channels alone without reading them', async () => {
    const convs = [
      rawConv('19:titled@thread.v2', { topic: 'Falcon' }),
      rawConv('19:aaaa_bbbb@unq.gbl.spaces', { sender: 'Pat Example' }),
      rawConv('19:meeting_x@thread.v2', { topic: 'Sync' }),
      rawConv('19:meeting_y@thread.v2', { fromMe: true }),
      rawConv('19:c@thread.tacv2', { topic: 'Design' }),
    ];
    const chatsvc = chatsvcWith(convs, {});
    const before = summariseAll(convs, ME);
    expect(await enrichTitles(chatsvc, convs, ME, before)).toEqual(before);
    expect(chatsvc.getMessages).not.toHaveBeenCalled();
  });

  it('caches by conversation and last message, so a repeat call reads nothing', async () => {
    const chatsvc = chatsvcWith(SAME_LAST_SENDER, MEMBERS);
    const rows = summariseAll(SAME_LAST_SENDER, ME);
    const first = await enrichTitles(chatsvc, SAME_LAST_SENDER, ME, rows);
    const second = await enrichTitles(chatsvc, SAME_LAST_SENDER, ME, rows);
    expect(second).toEqual(first);
    expect(chatsvc.getMessages).toHaveBeenCalledTimes(15);
  });

  it('reads a chat again once its last message changed', async () => {
    const conv = [rawConv(GROUPS[0], { at: T0 })];
    const chatsvc = chatsvcWith(conv, { [GROUPS[0]]: [from('Ann', T0)] });
    await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME));
    const moved = [rawConv(GROUPS[0], { at: T0 + MIN })];
    vi.mocked(chatsvc.getMessages).mockImplementation(perConversation({ [GROUPS[0]]: [from('Bob', T0 + MIN), from('Ann', T0)] }));
    const [row] = await enrichTitles(chatsvc, moved, ME, summariseAll(moved, ME));
    expect(row.title).toBe('Group chat: Bob; Ann');
    expect(chatsvc.getMessages).toHaveBeenCalledTimes(2);
  });

  it('does not cache a chat whose listing carries no last message id, so it can refresh', async () => {
    const { lastMessage, ...bare } = rawConv(GROUPS[0], { at: T0 });
    const { id: _id, ...lastWithoutId } = lastMessage ?? {};
    const conv: RawConversation[] = [{ ...bare, lastMessage: lastWithoutId }];
    const chatsvc = chatsvcWith(conv, { [GROUPS[0]]: [from('Ann', T0)] });
    await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME));
    vi.mocked(chatsvc.getMessages).mockImplementation(perConversation({ [GROUPS[0]]: [from('Bob', T0 + MIN), from('Ann', T0)] }));
    const [row] = await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME));
    expect(row.title).toBe('Group chat: Bob; Ann');
    expect(chatsvc.getMessages).toHaveBeenCalledTimes(2);

    const noLast: RawConversation[] = [bare];
    const chatsvc2 = chatsvcWith(noLast, { [GROUPS[0]]: [from('Ann', T0)] });
    await enrichTitles(chatsvc2, noLast, ME, summariseAll(noLast, ME));
    await enrichTitles(chatsvc2, noLast, ME, summariseAll(noLast, ME));
    expect(chatsvc2.getMessages).toHaveBeenCalledTimes(2);
  });

  it('keeps the old title when the read fails, never failing the call, and retries next time', async () => {
    const conv = [rawConv(GROUPS[0], { sender: 'King, Greg' })];
    const chatsvc = chatsvcWith(conv, { [GROUPS[0]]: new Error('boom') });
    const [row] = await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME));
    expect(row.title).toBe('King, Greg');
    expect(row).not.toHaveProperty('participants');

    vi.mocked(chatsvc.getMessages).mockImplementation(perConversation({ [GROUPS[0]]: [from('Ann', T0)] }));
    const [again] = await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME));
    expect(again.title).toBe('Group chat: Ann');
  });

  it('falls back to "(untitled group chat)" when only the user wrote', async () => {
    const conv = [rawConv(GROUPS[0], { fromMe: true })];
    const chatsvc = chatsvcWith(conv, { [GROUPS[0]]: [mine(T0)] });
    const [row] = await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME));
    expect(row.title).toBe('(untitled group chat)');
    expect(row).not.toHaveProperty('participants');
  });
});

describe('title reads', () => {
  it('read one small page of 30: a cold cache of many chats is not 200 messages each', () => {
    expect(TITLE_FETCH_MESSAGES).toBe(30);
    expect(TITLE_PAGE_SIZE).toBe(30);
  });
});

describe('enrichTitles for 1:1 chats', () => {
  const MY_OID = ME.slice('8:orgid:'.length);
  const PEER_OID = '00000000-0000-0000-0000-0000000000c3';
  const PEER = `8:orgid:${PEER_OID}`;
  const DM = `19:${MY_OID}_${PEER_OID}@unq.gbl.spaces`;
  const peerSays = (at: number, name = 'Rivera, Sam'): RawMessage => ({ ...rawMsg(at, { sender: name }), from: FROM(PEER) });

  it('names a 1:1 whose last message is mine after the other person, from their recent messages', async () => {
    const conv = [rawConv(DM, { at: T0, fromMe: true })];
    const chatsvc = chatsvcWith(conv, { [DM]: [mine(T0), from('Bystander, Bo', T0 - 2 * MIN), peerSays(T0 - MIN)] });
    const before = summariseAll(conv, ME).filter((s) => s.id === DM);
    expect(before[0].title).toBe('(1:1 chat)');

    const [row] = await enrichTitles(chatsvc, conv, ME, before);
    expect(row).toMatchObject({ title: 'Rivera, Sam', participants: ['Rivera, Sam'] });
    expect(vi.mocked(chatsvc.getMessages).mock.calls[0][1]).toMatchObject({ max: TITLE_FETCH_MESSAGES, pageSize: TITLE_PAGE_SIZE });
  });

  it('caches it like a group chat\'s title', async () => {
    const conv = [rawConv(DM, { at: T0, fromMe: true })];
    const chatsvc = chatsvcWith(conv, { [DM]: [mine(T0), peerSays(T0 - MIN)] });
    await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME));
    await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME));
    expect(chatsvc.getMessages).toHaveBeenCalledTimes(1);
  });

  it('does not read a 1:1 whose last message already names the other person', async () => {
    const conv: RawConversation[] = [{ ...rawConv(DM, { at: T0 }), lastMessage: { ...rawConv(DM, { at: T0 }).lastMessage!, from: FROM(PEER), imdisplayname: 'Rivera, Sam' } }];
    const chatsvc = chatsvcWith(conv, {});
    const [row] = await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME).filter((s) => s.id === DM));
    expect(row).toMatchObject({ title: 'Rivera, Sam', participants: ['Rivera, Sam'] });
    expect(chatsvc.getMessages).not.toHaveBeenCalled();
  });

  it('stays "(1:1 chat)" when the other person wrote nothing recent', async () => {
    const conv = [rawConv(DM, { at: T0, fromMe: true })];
    const chatsvc = chatsvcWith(conv, { [DM]: [mine(T0), from('Bystander, Bo', T0 - MIN)] });
    const [row] = await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME).filter((s) => s.id === DM));
    expect(row.title).toBe('(1:1 chat)');
    expect(row).not.toHaveProperty('participants');
  });

  it('leaves a 1:1 with a topic alone', async () => {
    const conv = [rawConv(DM, { at: T0, fromMe: true, topic: 'Planning' })];
    const chatsvc = chatsvcWith(conv, {});
    const [row] = await enrichTitles(chatsvc, conv, ME, summariseAll(conv, ME).filter((s) => s.id === DM));
    expect(row.title).toBe('Planning');
    expect(chatsvc.getMessages).not.toHaveBeenCalled();
  });
});
