import { describe, expect, it, vi } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import { knownPeople, normaliseName, personMri, resolvePeople, type KnownPerson } from '../../src/tools/people.js';
import { THIRD_PARTY_NAMES_NOTE } from '../../src/tools/untrusted.js';
import { ME, fakeChatsvc, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const ME_OID = ME.slice('8:orgid:'.length);
const GREG_OID = '00000000-0000-0000-0000-0000000000c3';
const GREG = `8:orgid:${GREG_OID}`;
const GROUP = '19:group0001@thread.v2';
const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;

function caught(fn: () => unknown): McpToolError {
  try {
    fn();
  } catch (e) {
    expect(e).toBeInstanceOf(McpToolError);
    return e as McpToolError;
  }
  throw new Error('expected a throw');
}

describe('personMri', () => {
  it('takes an MRI as is', () => {
    expect(personMri(` ${GREG} `)).toBe(GREG);
  });

  it('makes a bare GUID an orgid MRI', () => {
    expect(personMri(GREG_OID)).toBe(GREG);
    expect(personMri(GREG_OID.toUpperCase())).toBe(GREG);
  });

  it('is undefined for a name', () => {
    expect(personMri('Greg King')).toBeUndefined();
  });
});

describe('normaliseName', () => {
  it('treats "Last, First" and "First Last" as the same name, ignoring case and spacing', () => {
    expect(normaliseName('King, Greg')).toBe('greg king');
    expect(normaliseName('  Greg   KING ')).toBe('greg king');
  });
});

describe('knownPeople', () => {
  it("names each 1:1's other person (as the 1:1 is titled) and the chat's recent senders", async () => {
    const oneOnOne = `19:${ME_OID}_${GREG_OID}@unq.gbl.spaces`;
    const chatsvc = fakeChatsvc([{ ...rawConv(oneOnOne, { at: T0 }), lastMessage: { ...rawConv(oneOnOne).lastMessage, from: FROM(GREG), imdisplayname: 'King, Greg' } }], {
      getMessages: vi.fn(async (id: string) =>
        id === GROUP ? [{ ...rawMsg(T0, { sender: 'Bea Example' }), from: FROM('8:orgid:bea') }] : [],
      ),
    });
    const known = await knownPeople(chatsvc, ME, GROUP);
    expect(known).toEqual(
      expect.arrayContaining([
        { mri: GREG, name: 'King, Greg' },
        { mri: '8:orgid:bea', name: 'Bea Example' },
      ]),
    );
    expect(known).toHaveLength(2);
  });

  it('still returns what it has when the recent-sender read fails', async () => {
    const chatsvc = fakeChatsvc([], { getMessages: vi.fn().mockRejectedValue(new Error('boom')) });
    await expect(knownPeople(chatsvc, ME, GROUP)).resolves.toEqual([]);
  });
});

describe('resolvePeople', () => {
  const KNOWN: KnownPerson[] = [
    { mri: GREG, name: 'King, Greg' },
    { mri: '8:orgid:greta', name: 'Greta Kingsley' },
    { mri: '8:orgid:bea', name: 'Bea Example' },
  ];

  it('resolves an MRI, a bare GUID and names in either order, putting known names on them', () => {
    expect(resolvePeople([GREG, GREG_OID, 'greg king', 'King, Greg', 'bea'], KNOWN)).toEqual([
      { person: GREG, mri: GREG, name: 'King, Greg' },
      { person: GREG_OID, mri: GREG, name: 'King, Greg' },
      { person: 'greg king', mri: GREG, name: 'King, Greg' },
      { person: 'King, Greg', mri: GREG, name: 'King, Greg' },
      { person: 'bea', mri: '8:orgid:bea', name: 'Bea Example' },
    ]);
  });

  it('prefers an exact match over substrings', () => {
    expect(resolvePeople(['Greg King'], KNOWN)[0]?.mri).toBe(GREG);
  });

  it('refuses an ambiguous name, listing the candidates JSON-quoted with their MRIs', () => {
    const e = caught(() => resolvePeople(['king'], KNOWN));
    expect(e.message).toContain('"King, Greg" (8:orgid:00000000-0000-0000-0000-0000000000c3)');
    expect(e.message).toContain('"Greta Kingsley" (8:orgid:greta)');
    expect(e.message).toContain(THIRD_PARTY_NAMES_NOTE);
    expect(e.hint).toMatch(/MRI/);
  });

  it('refuses an unknown name, asking for the MRI and saying only people from your chats are found', () => {
    const e = caught(() => resolvePeople(['Zed Nobody'], KNOWN));
    expect(e.message).toMatch(/Zed Nobody/);
    expect(e.hint).toMatch(/8:orgid:/);
    expect(e.hint).toMatch(/only people from your chats/i);
  });

  it('leaves an MRI no one is known by without a name', () => {
    expect(resolvePeople(['8:orgid:stranger'], KNOWN)).toEqual([{ person: '8:orgid:stranger', mri: '8:orgid:stranger' }]);
  });
});
