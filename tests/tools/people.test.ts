import { describe, expect, it, vi } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import { knownPeople, normaliseName, personMri, resolvePeople, type KnownPerson } from '../../src/tools/people.js';
import type { DirectoryPerson } from '../../src/substrate.js';
import { THIRD_PARTY_NAMES_NOTE } from '../../src/tools/untrusted.js';
import { ME, fakeChatsvc, rawConv, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const ME_OID = ME.slice('8:orgid:'.length);
const GREG_OID = '00000000-0000-0000-0000-0000000000c3';
const GREG = `8:orgid:${GREG_OID}`;
const GROUP = '19:group0001@thread.v2';
const FROM = (mri: string) => `https://teams.cloud.microsoft/api/chatsvc/amer/v1/users/ME/contacts/${mri}`;

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

/** A directory that knows `people`, matching any whose name, email or UPN contains a query word. */
function fakeDirectory(people: DirectoryPerson[] = [], fail?: unknown) {
  const suggestPeople = vi.fn(async (text: string, _opts?: { size?: number; force?: boolean }): Promise<DirectoryPerson[]> => {
    if (fail !== undefined) throw fail;
    const words = text.toLowerCase().replace(/,/g, ' ').split(/\s+/).filter(Boolean);
    return people.filter((p) =>
      words.some((w) => [p.name, p.email ?? '', p.upn ?? ''].some((f) => f.toLowerCase().includes(w))),
    );
  });
  return { suggestPeople };
}

const GARY: DirectoryPerson = {
  mri: '8:orgid:00000000-0000-0000-0000-0000000000b1',
  name: 'Solomon, Gary',
  email: 'gary.solomon@example.test',
  upn: 'gsolomon@example.test',
};
const GARY2: DirectoryPerson = { mri: '8:orgid:00000000-0000-0000-0000-0000000000a2', name: 'Gary Solomon', email: 'gary.solomon2@example.test' };
const GARETH: DirectoryPerson = { mri: '8:orgid:00000000-0000-0000-0000-0000000000a3', name: 'Gareth Solomons', email: 'gareth@example.test' };

async function rejected(p: Promise<unknown>): Promise<McpToolError> {
  const e = await p.then(
    () => undefined,
    (err: unknown) => err,
  );
  expect(e).toBeInstanceOf(McpToolError);
  return e as McpToolError;
}

describe('resolvePeople', () => {
  const KNOWN: KnownPerson[] = [
    { mri: GREG, name: 'King, Greg' },
    { mri: '8:orgid:greta', name: 'Greta Kingsley' },
    { mri: '8:orgid:bea', name: 'Bea Example' },
  ];

  it('resolves an MRI, a bare GUID and names in either order, putting known names on them', async () => {
    const dir = fakeDirectory();
    expect(await resolvePeople([GREG, GREG_OID, 'greg king', 'King, Greg', 'bea'], KNOWN, dir)).toEqual([
      { person: GREG, mri: GREG, name: 'King, Greg', source: 'id' },
      { person: GREG_OID, mri: GREG, name: 'King, Greg', source: 'id' },
      { person: 'greg king', mri: GREG, name: 'King, Greg', source: 'chat' },
      { person: 'King, Greg', mri: GREG, name: 'King, Greg', source: 'chat' },
      { person: 'bea', mri: '8:orgid:bea', name: 'Bea Example', source: 'chat' },
    ]);
    expect(dir.suggestPeople).not.toHaveBeenCalled();
  });

  it('prefers an exact match over substrings', async () => {
    expect((await resolvePeople(['Greg King'], KNOWN, fakeDirectory()))[0]?.mri).toBe(GREG);
  });

  it('refuses an ambiguous name, listing the candidates JSON-quoted with their MRIs, without asking the directory', async () => {
    const dir = fakeDirectory([GARY]);
    const e = await rejected(resolvePeople(['Gary Solomon', 'king'], KNOWN, dir));
    expect(e.message).toContain('"King, Greg" (8:orgid:00000000-0000-0000-0000-0000000000c3)');
    expect(e.message).toContain('"Greta Kingsley" (8:orgid:greta)');
    expect(e.message).toContain(THIRD_PARTY_NAMES_NOTE);
    expect(e.hint).toMatch(/MRI/);
    expect(dir.suggestPeople).not.toHaveBeenCalled();
  });

  it('leaves an MRI no one is known by without a name', async () => {
    expect(await resolvePeople(['8:orgid:stranger'], KNOWN, fakeDirectory())).toEqual([
      { person: '8:orgid:stranger', mri: '8:orgid:stranger', source: 'id' },
    ]);
  });

  describe('directory fallback for a name', () => {
    it('only on a local miss: an exact normalised name in the directory wins, with its email', async () => {
      const dir = fakeDirectory([GARY, GARETH]);
      expect(await resolvePeople(['Gary Solomon'], KNOWN, dir)).toEqual([
        { person: 'Gary Solomon', mri: GARY.mri, name: 'Solomon, Gary', source: 'directory', email: GARY.email },
      ]);
      expect(dir.suggestPeople).toHaveBeenCalledTimes(1);
      expect(dir.suggestPeople).toHaveBeenCalledWith('Gary Solomon', { force: true });
    });

    it('falls back to the UPN as the email shown when the directory has no address', async () => {
      const dir = fakeDirectory([{ mri: GARY.mri, name: GARY.name, upn: GARY.upn }]);
      expect((await resolvePeople(['Gary Solomon'], KNOWN, dir))[0]).toMatchObject({ email: GARY.upn });
    });

    it('refuses several exact matches, listing name, email and MRI', async () => {
      const e = await rejected(resolvePeople(['gary solomon'], KNOWN, fakeDirectory([GARY, GARY2])));
      expect(e.message).toContain(`"Solomon, Gary" <${GARY.email}> (${GARY.mri})`);
      expect(e.message).toContain(`"Gary Solomon" <${GARY2.email}> (${GARY2.mri})`);
      expect(e.message).toContain(THIRD_PARTY_NAMES_NOTE);
      expect(e.hint).toMatch(/email|MRI/);
    });

    it('takes a lone suggestion the name is a run of whole words of', async () => {
      const r = await resolvePeople(['solomon'], KNOWN, fakeDirectory([GARY]));
      expect(r[0]).toMatchObject({ mri: GARY.mri, source: 'directory' });
    });

    it('refuses a lone suggestion that is only a partial-word match, naming it', async () => {
      const e = await rejected(resolvePeople(['solo'], KNOWN, fakeDirectory([GARY])));
      expect(e.message).toContain(`"Solomon, Gary" <${GARY.email}> (${GARY.mri})`);
      expect(e.hint).toMatch(/email|MRI/);
    });

    it('refuses several suggestions none of which is an exact match, listing them', async () => {
      const e = await rejected(resolvePeople(['solomon'], KNOWN, fakeDirectory([GARY, GARETH])));
      expect(e.message).toContain('"Solomon, Gary"');
      expect(e.message).toContain('"Gareth Solomons"');
    });

    it('refuses a name found nowhere, saying the directory was searched too', async () => {
      const e = await rejected(resolvePeople(['Zed Nobody'], KNOWN, fakeDirectory([GARY])));
      expect(e.message).toMatch(/Zed Nobody/);
      expect(e.message).toMatch(/directory/i);
      expect(e.hint).toMatch(/8:orgid:/);
      expect(e.hint).toMatch(/directory/i);
    });

    it('asks the directory once per distinct name', async () => {
      const dir = fakeDirectory([GARY]);
      const r = await resolvePeople(['Gary Solomon', 'gary  solomon', 'Solomon, Gary'], KNOWN, dir);
      expect(r.map((p) => p.mri)).toEqual([GARY.mri, GARY.mri, GARY.mri]);
      expect(dir.suggestPeople).toHaveBeenCalledTimes(1);
    });

    it('says the directory could not be searched, keeping the capture failure and its remedy', async () => {
      const dir = fakeDirectory([], new Error('could not capture a Substrate search token from the browser (timed out)'));
      const e = await rejected(resolvePeople(['Zed Nobody'], KNOWN, dir));
      expect(e.message).toMatch(/Zed Nobody/);
      expect(e.message).toMatch(/directory could not be searched/i);
      expect(e.message).toMatch(/timed out/);
      expect(e.hint).toMatch(/search box/i);
      expect(e.hint).toMatch(/8:orgid:/);
    });

    it("keeps a failure's own hint (a disabled bridge)", async () => {
      const off = new McpToolError('Substrate search is disabled (no token source).', {
        kind: 'no_credential',
        hint: 'Unset TEAMS_DISABLE_FETCHPROXY so a Substrate token can be captured from a signed-in Teams tab.',
      });
      const e = await rejected(resolvePeople(['Zed Nobody'], KNOWN, fakeDirectory([], off)));
      expect(e.message).toMatch(/disabled/);
      expect(e.hint).toMatch(/TEAMS_DISABLE_FETCHPROXY/);
      expect(e.hint).toMatch(/8:orgid:/);
      expect(e.kind).toBe('no_credential');
    });

    it('stops asking after the first failure', async () => {
      const dir = fakeDirectory([], new Error('timed out'));
      await rejected(resolvePeople(['Zed Nobody', 'Yan Other'], KNOWN, dir));
      expect(dir.suggestPeople).toHaveBeenCalledTimes(1);
    });
  });

  describe('an email address', () => {
    it('is looked up in the directory and matched on its email, ignoring case', async () => {
      const dir = fakeDirectory([GARY, GARY2]);
      expect(await resolvePeople(['Gary.Solomon@Example.test'], KNOWN, dir)).toEqual([
        { person: 'Gary.Solomon@Example.test', mri: GARY.mri, name: 'Solomon, Gary', source: 'directory', email: GARY.email },
      ]);
      expect(dir.suggestPeople).toHaveBeenCalledWith('Gary.Solomon@Example.test', { force: true });
    });

    it('is matched on the UPN too', async () => {
      expect((await resolvePeople(['GSOLOMON@example.test'], KNOWN, fakeDirectory([GARY])))[0]?.mri).toBe(GARY.mri);
    });

    it('is refused when no one has exactly that address, even with suggestions', async () => {
      const e = await rejected(resolvePeople(['gary@example.test'], KNOWN, fakeDirectory([GARY])));
      expect(e.message).toMatch(/gary@example\.test/);
      expect(e.message).toMatch(/directory/i);
      expect(e.hint).toMatch(/8:orgid:/);
    });

    it('surfaces a directory failure', async () => {
      const e = await rejected(resolvePeople(['gary@example.test'], KNOWN, fakeDirectory([], new Error('pairing required'))));
      expect(e.message).toMatch(/pairing required/);
      expect(e.hint).toMatch(/8:orgid:/);
    });
  });
});
