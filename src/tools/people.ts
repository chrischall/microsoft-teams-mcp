/**
 * People by name, for the tools that act on someone: a group chat's members
 * (`teams_add_chat_members`, `teams_remove_chat_members`,
 * `teams_create_group_chat`) and a message's
 * @mentions (`teams_send_message`).
 *
 * A person is, in this order:
 *   - an MRI (`8:orgid:<oid>`) or a bare oid — taken as is (source `id`);
 *   - an email address (has an `@`) — looked up in the directory (Substrate's
 *     people suggestions, `substrate.ts`), and taken only when exactly one
 *     person there has that email or UPN, ignoring case;
 *   - a display name — matched first among the people this MCP can already
 *     see (source `chat`): the other person of each 1:1 chat (named as
 *     `tools/titles.ts` names that 1:1) and the recent senders of the chat
 *     being changed; exact (as {@link normaliseName} compares), else
 *     substring; several is refused. ONLY when that finds no one is the
 *     directory asked (source `directory`), and a suggestion is taken only
 *     when it is (a) the one suggestion whose normalised name equals the
 *     query's, or (b) the ONLY suggestion and the normalised query is a run of
 *     whole words of its normalised name ("solomon" or "gary" for "Solomon,
 *     Gary", never "sol"). Anything else is refused, listing the suggestions.
 *
 * A name is never guessed: ambiguous or unknown is an error. The chat
 * service's own user-search endpoints answer 401 with the chat-service token
 * (live, 2026-11-09), so the directory is reached through Substrate with the
 * token search already captures.
 */
import { McpToolError } from '@chrischall/mcp-utils';
import type { ChatServiceClient } from '../chatsvc.js';
import type { DirectoryPerson, SubstrateSearch } from '../substrate.js';
import { isUserMessage, oneOnOnePeerMri, senderMri, summariseAll } from '../model.js';
import { enrichTitles } from './titles.js';
import { THIRD_PARTY_NAMES_NOTE } from './untrusted.js';

/** An MRI such as `8:orgid:<oid>` (a person) or `28:<id>` (a bot). */
export const MRI = /^\d+:\S+$/;

/** A bare Entra object id. */
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Recent messages scanned for senders. */
const RECENT_MESSAGES = 200;

/** Candidates an error lists. */
const MAX_CANDIDATES = 10;

export interface KnownPerson {
  mri: string;
  name: string;
}

/** Where a person was found: taken as the MRI / object id given, from the user's chats, or from the directory. */
export type PersonSource = 'id' | 'chat' | 'directory';

export interface ResolvedPerson {
  /** What the caller passed. */
  person: string;
  mri: string;
  /** A known display name, when there is one. */
  name?: string;
  source: PersonSource;
  /** A directory person's email (or UPN), so the user can see exactly who it is. */
  email?: string;
}

/** What resolution needs of the directory. */
export type PeopleDirectory = Pick<SubstrateSearch, 'suggestPeople'>;

/** The MRI an input names directly — an MRI as is, a bare GUID as `8:orgid:<guid>` — else `undefined` (a name). */
export function personMri(input: string): string | undefined {
  const s = input.trim();
  if (MRI.test(s)) return s;
  if (GUID.test(s)) return `8:orgid:${s.toLowerCase()}`;
  return undefined;
}

/**
 * A name for comparing: lower-cased, spacing collapsed, and a "Last, First"
 * display name (common in Teams) turned round to "first last".
 */
export function normaliseName(name: string): string {
  const s = name.trim().toLowerCase().replace(/\s+/g, ' ');
  const m = /^([^,]+),\s*([^,]+)$/.exec(s);
  return m ? `${m[2].trim()} ${m[1].trim()}` : s;
}

/** Who has written in the conversation lately, by MRI → display name. */
export async function recentSenders(chatsvc: ChatServiceClient, conversationId: string): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  for (const m of await chatsvc.getMessages(conversationId, { max: RECENT_MESSAGES })) {
    const mri = senderMri(m.from);
    const name = m.imdisplayname?.trim();
    if (isUserMessage(m) && mri !== undefined && name && !out.has(mri)) out.set(mri, name);
  }
  return out;
}

/**
 * Everyone this MCP can put a name to for `conversationId`: the other person
 * of every 1:1 the user has (as that 1:1 is titled), then the chat's recent
 * senders. With no `conversationId` (a chat not yet made), the 1:1 people
 * only. One entry per MRI (ignoring case), the 1:1 name winning. Never
 * throws — a failed read just knows fewer people.
 */
export async function knownPeople(chatsvc: ChatServiceClient, me: string, conversationId?: string): Promise<KnownPerson[]> {
  const byMri = new Map<string, KnownPerson>();
  const add = (mri: string, name: string) => {
    const key = mri.toLowerCase();
    if (name.trim() && !byMri.has(key)) byMri.set(key, { mri, name: name.trim() });
  };
  try {
    const raw = await chatsvc.listConversations();
    const oneOnOnes = summariseAll(raw, me).filter((s) => s.kind === 'oneOnOne');
    for (const s of await enrichTitles(chatsvc, raw, me, oneOnOnes)) {
      const peer = oneOnOnePeerMri(s.id, me);
      const name = s.participants?.[0];
      if (peer !== undefined && name !== undefined) add(peer, name);
    }
  } catch {
    // No 1:1 names; the chat's own senders may still do.
  }
  if (conversationId === undefined) return [...byMri.values()];
  try {
    for (const [mri, name] of await recentSenders(chatsvc, conversationId)) add(mri, name);
  } catch {
    // Fewer names, not a failure.
  }
  return [...byMri.values()];
}

/** The known name of `mri`, if any. */
export function knownName(mri: string, known: readonly KnownPerson[]): string | undefined {
  const key = mri.toLowerCase();
  return known.find((k) => k.mri.toLowerCase() === key)?.name;
}

/** An email address, as opposed to a name. */
function isEmail(input: string): boolean {
  return /^[^@\s]+@[^@\s]+$/.test(input.trim());
}

/** Candidates, JSON-quoted (they are other people's display names), with email when known and MRI. */
function candidateList(people: readonly { name: string; mri: string; email?: string; upn?: string }[]): string {
  return people
    .slice(0, MAX_CANDIDATES)
    .map((k) => {
      const email = k.email ?? k.upn;
      return `${JSON.stringify(k.name)}${email !== undefined ? ` <${email}>` : ''} (${k.mri})`;
    })
    .join('; ');
}

/** The local match for a name: one person, none (`undefined`), or a refusal when several match. */
function matchKnown(person: string, known: readonly KnownPerson[]): KnownPerson | undefined {
  const q = normaliseName(person);
  const raw = person.trim().toLowerCase();
  const exact = known.filter((k) => normaliseName(k.name) === q);
  const found =
    exact.length > 0
      ? exact
      : known.filter((k) => normaliseName(k.name).includes(q) || k.name.toLowerCase().includes(raw));
  if (found.length === 1) return found[0];
  if (found.length > 1) {
    throw new McpToolError(
      `"${person}" matches ${found.length} people: ${candidateList(found)}. ${THIRD_PARTY_NAMES_NOTE}`,
      { hint: 'Pass the MRI (8:orgid:…) of the one you mean.' },
    );
  }
  return undefined;
}

const PASS_ID = "Pass the person's email address or MRI (8:orgid:…) instead.";

/** A directory failure as the refusal for `person`, keeping the failure's message, kind and remedy. */
function directoryFailed(person: string, err: unknown, email: boolean): McpToolError {
  const message = err instanceof Error ? err.message : String(err);
  const own = err instanceof McpToolError ? err.hint : undefined;
  const remedy =
    own ??
    'Run any search in the Teams search box of a signed-in Teams tab while retrying (that is what lets this ' +
      'server capture the token it searches the directory with).';
  const lead = email ? `Could not look up "${person}" because` : `No one in your chats is called "${person}", and`;
  return new McpToolError(`${lead} the directory could not be searched: ${message}`, {
    hint: `${remedy} Or pass the person's MRI (8:orgid:…) instead.`,
    ...(err instanceof McpToolError && err.kind !== undefined ? { kind: err.kind } : {}),
    cause: err,
  });
}

function fromDirectory(person: string, p: DirectoryPerson): ResolvedPerson {
  const email = p.email ?? p.upn;
  return { person, mri: p.mri, name: p.name, source: 'directory', ...(email !== undefined ? { email } : {}) };
}

/** The one directory person an email names: their email or UPN equal to it, ignoring case. */
function pickByEmail(person: string, found: readonly DirectoryPerson[]): ResolvedPerson {
  const want = person.trim().toLowerCase();
  const hits = found.filter((p) => p.email?.toLowerCase() === want || p.upn?.toLowerCase() === want);
  if (hits.length === 1) return fromDirectory(person, hits[0]);
  if (hits.length > 1) {
    throw new McpToolError(
      `"${person}" is the address of ${hits.length} people in the directory: ${candidateList(hits)}. ${THIRD_PARTY_NAMES_NOTE}`,
      { hint: 'Pass the MRI (8:orgid:…) of the one you mean.' },
    );
  }
  throw new McpToolError(`No one in your organisation's directory has the address "${person}".`, {
    hint:
      "Check the address, or pass the person's name or MRI (8:orgid:…) instead. Only people in your " +
      'organisation (not distribution lists or personal contacts) are found there.',
  });
}

/**
 * The one directory person a name means — see the module comment for the
 * rule: a unique exact normalised match, or the only suggestion when the
 * query is a run of whole words of its name. Otherwise refused.
 */
function pickByName(person: string, found: readonly DirectoryPerson[]): ResolvedPerson {
  const q = normaliseName(person);
  const exact = found.filter((p) => normaliseName(p.name) === q);
  if (exact.length === 1) return fromDirectory(person, exact[0]);
  if (exact.length > 1) {
    throw new McpToolError(
      `"${person}" matches ${exact.length} people in the directory: ${candidateList(exact)}. ${THIRD_PARTY_NAMES_NOTE}`,
      { hint: 'Pass the email address or MRI (8:orgid:…) of the one you mean.' },
    );
  }
  if (found.length === 1 && ` ${normaliseName(found[0].name)} `.includes(` ${q} `)) return fromDirectory(person, found[0]);
  if (found.length > 0) {
    throw new McpToolError(
      `No one in your chats is called "${person}", and no one in the directory is called exactly that; it ` +
        `suggested: ${candidateList(found)}. ${THIRD_PARTY_NAMES_NOTE}`,
      { hint: 'If one of them is who you mean, pass their email address or MRI (8:orgid:…).' },
    );
  }
  throw new McpToolError(`No one in your chats or your organisation's directory is called "${person}".`, {
    hint:
      `${PASS_ID} A name is matched among the other person of each of your 1:1 chats and (for an existing chat) ` +
      "its recent senders, then searched in your organisation's directory.",
  });
}

/**
 * Each requested person's MRI (and name, source and — from the directory —
 * email), by the rules in the module comment. Every person resolves or the
 * call throws, naming the candidates or the failure. Local matching runs for
 * every input first, so a local ambiguity refuses before the directory is
 * asked; then the directory is asked once per distinct input that needs it
 * (`force`, so a retry after running a Teams search really re-captures).
 */
export async function resolvePeople(
  inputs: readonly string[],
  known: readonly KnownPerson[],
  directory: PeopleDirectory,
): Promise<ResolvedPerson[]> {
  type Pending = { person: string; email: boolean };
  const local: (ResolvedPerson | Pending)[] = inputs.map((person) => {
    const direct = personMri(person);
    if (direct !== undefined) {
      const name = knownName(direct, known);
      return { person, mri: direct, ...(name !== undefined ? { name } : {}), source: 'id' as const };
    }
    if (isEmail(person)) return { person, email: true };
    const k = matchKnown(person, known);
    return k !== undefined ? { person, mri: k.mri, name: k.name, source: 'chat' as const } : { person, email: false };
  });

  const asked = new Map<string, DirectoryPerson[]>();
  const out: ResolvedPerson[] = [];
  for (const r of local) {
    if ('mri' in r) {
      out.push(r);
      continue;
    }
    const text = r.person.trim();
    const key = r.email ? `@${text.toLowerCase()}` : normaliseName(text);
    let found = asked.get(key);
    if (found === undefined) {
      try {
        found = await directory.suggestPeople(text, { force: true });
      } catch (err) {
        throw directoryFailed(r.person, err, r.email);
      }
      asked.set(key, found);
    }
    out.push(r.email ? pickByEmail(r.person, found) : pickByName(r.person, found));
  }
  return out;
}
