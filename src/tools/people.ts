/**
 * People by name, for the tools that act on someone: a group chat's members
 * (`teams_add_chat_members`, `teams_remove_chat_members`) and a message's
 * @mentions (`teams_send_message`).
 *
 * There is no directory lookup: the Teams user-search endpoints answer 401
 * with the chat-service token (live, 2026-11-09). So a person is an MRI
 * (`8:orgid:<oid>`), a bare oid, or a display name matched among the people
 * this MCP can already see — the other person of each 1:1 chat (named as
 * `tools/titles.ts` names that 1:1) and the recent senders of the chat being
 * changed. A name is never guessed: ambiguous or unknown is an error.
 */
import { McpToolError } from '@chrischall/mcp-utils';
import type { ChatServiceClient } from '../chatsvc.js';
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

export interface ResolvedPerson {
  /** What the caller passed. */
  person: string;
  mri: string;
  /** A known display name, when there is one. */
  name?: string;
}

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
 * senders. One entry per MRI (ignoring case), the 1:1 name winning. Never
 * throws — a failed read just knows fewer people.
 */
export async function knownPeople(chatsvc: ChatServiceClient, me: string, conversationId: string): Promise<KnownPerson[]> {
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

/**
 * Each requested person's MRI (and known name). An MRI or bare GUID is taken
 * as is; a name must match exactly one known person — exactly (as
 * {@link normaliseName} compares), else as a substring. Ambiguous or unknown
 * throws, naming the candidates: every person resolves or nothing does.
 */
export function resolvePeople(inputs: readonly string[], known: readonly KnownPerson[]): ResolvedPerson[] {
  return inputs.map((person) => {
    const direct = personMri(person);
    if (direct !== undefined) {
      const name = knownName(direct, known);
      return { person, mri: direct, ...(name !== undefined ? { name } : {}) };
    }
    const q = normaliseName(person);
    const raw = person.trim().toLowerCase();
    const exact = known.filter((k) => normaliseName(k.name) === q);
    const found =
      exact.length > 0
        ? exact
        : known.filter((k) => normaliseName(k.name).includes(q) || k.name.toLowerCase().includes(raw));
    if (found.length === 1) return { person, mri: found[0].mri, name: found[0].name };
    // Candidate names are other people's display names: JSON-quoted and flagged as data.
    if (found.length > 1) {
      const named = found
        .slice(0, MAX_CANDIDATES)
        .map((k) => `${JSON.stringify(k.name)} (${k.mri})`)
        .join('; ');
      throw new McpToolError(`"${person}" matches ${found.length} people: ${named}. ${THIRD_PARTY_NAMES_NOTE}`, {
        hint: 'Pass the MRI (8:orgid:…) of the one you mean.',
      });
    }
    throw new McpToolError(`No one in your chats is called "${person}".`, {
      hint:
        "Pass the person's MRI (8:orgid:…) instead. Only people from your chats can be found by name — " +
        'the other person of one of your 1:1 chats, or someone who wrote recently in this chat.',
    });
  });
}
