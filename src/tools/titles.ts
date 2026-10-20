/**
 * Names for conversations: untitled group chats and 1:1s titled from who is in
 * them, and {@link nameConversations} for any tool whose result carries a
 * conversation id.
 *
 * A group chat with no topic has nothing in the listing to call it by: the
 * listing's stand-in is the last message's sender (`conversationSummary`).
 * Live (2026-10-10, a real tenant) that made 15 different group chats all
 * "King, Greg" — one person who posts everywhere — so a list could not tell
 * them apart and name resolution was always ambiguous.
 *
 * This server has no member-list call: nothing in the recorded API notes
 * (`docs/plans/2026-10-10-chatsvc-api.md`) or the fixtures shows one (a
 * `threads/{id}` read, say) working with the captured chat-service token, so
 * none is guessed at. Every message does carry its sender's name and MRI, so
 * reading a chat's recent messages names the people in it — the members who
 * WROTE in that sample, not every member.
 *
 * The title is a pure function of that sampled set of people: they are sorted
 * by display name (lower-cased, compared by code unit — no locale rules — and
 * by MRI on a tie), the first three named and the rest counted as "+N". Who
 * posted most recently does not move it, so a chat keeps its name from one
 * message to the next; it changes only when the sample's people do (someone
 * new writes, or someone's last message ages out of the page read).
 *
 * An untitled 1:1 is named after the other person, whose MRI its id carries
 * (`oneOnOnePeerMri`). The listing has their name only when they sent the last
 * message; when the user did, the 1:1 was "(untitled chat)" and a query for
 * the other person's name could not find it (live, 2026-10-10: a favorite
 * 1:1). Their name is read from their recent messages here, the same way.
 *
 * Reads are bounded and cached: one page of {@link TITLE_FETCH_MESSAGES}
 * messages per chat (a cold cache of ~145 chats took 29s at the default
 * 200-message page), {@link TITLE_CONCURRENCY} chats at once, and the result is kept
 * in-process against the chat's last message id, so it is read again only
 * once something new arrives. Each tool reads only the chats it needs (the
 * rows it returns; every untitled chat only when matching a name), and
 * a failed read keeps the listing's title — it never fails the call.
 */
import { mapWithConcurrency } from '@chrischall/mcp-utils';
import type { ChatServiceClient, RawConversation, RawMessage } from '../chatsvc.js';
import {
  conversationKind,
  GROUP_TITLE_PREFIX,
  isFromMri,
  isUserMessage,
  oneOnOnePeerMri,
  senderMri,
  UNTITLED_GROUP_TITLE,
  UNTITLED_ONE_ON_ONE_TITLE,
  summariseAll,
  type ConversationKind,
  type ConversationSummary,
} from '../model.js';

/** Recent messages read per untitled chat. */
export const TITLE_FETCH_MESSAGES = 30;
/** Their page size, so the read is one small request rather than a 200-message page. */
export const TITLE_PAGE_SIZE = 30;
/** Untitled group chats read at once. */
export const TITLE_CONCURRENCY = 4;
/** Participants named in a title; the rest are counted as "+N". */
const NAMED_PARTICIPANTS = 3;
/**
 * Between names: Teams display names are often "Last, First", so a comma
 * would run two people together.
 */
const NAME_SEPARATOR = '; ';

export interface GroupTitle {
  title: string;
  /**
   * Every distinct sender other than the user seen in the messages read,
   * sorted by name (see {@link groupTitleFrom}) — the title names only the
   * first three.
   */
  participants: string[];
}

/** A group chat with no topic — the only kind whose title is a guess. */
export function isUntitledGroup(c: RawConversation): boolean {
  return conversationKind(c) === 'group' && !c.threadProperties?.topic?.trim();
}

/** Lower-cased names by code unit, then MRI: the same order on every machine, whatever its locale. */
export function byNameThenMri(a: { name: string; mri: string }, b: { name: string; mri: string }): number {
  const x = a.name.toLowerCase();
  const y = b.name.toLowerCase();
  if (x !== y) return x < y ? -1 : 1;
  return a.mri < b.mri ? -1 : a.mri > b.mri ? 1 : 0;
}

/**
 * The title for a group chat from its messages (newest first): its distinct
 * senders other than the user, sorted by name ({@link byNameThenMri}), the
 * first three named and the rest counted as "+N" — all of whom are its
 * `participants`, in that order. Who wrote most recently does not change it:
 * it depends only on who is in the messages read. A sender is distinct by MRI
 * (named by their newest message, so a rename shows) and by display name
 * ignoring case (of two accounts under one name, the lower MRI is kept), so
 * one person under two accounts appears once. System events are skipped —
 * their display name is whoever triggered them. `undefined` when no one else
 * wrote.
 */
export function groupTitleFrom(messagesNewestFirst: readonly RawMessage[], meMri: string): GroupTitle | undefined {
  const byMri = new Map<string, string>();
  for (const m of messagesNewestFirst) {
    if (!isUserMessage(m)) continue;
    const mri = senderMri(m.from);
    const name = m.imdisplayname?.trim();
    if (mri === undefined || mri === meMri || !name || byMri.has(mri)) continue;
    byMri.set(mri, name);
  }
  const byName = new Map<string, { name: string; mri: string }>();
  for (const [mri, name] of byMri) {
    const key = name.toLowerCase();
    const kept = byName.get(key);
    if (kept === undefined || mri < kept.mri) byName.set(key, { name, mri });
  }
  const participants = [...byName.values()].sort(byNameThenMri).map((p) => p.name);
  if (participants.length === 0) return undefined;
  const named = participants.slice(0, NAMED_PARTICIPANTS);
  const more = participants.length - named.length;
  return {
    title: `${GROUP_TITLE_PREFIX}${named.join(NAME_SEPARATOR)}${more > 0 ? ` +${more}` : ''}`,
    participants,
  };
}

/**
 * The other person's MRI in an untitled 1:1 whose listing does not already
 * name them (they did not send its last message) — the 1:1s worth reading.
 */
export function unnamedOneOnOnePeer(c: RawConversation, meMri: string): string | undefined {
  if (conversationKind(c) !== 'oneOnOne' || c.threadProperties?.topic?.trim()) return undefined;
  const peer = oneOnOnePeerMri(c.id, meMri);
  if (peer === undefined) return undefined;
  const last = c.lastMessage;
  const named = last !== undefined && isUserMessage(last) && isFromMri(last.from, peer) && !!last.imdisplayname?.trim();
  return named ? undefined : peer;
}

/**
 * A 1:1's title from its messages (newest first): the other person's most
 * recent display name. `undefined` when they wrote nothing in what was read.
 */
export function oneOnOneTitleFrom(messagesNewestFirst: readonly RawMessage[], peerMri: string): GroupTitle | undefined {
  for (const m of messagesNewestFirst) {
    const name = m.imdisplayname?.trim();
    if (isUserMessage(m) && isFromMri(m.from, peerMri) && name) return { title: name, participants: [name] };
  }
  return undefined;
}

/** A cached read: the chat's last message id when read, and what it found (null: no one else). */
interface CacheEntry {
  lastMessageId: string;
  title: GroupTitle | null;
}

/** One cache per client — in production the process's only one; in tests, one per fake. */
const caches = new WeakMap<ChatServiceClient, Map<string, CacheEntry>>();

function cacheFor(chatsvc: ChatServiceClient): Map<string, CacheEntry> {
  let cache = caches.get(chatsvc);
  if (cache === undefined) {
    cache = new Map();
    caches.set(chatsvc, cache);
  }
  return cache;
}

/**
 * Participant titles for the untitled group chats and unnamed 1:1s among `rows`, by id. Cached
 * chats whose last message has not changed are not read again; a chat whose
 * read fails is left out (and not cached, so the next call retries it), and
 * so is one whose listing has no last message id (it is read every time).
 */
export async function groupTitles(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  rows: readonly Pick<ConversationSummary, 'id' | 'kind'>[],
): Promise<Map<string, GroupTitle | null>> {
  const byId = new Map(raw.map((c) => [c.id, c]));
  const cache = cacheFor(chatsvc);
  const out = new Map<string, GroupTitle | null>();
  const toRead: { id: string; lastMessageId: string | undefined; titleFrom: (m: readonly RawMessage[]) => GroupTitle | undefined }[] = [];
  for (const row of rows) {
    const conv = byId.get(row.id);
    if (conv === undefined || out.has(row.id)) continue;
    const peer = row.kind === 'oneOnOne' ? unnamedOneOnOnePeer(conv, me) : undefined;
    const titleFrom =
      peer !== undefined
        ? (m: readonly RawMessage[]) => oneOnOneTitleFrom(m, peer)
        : row.kind === 'group' && isUntitledGroup(conv)
          ? (m: readonly RawMessage[]) => groupTitleFrom(m, me)
          : undefined;
    if (titleFrom === undefined) continue;
    // With no last message id there is nothing to tell a stale entry by, so it is never cached.
    const lastMessageId = conv.lastMessage?.id || undefined;
    const hit = lastMessageId === undefined ? undefined : cache.get(row.id);
    if (hit !== undefined && hit.lastMessageId === lastMessageId) out.set(row.id, hit.title);
    else if (!toRead.some((r) => r.id === row.id)) toRead.push({ id: row.id, lastMessageId, titleFrom });
  }

  await mapWithConcurrency(toRead, TITLE_CONCURRENCY, async ({ id, lastMessageId, titleFrom }) => {
    try {
      const messages = await chatsvc.getMessages(id, { max: TITLE_FETCH_MESSAGES, pageSize: TITLE_PAGE_SIZE });
      const title = titleFrom(messages) ?? null;
      if (lastMessageId !== undefined) cache.set(id, { lastMessageId, title });
      else cache.delete(id);
      out.set(id, title);
    } catch {
      // Keep the listing's title; the next call reads it again.
    }
  });
  return out;
}

/**
 * `rows` with every untitled group chat titled by its participants, and every
 * unnamed 1:1 by the other person (and given `participants`). A group chat no
 * one else wrote in is "(untitled group chat)", a 1:1 whose other person wrote
 * nothing recent "(1:1 chat)"; one whose read failed keeps its title.
 * Everything else is returned as it was.
 */
export async function enrichTitles<T extends ConversationSummary>(
  chatsvc: ChatServiceClient,
  raw: readonly RawConversation[],
  me: string,
  rows: readonly T[],
): Promise<T[]> {
  const titles = await groupTitles(chatsvc, raw, me, rows);
  return rows.map((row) => {
    if (!titles.has(row.id)) return row;
    const t = titles.get(row.id);
    if (t) return { ...row, title: t.title, participants: t.participants };
    return { ...row, title: row.kind === 'oneOnOne' ? UNTITLED_ONE_ON_ONE_TITLE : UNTITLED_GROUP_TITLE };
  });
}

/** What a result says a conversation is called: the fields a listing row names it by. */
export interface ConversationName {
  id: string;
  kind: ConversationKind;
  title: string;
  team?: string;
  participants?: string[];
}

/** `{ participants }` when a (titled) summary has them — an untitled chat's — else nothing, for spreading into a row. */
export function participantsOf(s: Pick<ConversationSummary, 'participants'>): { participants?: string[] } {
  return s.participants !== undefined ? { participants: s.participants } : {};
}

/** The naming fields of a (titled) summary, for a result that carries a conversation by id. */
export function conversationName(s: Pick<ConversationSummary, 'id' | 'kind' | 'title' | 'team' | 'participants'>): ConversationName {
  return {
    id: s.id,
    kind: s.kind,
    title: s.title,
    ...(s.team !== undefined ? { team: s.team } : {}),
    ...participantsOf(s),
  };
}

/** Names for conversation ids, from a listing read that is already under way (see {@link conversationNamer}). */
export type ConversationNamer = (ids: readonly string[]) => Promise<Map<string, ConversationName>>;

/**
 * Starts the listing read (conversations and the user's MRI) NOW and returns
 * a namer that uses it, so a tool can overlap that read with its own work —
 * the Activity feed scan — rather than waiting for it afterwards. A failed
 * listing is swallowed as soon as it settles (no unhandled rejection, even
 * if the namer is never called) and names nothing.
 */
export function conversationNamer(chatsvc: ChatServiceClient): ConversationNamer {
  const listing = Promise.all([chatsvc.listConversations(), chatsvc.meMri()]).then(
    ([raw, me]) => ({ raw, me }),
    () => undefined,
  );
  return async (ids) => {
    const out = new Map<string, ConversationName>();
    const wanted = new Set(ids);
    if (wanted.size === 0) return out;
    const read = await listing;
    if (read === undefined) return out;
    try {
      const rows = summariseAll(read.raw, read.me).filter((s) => wanted.has(s.id));
      for (const s of await enrichTitles(chatsvc, read.raw, read.me, rows)) out.set(s.id, conversationName(s));
    } catch {
      // Unnamed, not failed: the caller still has every id.
    }
    return out;
  };
}

/**
 * Names for the conversations `ids`, for a tool whose result carries ids it
 * did not get from the listing (the Activity feed, an attachment): one listing
 * read, then {@link enrichTitles} on just those conversations, so an untitled
 * group chat gets its participant title and a 1:1 the other person's name. An
 * id the listing does not carry is left out. Never throws — a failed listing
 * names nothing, a failed title read keeps the listing's title — since a name
 * is never worth failing the call over.
 */
export async function nameConversations(chatsvc: ChatServiceClient, ids: readonly string[]): Promise<Map<string, ConversationName>> {
  if (ids.length === 0) return new Map();
  return conversationNamer(chatsvc)(ids);
}
