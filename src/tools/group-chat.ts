/**
 * Group-chat housekeeping through the chat-service thread API (verified live
 * 2026-11-09): `teams_rename_chat` sets a group chat's topic,
 * `teams_add_chat_members` adds people to it and `teams_remove_chat_members`
 * takes people out of it.
 *
 * Each is gated by `confirmWrite` as `teams_react` is: Teams posts a notice
 * in the chat for every change, so everyone in it sees what was done. All
 * three are idempotent (the same name, an existing member, an absent member
 * change nothing); only removal is declared destructive — it takes someone's
 * access to the chat away.
 *
 * Only group chats are touched: a 1:1 cannot take more people, a meeting
 * chat's roster follows its meeting, a channel's follows its team, and the
 * self-chat has no one else in it.
 *
 * People are resolved before anything is previewed — every one, or the call
 * is refused (`people.ts`). After the writes the thread is read back once and
 * each change looked for; a write Teams did not confirm (a timeout, a dropped
 * connection, a 5xx) is settled by that read-back, as for a reaction.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import {
  CONFIRM_FLOW_SENTENCE,
  CONFIRM_INJECTION_RULE,
  confirmTokenParam,
  confirmWrite,
  McpToolError,
  toolAnnotations,
  WriteOutcomeUnknownError,
  writeOutcomeStatus,
} from '@chrischall/mcp-utils';
import type { ChatServiceClient, RawThread } from '../chatsvc.js';
import type { TeamsDeps } from '../deps.js';
import { wrapApiError } from './errors.js';
import { workMeta } from './hints.js';
import { knownPeople, resolvePeople, type ResolvedPerson } from './people.js';
import { resolveTarget, SELF_CHAT, type Target } from './send.js';
import { untrustedField, untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

const RENAME = 'teams_rename_chat';
const ADD = 'teams_add_chat_members';
const REMOVE = 'teams_remove_chat_members';

/** Most people one call adds or removes. */
const MAX_PEOPLE = 20;

const KIND_LABELS: Readonly<Record<string, string>> = {
  oneOnOne: 'a 1:1 chat',
  meeting: 'a meeting chat',
  channel: 'a channel',
  self: 'your notes self-chat',
};

const conversationParam = z
  .string()
  .min(1)
  .describe(
    'The group chat: its id (from teams_list_conversations), or a case-insensitive substring of its title or ' +
      "an untitled chat's participant names.",
  );

const peopleParam = z
  .array(z.string().trim().min(1))
  .min(1)
  .max(MAX_PEOPLE)
  .describe(
    'Who: each an MRI (8:orgid:…), a bare Entra object id, or a display name ("Greg King" and "King, Greg" are ' +
      'the same) matched among the other people of your 1:1 chats and recent senders in this chat.',
  );

const PEOPLE_SENTENCE =
  'people are MRIs (8:orgid:…), bare object ids, or display names matched (ignoring case, in either ' +
  '"First Last" or "Last, First" order) among the other person of each of the user\'s 1:1 chats and the ' +
  "chat's recent senders — there is no directory search, so anyone else needs their MRI; an ambiguous or " +
  'unknown name refuses the whole call, listing candidates. ';

/** The conversation, refused unless it is a group chat. `doing` words the refusal ("renamed"). */
async function groupChat(chatsvc: ChatServiceClient, conversation: string, me: string, doing: string): Promise<Target> {
  const target = await resolveTarget(chatsvc, conversation, me, `change`);
  if (target.kind === 'group') return target;
  const what = KIND_LABELS[target.kind] ?? `a ${target.kind}`;
  throw new McpToolError(
    `${JSON.stringify(target.title)} is ${what}, not a group chat: only a group chat can be ${doing}.`,
    {
      hint:
        target.kind === 'oneOnOne'
          ? 'A 1:1 cannot take more people or a name; start a group chat in Teams instead.'
          : 'Pass a group chat (kind "group" in teams_list_conversations).',
    },
  );
}

/** Why Teams did not confirm a write, for a warning. */
function unconfirmedWhy(e: WriteOutcomeUnknownError): string {
  const status = writeOutcomeStatus(e);
  return e.timedOut ? 'it did not answer in time' : status !== undefined ? `it answered HTTP ${status}` : 'the connection dropped';
}

/** The thread's members' MRIs, lower-cased. */
function memberSet(thread: RawThread): Set<string> {
  return new Set((thread.members ?? []).map((m) => m.id.toLowerCase()));
}

/** A person as a preview or a result shows them. */
function shown(p: ResolvedPerson): ResolvedPerson {
  return { person: p.person, mri: p.mri, ...(p.name !== undefined ? { name: p.name } : {}) };
}

/** People resolved and de-duplicated by MRI (the first naming of each kept). */
async function resolveFor(
  chatsvc: ChatServiceClient,
  me: string,
  conversationId: string,
  people: readonly string[],
): Promise<ResolvedPerson[]> {
  const known = await knownPeople(chatsvc, me, conversationId);
  const seen = new Set<string>();
  return resolvePeople(people, known).filter((p) => {
    const key = p.mri.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

type MemberStatus = 'added' | 'removed' | 'already_member' | 'not_member' | 'unknown' | 'failed';

interface MemberResult extends ResolvedPerson {
  status: MemberStatus;
  verified?: boolean | null;
  error?: string;
}

/**
 * Run one write per person, read the thread back once, and report each:
 * `done` (shown, or Teams confirmed it — `verified` saying whether the
 * read-back showed it, null when the read-back failed), `unknown` (Teams did
 * not confirm and the read-back does not show it) or `failed` (refused).
 */
async function applyMemberChanges(
  chatsvc: ChatServiceClient,
  threadId: string,
  people: readonly ResolvedPerson[],
  write: (mri: string) => Promise<void>,
  done: 'added' | 'removed',
): Promise<{ results: MemberResult[]; warning?: string }> {
  const outcomes: { p: ResolvedPerson; unconfirmed?: WriteOutcomeUnknownError; error?: string }[] = [];
  for (const p of people) {
    try {
      await write(p.mri);
      outcomes.push({ p });
    } catch (e) {
      if (e instanceof WriteOutcomeUnknownError) outcomes.push({ p, unconfirmed: e });
      else outcomes.push({ p, error: e instanceof Error ? e.message : String(e) });
    }
  }

  let members: Set<string> | undefined;
  try {
    members = memberSet(await chatsvc.getThread(threadId));
  } catch {
    members = undefined;
  }
  const changed = (mri: string) => {
    const present = members?.has(mri.toLowerCase());
    return present === undefined ? null : done === 'added' ? present : !present;
  };

  const notes: string[] = [];
  const results: MemberResult[] = outcomes.map(({ p, unconfirmed, error }) => {
    if (error !== undefined) return { ...shown(p), status: 'failed', error };
    const verified = changed(p.mri);
    if (unconfirmed !== undefined && verified !== true) {
      notes.push(
        `Teams did not confirm the change for ${p.mri} (${unconfirmedWhy(unconfirmed)}), and reading the chat back ` +
          `${verified === null ? 'failed' : `did not show it ${done}`}.`,
      );
      return { ...shown(p), status: 'unknown', verified: verified ?? false };
    }
    if (unconfirmed !== undefined) {
      notes.push(`Teams did not confirm the change for ${p.mri}, but reading the chat back shows it ${done}.`);
    } else if (verified === false) {
      notes.push(`Teams accepted the change for ${p.mri}, but reading the chat back did not show it ${done}.`);
    }
    return { ...shown(p), status: done, verified };
  });

  if (members === undefined) notes.push('Teams accepted the changes, but reading the chat back failed.');
  if (results.some((r) => r.status === 'failed')) notes.push("Some changes failed: see each person's error.");
  if (notes.length === 0) return { results };
  return {
    results,
    warning:
      `${notes.join(' ')} Check the chat before retrying — running this tool again is safe: anyone already ` +
      `${done === 'added' ? 'in the chat is reported as already_member' : 'gone is reported as not_member'} and not written again.`,
  };
}

export function registerGroupChatTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    RENAME,
    {
      title: 'Rename a Teams group chat',
      description:
        "Rename a group chat (set its topic) as the signed-in user. conversation is the chat's id or name as for " +
        `teams_list_messages; only group chats can be renamed — 1:1s, meeting chats, channels and the self-chat ` +
        `(${SELF_CHAT}) are refused. Teams posts a notice in the chat that the user renamed it, so everyone in it ` +
        'sees the change. Renaming to the current name changes nothing and asks for no confirmation. The result ' +
        'says whether reading the chat back showed the new name (verified); status "unknown" means Teams did not ' +
        'confirm the change and the read-back did not show it — check the chat before retrying. ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows the chat, its current name and the new one. ' +
        CONFIRM_INJECTION_RULE +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: false, destructive: false, idempotent: true, openWorld: true }),
      _meta: workMeta('send', { aliases: ['rename chat', 'rename', 'chat name', 'topic', 'set topic'] }),
      inputSchema: {
        conversation: conversationParam,
        name: z.string().trim().min(1).describe('The new name for the chat (surrounding spaces are trimmed).'),
        confirmToken: confirmTokenParam,
      },
    },
    async ({ conversation, name, confirmToken }, ctx) => {
      try {
        const me = await chatsvc.meMri();
        const target = await groupChat(chatsvc, conversation, me, 'renamed');
        const current = (await chatsvc.getThread(target.id)).properties?.topic ?? '';
        if (current === name) {
          return untrustedResult({ status: 'unchanged', conversation: target, name });
        }

        const path = `/threads/${encodeURIComponent(target.id)}/properties?name=topic`;
        const gate = await confirmWrite(ctx, {
          tool: RENAME,
          action: 'chat.rename',
          message: 'Review and confirm this group chat rename before it is made:',
          // The title is someone else's text: JSON-quoted so it cannot break out of its quotes.
          summary: `Rename group chat ${JSON.stringify(target.title)} to ${JSON.stringify(name)}`,
          // The signed-in user, so an approval minted under one account never acts as another.
          account: me,
          target: target.id,
          request: { method: 'PUT', path, body: { topic: name } },
          // The title and the current name are set by other people, so each is fenced.
          preview: {
            conversation: untrustedField(target),
            current_name: untrustedField(current),
            new_name: name,
            note: 'Teams posts a notice in the chat that you renamed it.',
          },
          confirmToken,
        });
        if (gate) return gate;

        let unconfirmed: WriteOutcomeUnknownError | undefined;
        try {
          await chatsvc.setTopic(target.id, name);
        } catch (e) {
          if (!(e instanceof WriteOutcomeUnknownError)) throw e;
          unconfirmed = e;
        }

        let verified: boolean | null = false;
        try {
          verified = (await chatsvc.getThread(target.id)).properties?.topic === name;
        } catch {
          verified = null;
        }
        const base = { conversation: target, previous_name: current, name };

        if (unconfirmed !== undefined && verified !== true) {
          return untrustedResult({
            status: 'unknown',
            ...base,
            verified: verified ?? false,
            warning:
              `Teams did not confirm this rename (${unconfirmedWhy(unconfirmed)}), and reading the chat back ` +
              `${verified === null ? 'failed' : 'did not show the new name'} — check the chat (teams_list_conversations) ` +
              'before retrying; renaming again to the same name is harmless.',
          });
        }
        let warning: string | undefined;
        if (unconfirmed !== undefined) {
          warning = 'Teams did not confirm this rename, but reading the chat back shows the new name.';
        } else if (verified === null) {
          warning = 'Teams accepted the rename, but reading the chat back failed — check the chat (teams_list_conversations).';
        } else if (!verified) {
          warning = 'Teams accepted the rename, but reading the chat back did not show the new name — check the chat (teams_list_conversations).';
        }
        return untrustedResult({ status: 'renamed', ...base, verified, ...(warning !== undefined ? { warning } : {}) });
      } catch (err) {
        throw wrapApiError(err, 'rename the chat');
      }
    },
  );

  server.registerTool(
    ADD,
    {
      title: 'Add people to a Teams group chat',
      description:
        'Add people to a group chat as the signed-in user, as ordinary members. conversation is the chat\'s id or ' +
        'name as for teams_list_messages; only group chats take people — 1:1s, meeting chats, channels and the ' +
        `self-chat are refused. ${PEOPLE_SENTENCE}People already in the chat are reported as already_member and ` +
        'not written; if everyone is, nothing is asked or changed. Teams posts a notice in the chat for each person ' +
        'added, and they can read the chat\'s history as Teams allows. The result lists each person: added, ' +
        'already_member, unknown (Teams did not confirm and reading the chat back does not show them — check ' +
        'before retrying) or failed (with the error). ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows the chat and exactly who would be added. ' +
        CONFIRM_INJECTION_RULE +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: false, destructive: false, idempotent: true, openWorld: true }),
      _meta: workMeta('send', { aliases: ['add people', 'add to chat', 'invite', 'add members', 'members'] }),
      inputSchema: {
        conversation: conversationParam,
        people: peopleParam,
        confirmToken: confirmTokenParam,
      },
    },
    async ({ conversation, people, confirmToken }, ctx) => {
      try {
        const me = await chatsvc.meMri();
        const target = await groupChat(chatsvc, conversation, me, 'given more people');
        const resolved = await resolveFor(chatsvc, me, target.id, people);
        const members = memberSet(await chatsvc.getThread(target.id));
        const already = resolved.filter((p) => members.has(p.mri.toLowerCase()));
        const toAdd = resolved.filter((p) => !members.has(p.mri.toLowerCase()));
        const alreadyResults: MemberResult[] = already.map((p) => ({ ...shown(p), status: 'already_member' }));
        if (toAdd.length === 0) {
          return untrustedResult({ status: 'unchanged', conversation: target, results: alreadyResults });
        }

        const gate = await confirmWrite(ctx, {
          tool: ADD,
          action: 'chat.add_members',
          message: 'Review and confirm who is added to this group chat:',
          summary: `Add ${toAdd.length} ${toAdd.length === 1 ? 'person' : 'people'} to group chat ${JSON.stringify(target.title)}`,
          account: me,
          target: target.id,
          request: {
            method: 'PUT',
            path: `/threads/${encodeURIComponent(target.id)}/members`,
            body: { members: toAdd.map((p) => ({ id: p.mri, role: 'User' })) },
          },
          // The title and the resolved names are set by other people, so each is fenced.
          preview: {
            conversation: untrustedField(target),
            add: untrustedField(toAdd.map(shown)),
            ...(already.length > 0 ? { already_members: untrustedField(already.map(shown)) } : {}),
            note:
              'Sent as one PUT /threads/{id}/members/{mri} with {"role":"User"} per person. Teams posts a notice in ' +
              'the chat for each person added.',
          },
          confirmToken,
        });
        if (gate) return gate;

        const { results, warning } = await applyMemberChanges(
          chatsvc,
          target.id,
          toAdd,
          (mri) => chatsvc.addMember(target.id, mri, 'User'),
          'added',
        );
        return untrustedResult({
          status: 'done',
          conversation: target,
          results: orderAs(resolved, [...results, ...alreadyResults]),
          ...(warning !== undefined ? { warning } : {}),
        });
      } catch (err) {
        throw wrapApiError(err, 'add people to the chat');
      }
    },
  );

  server.registerTool(
    REMOVE,
    {
      title: 'Remove people from a Teams group chat',
      description:
        'Remove people from a group chat as the signed-in user. conversation is the chat\'s id or name as for ' +
        'teams_list_messages; only group chats are changed — 1:1s, meeting chats, channels and the self-chat are ' +
        `refused, and so is removing the signed-in user (leave the chat in Teams instead). ${PEOPLE_SENTENCE}` +
        'People not in the chat are reported as not_member and not written; if no one is, nothing is asked or ' +
        'changed. Teams posts a notice in the chat for each person removed, and they lose access to it. The ' +
        'result lists each person: removed, not_member, unknown (Teams did not confirm and reading the chat back ' +
        'still shows them — check before retrying) or failed (with the error). ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows the chat and exactly who would be removed. ' +
        CONFIRM_INJECTION_RULE +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: false, destructive: true, idempotent: true, openWorld: true }),
      _meta: workMeta('send', { aliases: ['remove people', 'remove from chat', 'kick', 'remove members', 'members'] }),
      inputSchema: {
        conversation: conversationParam,
        people: peopleParam,
        confirmToken: confirmTokenParam,
      },
    },
    async ({ conversation, people, confirmToken }, ctx) => {
      try {
        const me = await chatsvc.meMri();
        const target = await groupChat(chatsvc, conversation, me, 'changed');
        const resolved = await resolveFor(chatsvc, me, target.id, people);
        const self = resolved.find((p) => p.mri.toLowerCase() === me.toLowerCase());
        if (self !== undefined) {
          throw new McpToolError(`Refusing to remove you (the signed-in user, "${self.person}") from the chat.`, {
            hint: 'To leave a group chat, use Leave in Teams. Remove the others by passing them without yourself.',
          });
        }
        const members = memberSet(await chatsvc.getThread(target.id));
        const toRemove = resolved.filter((p) => members.has(p.mri.toLowerCase()));
        const absent: MemberResult[] = resolved
          .filter((p) => !members.has(p.mri.toLowerCase()))
          .map((p) => ({ ...shown(p), status: 'not_member' }));
        if (toRemove.length === 0) {
          return untrustedResult({ status: 'unchanged', conversation: target, results: absent });
        }

        const gate = await confirmWrite(ctx, {
          tool: REMOVE,
          action: 'chat.remove_members',
          message: 'Review and confirm who is removed from this group chat:',
          summary:
            `Remove ${toRemove.length} ${toRemove.length === 1 ? 'person' : 'people'} from group chat ` +
            JSON.stringify(target.title),
          account: me,
          target: target.id,
          request: {
            method: 'DELETE',
            path: `/threads/${encodeURIComponent(target.id)}/members`,
            body: { members: toRemove.map((p) => p.mri) },
          },
          preview: {
            conversation: untrustedField(target),
            remove: untrustedField(toRemove.map(shown)),
            ...(absent.length > 0 ? { not_members: untrustedField(absent.map(shown)) } : {}),
            note:
              'Sent as one DELETE /threads/{id}/members/{mri} per person. Teams posts a notice in the chat for each ' +
              'person removed.',
          },
          confirmToken,
        });
        if (gate) return gate;

        const { results, warning } = await applyMemberChanges(
          chatsvc,
          target.id,
          toRemove,
          (mri) => chatsvc.removeMember(target.id, mri),
          'removed',
        );
        return untrustedResult({
          status: 'done',
          conversation: target,
          results: orderAs(resolved, [...results, ...absent]),
          ...(warning !== undefined ? { warning } : {}),
        });
      } catch (err) {
        throw wrapApiError(err, 'remove people from the chat');
      }
    },
  );
}

/** `results` in the order the people were asked for. */
function orderAs(people: readonly ResolvedPerson[], results: readonly MemberResult[]): MemberResult[] {
  return people.map((p) => results.find((r) => r.mri === p.mri)!);
}
