/**
 * What `teams_edit_message` and `teams_delete_message` share: finding one of
 * the user's own messages (refusing anything else before any preview), and
 * settling a write by reading the message back.
 */
import { McpToolError, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import type { ChatServiceClient, RawMessage } from '../chatsvc.js';
import { isUserMessage, senderMri } from '../model.js';

export type OwnMessageVerb = 'edit' | 'delete';
const PARTICIPLE: Record<OwnMessageVerb, string> = { edit: 'edited', delete: 'deleted' };

/**
 * One of the user's own, person-written messages — the only kind this server
 * edits or deletes — or the reason it is not one. Whether it is already
 * deleted is the caller's to judge.
 */
export async function ownMessage(
  chatsvc: ChatServiceClient,
  conversationId: string,
  id: string,
  me: string,
  verb: OwnMessageVerb,
): Promise<RawMessage> {
  const m = await chatsvc.getMessage(conversationId, id);
  if (m === undefined) {
    throw new McpToolError(`There is no message ${id} in this conversation to ${verb}.`, {
      hint: 'Take message ids from teams_list_messages.',
    });
  }
  if (senderMri(m.from)?.toLowerCase() !== me.toLowerCase()) {
    throw new McpToolError(`Message ${id} was not sent by you: you can ${verb} only your own messages.`, {
      hint: 'Pick one of your own messages (fromMe: true in teams_list_messages).',
    });
  }
  if (!isUserMessage(m)) {
    throw new McpToolError(`Message ${id} is a system event (${m.messagetype}), not a message anyone wrote.`, {
      hint: `Only messages you wrote can be ${PARTICIPLE[verb]}.`,
    });
  }
  return m;
}

/** How a write to one message went, as the tool's result reports it. */
export interface Settled {
  status: 'done' | 'unknown';
  verified: boolean;
  warning?: string;
}

/**
 * Make the write, then read the message back. A write Teams did not confirm
 * (a `WriteOutcomeUnknownError`) is settled by the read-back: `done` with a
 * warning when it shows the change, `unknown` when it does not. Any other
 * failure is a real failure and throws.
 *
 * `shows` judges the read-back (`undefined`: the message is gone, a 404);
 * `verb` words the warnings ("edit", "delete"), `change` what the read-back
 * should show ("the new text", "it deleted").
 */
export async function writeAndReadBack(
  write: () => Promise<void>,
  readBack: () => Promise<RawMessage | undefined>,
  shows: (m: RawMessage | undefined) => boolean,
  verb: OwnMessageVerb,
  change: string,
): Promise<Settled> {
  let unconfirmed = false;
  try {
    await write();
  } catch (e) {
    if (!(e instanceof WriteOutcomeUnknownError)) throw e;
    unconfirmed = true;
  }

  let verified = false;
  let readBackFailed = false;
  try {
    verified = shows(await readBack());
  } catch {
    readBackFailed = true;
  }

  const check = 'check the conversation (teams_list_messages)';
  if (unconfirmed && !verified) {
    return {
      status: 'unknown',
      verified,
      warning:
        `Teams did not confirm this ${verb}, and reading the message back ${readBackFailed ? 'failed' : `did not show ${change}`} — ` +
        `${check} before retrying.`,
    };
  }
  const warning = unconfirmed
    ? `Teams did not confirm this ${verb}, but reading the message back shows ${change}.`
    : readBackFailed
      ? `Teams accepted the ${verb}, but reading the message back failed — ${check}.`
      : !verified
        ? `Teams accepted the ${verb}, but reading the message back did not show ${change} — ${check}.`
        : undefined;
  return { status: 'done', verified, ...(warning !== undefined ? { warning } : {}) };
}
