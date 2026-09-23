/**
 * Conversation identity for the "open chat/channel" tools. Both tools can
 * only read whatever is currently displayed (see `client.ts`), so their rows
 * are only useful if the model can check they came from the conversation the
 * user meant. This reads the selected sidebar row and turns it into a
 * `conversation` field plus a `conversation_check` instruction.
 */
import type { OpenConversationRow, TeamsClient } from '../client.js';

export const OPEN_CONVERSATION_DESCRIPTION =
  ' The result names the open conversation (`conversation.title`, and ' +
  '`conversation.conversationKey`, which embeds the thread id): before summarizing, confirm ' +
  'it is the one the user asked about, and if it is not — or is null — tell the user which ' +
  'conversation was read and ask them to open the right one. If the user has more than one ' +
  'teams.cloud.microsoft tab open, the bridge reads whichever tab answers first, so the ' +
  'content may come from a different tab than the one they are looking at.';

export interface ConversationIdentity {
  conversation: OpenConversationRow | null;
  conversation_check: string;
}

/**
 * Best-effort: the identity read must never cost the user the messages
 * themselves, so a failed read degrades to `null` plus a warning.
 */
export async function readOpenConversation(
  client: TeamsClient,
  kind: 'chat' | 'channel',
): Promise<ConversationIdentity> {
  let conversation: OpenConversationRow | null = null;
  try {
    conversation = await client.getOpenConversation();
  } catch {
    conversation = null;
  }
  if (conversation?.title) {
    return {
      conversation,
      conversation_check:
        `These rows were read from the ${kind} "${conversation.title}". Confirm that is the ` +
        `${kind} the user asked about before attributing anything to it.`,
    };
  }
  return {
    conversation,
    conversation_check:
      `Could not identify which ${kind} is open. Do not assume these rows belong to the ` +
      `${kind} the user named — tell them you cannot confirm the source and ask them to ` +
      `check which ${kind} is open (and that only one Teams tab is open).`,
  };
}
