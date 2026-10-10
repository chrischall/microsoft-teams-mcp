/**
 * The conversation listing shared by `teams_list_chats` and
 * `teams_list_teams_and_channels`: one `listConversations` call, normalised
 * against the signed-in user, then filtered by kind, unread state and query.
 */
import { z } from 'zod';
import type { ChatServiceClient } from '../chatsvc.js';
import {
  conversationSummary,
  matchConversations,
  type ConversationKind,
  type ConversationSummary,
} from '../model.js';

export interface ListFilter {
  kinds: readonly ConversationKind[];
  unread_only?: boolean;
  /** With `unread_only`, keep channels hidden in Teams too (default false). */
  include_hidden?: boolean;
  query?: string;
}

export const unreadOnlyArg = z
  .boolean()
  .optional()
  .describe(
    'Only return conversations with a real message (not a call recording or other system event) ' +
      'the user has not read yet (default false). Reading through this MCP never marks anything read.',
  );

export const includeHiddenArg = z
  .boolean()
  .optional()
  .describe(
    'Also count channels the user has hidden in Teams as unread (default false: Teams shows no ' +
      'unread badge for a hidden channel, so neither does this).',
  );

/**
 * Whether an unread conversation counts as unread for the user: a channel
 * hidden in Teams does not, unless the caller asked for hidden ones.
 */
export function countsAsUnread(s: ConversationSummary, includeHidden: boolean | undefined): boolean {
  return s.unread && (includeHidden === true || s.hidden !== true);
}

export function queryArg(what: string) {
  return z
    .string()
    .min(1)
    .optional()
    .describe(`Case-insensitive substring of the ${what}, or an exact conversation id.`);
}

/**
 * Every conversation of the given kinds, as summaries. Unsorted: each tool
 * orders its own result. A team's name comes from its own thread, so the
 * summaries are built against the FULL listing before any filtering.
 */
export async function listSummaries(chatsvc: ChatServiceClient, filter: ListFilter): Promise<ConversationSummary[]> {
  const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
  let rows = raw.map((c) => conversationSummary(c, me, raw)).filter((s) => filter.kinds.includes(s.kind));
  if (filter.unread_only) rows = rows.filter((s) => countsAsUnread(s, filter.include_hidden));
  if (filter.query !== undefined) rows = matchConversations(rows, filter.query);
  return rows;
}
