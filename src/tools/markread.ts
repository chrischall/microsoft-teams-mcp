/**
 * `teams_mark_read` — move the signed-in user's read marker (the
 * conversation's `consumptionhorizon`) forward, so Teams clears the unread
 * badge — and the pieces the read tools' opt-in `mark_read` share with it.
 *
 * Gated by `confirmWrite` exactly as `teams_react` is. Only the user sees the
 * read marker, it replaces nothing that cannot be read again, and moving it to
 * the same message twice is one move, so it is declared non-destructive and
 * idempotent.
 *
 * The marker only ever moves FORWARDS: a target at or before the current
 * marker is a no-op reported as such, never a write, and never asks for
 * confirmation. (Live 2026-10-12, the service answered a backwards move 200
 * and kept the marker where it was — so sending one would only have reported
 * a move that never happened.)
 */
import type { McpServer, ServerContext } from '@modelcontextprotocol/server';
import { z } from 'zod';
import {
  CONFIRM_FLOW_SENTENCE,
  CONFIRM_INJECTION_RULE,
  confirmTokenParam,
  confirmWrite,
  mapWithConcurrency,
  McpToolError,
  toolAnnotations,
  WriteOutcomeUnknownError,
  writeOutcomeStatus,
} from '@chrischall/mcp-utils';
import { type ChatServiceClient, type RawConversation, type RawMessage } from '../chatsvc.js';
import { clipText } from '../compose.js';
import type { TeamsDeps } from '../deps.js';
import { htmlToText, isUnread, readHorizon, SELF_CHAT, senderMri, type ConversationSummary } from '../model.js';
import { confirmedSummary } from './confirm.js';
import { resolveListed } from './conversation.js';
import { wrapApiError } from './errors.js';
import { untrustedField, untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

const TOOL = 'teams_mark_read';

/** Messages read to find the newest one — one page; a page is not strictly ordered, so all of it is looked at. */
const NEWEST_PAGE = 50;

/** Previewed text of the message the marker moves to. */
const PREVIEW_CHARS = 200;

/** Markers moved at once by a read tool's `mark_read`. */
const MARK_CONCURRENCY = 4;

/** The `mark_read` argument the read tools share. */
export const markReadArg = z
  .boolean()
  .optional()
  .describe(
    'After the read, also mark it read in Teams: move the read marker of each conversation returned up to ' +
      'the newest message returned (default false — the read leaves the unread badges alone). Never moves a ' +
      'marker backwards, and never past an unread message the caller did not receive: a conversation whose ' +
      'read was cut short is left as it was (marked_read: { skipped, hint }). Not allowed with format ' +
      '"summary", which returns no message text. A write, so it asks the user to confirm first, as ' +
      'teams_mark_read does.',
  );

/**
 * Refuse `mark_read` with `format: "summary"`: the caller would get no message
 * text, so marking would clear badges on messages nobody read.
 */
export function assertMarkReadFormat(markRead: boolean | undefined, format: string | undefined): void {
  if (markRead && format === 'summary') {
    throw new McpToolError('mark_read cannot be combined with format "summary": a summary returns no messages, so nothing would have been read.', {
      hint: 'Read with format "full" or "compact" to mark what is returned, or use teams_mark_read to mark a conversation read deliberately.',
    });
  }
}

/** Why a read tool's `mark_read` left a conversation's marker alone though it returned unread messages. */
export interface MarkSkip {
  /**
   * `truncated`: a limit cut the read, so unread messages older than those
   * returned were not returned. `since`: the `since` cutoff did.
   */
  skipped: 'truncated' | 'since';
  hint: string;
}

/** Description text for a read tool that takes `mark_read`. */
export const MARK_READ_DESCRIPTION =
  'By default this does NOT mark anything read — the user\'s Teams unread badges are unaffected; pass ' +
  'mark_read: true (or use teams_mark_read) to clear them. mark_read never marks a message the caller did ' +
  'not receive: it is refused with format "summary" (compact is fine), and a conversation whose read was ' +
  'cut short, leaving unread messages older than those returned, keeps its marker and reports ' +
  'marked_read: { skipped: "truncated" (or "since"), hint }. ';

/** The request path a marker move PUTs to — the same string the client builds. */
export function horizonPath(conversationId: string): string {
  return `/users/ME/conversations/${encodeURIComponent(conversationId)}/properties?name=consumptionhorizon`;
}

/** The id of the newest of `messages` (a message's id is its arrival ms), or `undefined` for none. */
export function newestMessageId(messages: readonly Pick<RawMessage, 'id'>[]): string | undefined {
  let best: string | undefined;
  for (const m of messages) {
    const n = Number(m.id);
    if (Number.isFinite(n) && (best === undefined || n > Number(best))) best = m.id;
  }
  return best;
}

/** The current marker as a message id (its arrival ms), or `null` for a conversation never opened. */
export function horizonMessageId(conv: RawConversation): string | null {
  const h = readHorizon(conv);
  return h > 0 ? String(h) : null;
}

/** Whether moving the marker to `target` moves it forwards. */
export function advancesHorizon(conv: RawConversation, target: string | undefined): target is string {
  if (target === undefined) return false;
  const n = Number(target);
  return Number.isFinite(n) && n > readHorizon(conv);
}

/** One marker move: a conversation and the message its marker moves up to. */
export interface MarkPlan {
  summary: ConversationSummary;
  /** The marker before the move, or `null` when there was none. */
  previous: string | null;
  target: string;
}

/**
 * The confirmation for a read tool's `mark_read`: one approval covering every
 * move, each conversation's title fenced as a read tool's result is.
 * Resolves `undefined` when the moves may proceed (see `confirmWrite`).
 */
export async function confirmMarks(
  ctx: ServerContext,
  opts: { tool: string; me: string; plans: readonly MarkPlan[]; confirmToken: string | undefined; note?: string },
) {
  const { plans } = opts;
  const many = plans.length > 1;
  return confirmWrite(ctx, {
    tool: opts.tool,
    action: 'chat.mark_read',
    message: 'Review and confirm marking these Teams messages read:',
    summary: many
      ? `Mark ${plans.length} conversations read up to their newest returned message`
      : `Mark a ${plans[0].summary.kind} ${JSON.stringify(plans[0].summary.title)} read up to message ${plans[0].target}`,
    account: opts.me,
    ...(many
      ? { payload: plans.map((p) => ({ conversation: p.summary.id, up_to_message_id: p.target })) }
      : { target: `${plans[0].summary.id}/${plans[0].target}`, request: { method: 'PUT', path: horizonPath(plans[0].summary.id) } }),
    preview: {
      conversations: untrustedField(
        plans.map((p) => ({
          id: p.summary.id,
          title: p.summary.title,
          kind: p.summary.kind,
          previous_horizon_message_id: p.previous,
          up_to_message_id: p.target,
        })),
      ),
      ...(opts.note !== undefined ? { note: opts.note } : {}),
    },
    confirmToken: opts.confirmToken,
  });
}

/** What a read tool reports per conversation after `mark_read`. */
export interface MarkOutcome {
  marked_read: boolean | MarkSkip;
  mark_read_warning?: string;
}

/** Why a write went unconfirmed, for a warning. */
function unconfirmedWhy(e: WriteOutcomeUnknownError): string {
  const status = writeOutcomeStatus(e);
  return e.timedOut ? 'it did not answer in time' : status !== undefined ? `it answered HTTP ${status}` : 'the connection dropped';
}

/**
 * Make every move, a few at a time. One failing never costs the others: an
 * unconfirmed or refused move is `marked_read: false` with a warning.
 */
export async function applyMarks(chatsvc: ChatServiceClient, plans: readonly MarkPlan[]): Promise<Map<string, MarkOutcome>> {
  const results = await mapWithConcurrency(plans, MARK_CONCURRENCY, async (p): Promise<MarkOutcome> => {
    try {
      await chatsvc.setConsumptionHorizon(p.summary.id, p.target);
      return { marked_read: true };
    } catch (e) {
      if (e instanceof WriteOutcomeUnknownError) {
        return {
          marked_read: false,
          mark_read_warning:
            `Teams did not confirm moving the read marker (${unconfirmedWhy(e)}); it may have moved — ` +
            'retrying is harmless.',
        };
      }
      return { marked_read: false, mark_read_warning: wrapApiError(e, 'mark this conversation read').message };
    }
  });
  return new Map(plans.map((p, i) => [p.summary.id, results[i]]));
}

export function registerMarkReadTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    TOOL,
    {
      title: 'Mark a Teams conversation read',
      description:
        'Mark a chat or channel read in Teams, clearing its unread badge: moves the signed-in user\'s read ' +
        'marker up to up_to_message_id (from teams_read_conversation), or by default up to the newest message ' +
        'in the conversation. conversation is an id or name as for teams_read_conversation; ' +
        `"${SELF_CHAT}" (or "notes" / "self") is the user's own notes self-chat — use it to try this tool. ` +
        'The marker never moves backwards: when it is already at or past the target, nothing is written and ' +
        'status is "already_read". Returns the conversation, previous_horizon_message_id (null if it was never ' +
        'opened) and new_horizon_message_id; status "unknown" means Teams did not confirm the move — retrying is ' +
        'harmless. teams_read_conversation and teams_get_unread can do the same after a read with mark_read. ' +
        CONFIRM_FLOW_SENTENCE +
        ' The preview shows the conversation, the message the marker moves to and where it was. ' +
        CONFIRM_INJECTION_RULE +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: false, destructive: false, idempotent: true, openWorld: true }),
      inputSchema: {
        conversation: z
          .string()
          .min(1)
          .describe(
            `The conversation id (or "${SELF_CHAT}" — also by the name "notes" or "self" — for the self-chat), ` +
              "or a case-insensitive substring of the chat title, channel name, team name or an untitled chat's participant names.",
          ),
        up_to_message_id: z
          .string()
          .min(1)
          .optional()
          .describe(
            'Mark read up to and including this message, as teams_read_conversation lists it (default: the ' +
              'newest message in the conversation). Messages after it stay unread.',
          ),
        confirmToken: confirmTokenParam,
      },
    },
    async ({ conversation, up_to_message_id, confirmToken }, ctx) => {
      try {
        const [raw, me] = await Promise.all([chatsvc.listConversations(), chatsvc.meMri()]);
        const summary = await resolveListed(chatsvc, raw, me, conversation, 'mark read');
        // The self-chat resolves even when the listing does not carry it.
        const conv = raw.find((c) => c.id === summary.id) ?? { id: summary.id };
        const previous = horizonMessageId(conv);

        const page = await chatsvc.getMessages(summary.id, { max: NEWEST_PAGE, pageSize: NEWEST_PAGE });
        let target: RawMessage | undefined;
        if (up_to_message_id !== undefined) {
          target = await chatsvc.getMessage(summary.id, up_to_message_id);
          if (target === undefined) {
            throw new McpToolError(`There is no message ${up_to_message_id} in this conversation to mark read up to.`, {
              hint: 'Take message ids from teams_read_conversation, or leave up_to_message_id out to mark it all read.',
            });
          }
        } else {
          const newest = newestMessageId(page);
          target = page.find((m) => m.id === newest);
        }

        if (!advancesHorizon(conv, target?.id)) {
          return untrustedResult({
            status: 'already_read',
            conversation: summary,
            previous_horizon_message_id: previous,
            new_horizon_message_id: previous,
            hint:
              target === undefined
                ? 'The conversation has no messages, so there is nothing to mark read.'
                : `The read marker is already at or past message ${target.id}; it never moves backwards, so nothing was changed.`,
          });
        }
        const to = target.id;

        const gate = await confirmWrite(ctx, {
          tool: TOOL,
          action: 'chat.mark_read',
          message: 'Review and confirm marking this Teams conversation read:',
          // The title is someone else's text: JSON-quoted so it cannot break out of its quotes.
          summary: `Mark ${summary.kind} ${JSON.stringify(summary.title)} read up to message ${to}`,
          // The signed-in user, so an approval minted under one account never marks another's chats.
          account: me,
          target: `${summary.id}/${to}`,
          // The client time in the body is minted when the marker moves, so it is not bound.
          request: { method: 'PUT', path: horizonPath(summary.id) },
          preview: {
            conversation: untrustedField({ id: summary.id, title: summary.title, kind: summary.kind }),
            up_to_message: untrustedField({
              id: to,
              sender: target.imdisplayname?.trim() || senderMri(target.from) || 'unknown',
              preview: clipText(htmlToText(target.content ?? ''), PREVIEW_CHARS),
            }),
            previous_horizon_message_id: previous,
          },
          confirmToken,
        });
        if (gate) return gate;

        // What the page shows past the new marker: anything from someone else is still unread.
        const stillUnread = page.some((m) => isUnread(m, Number(to), me));
        const after = confirmedSummary(summary, stillUnread);
        try {
          await chatsvc.setConsumptionHorizon(summary.id, to);
        } catch (e) {
          if (!(e instanceof WriteOutcomeUnknownError)) throw e;
          return untrustedResult({
            status: 'unknown',
            conversation: summary,
            previous_horizon_message_id: previous,
            new_horizon_message_id: to,
            warning:
              `Teams did not confirm moving the read marker (${unconfirmedWhy(e)}); it may have moved. ` +
              'Retrying is harmless — moving it to the same message again is one move.',
          });
        }
        return untrustedResult({
          status: 'marked_read',
          conversation: after,
          previous_horizon_message_id: previous,
          new_horizon_message_id: to,
        });
      } catch (err) {
        throw wrapApiError(err, 'mark the conversation read');
      }
    },
  );
}
