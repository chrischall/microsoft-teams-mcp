/**
 * `teams_get_activity` — the Activity feed (the bell icon in Teams' left nav),
 * read through the chat-service API rather than off the page.
 *
 * Verified live 2026-10-10 against a real tenant: the feed IS a chat-service
 * conversation, `48:notifications`, and `getMessages` reads it like any other
 * (`listConversations` drops `48:*` ids, so it is read directly by id). 198 of
 * the 200 newest entries carried `properties.activity` — a JSON string, or an
 * object — describing what happened (`activityType` / `activitySubtype`), who
 * did it, where (`sourceThreadId` + `sourceMessageId`, which
 * `teams_read_conversation` takes) and a message preview; `properties.isread`
 * is `"true"` / `"false"`. `48:mentions` holds the mention subset in the same
 * shape, so it is not read separately. An entry with `deletetime` set is
 * deleted and skipped, as is one with no activity payload.
 *
 * Reading the feed never writes its read marker: the bell's badge is exactly
 * as it was.
 *
 * The feed's `sourceThreadTopic` is empty for an untitled group chat or 1:1,
 * so each returned row's conversation is named from the listing instead, the
 * way every other tool names it (`titles.ts`): one listing read plus the
 * bounded, cached title reads for just those conversations. That naming never
 * fails the call — the feed's own topic stands in.
 */
import type { McpServer } from '@modelcontextprotocol/server';
import { z } from 'zod';
import { toolAnnotations } from '@chrischall/mcp-utils';
import type { TeamsDeps } from '../deps.js';
import type { RawMessage } from '../chatsvc.js';
import { htmlToText } from '../model.js';
import { arrivalMs } from './conversation.js';
import { sinceCutoff, DEFAULT_SINCE_DAYS, MAX_SINCE_DAYS } from './confirm.js';
import { limitArg, trimRows } from './limit.js';
import { conversationNamer, type ConversationNamer } from './titles.js';
import { wrapApiError } from './errors.js';
import { untrustedResult, UNTRUSTED_DESCRIPTION_SUFFIX } from './untrusted.js';

/** The chat-service conversation that holds the Activity feed. */
export const ACTIVITY_FEED_ID = '48:notifications';

const MAX_ROWS = 200;
/** Feed entries read before giving up — five pages; the real bound is `since_days`. */
const MAX_RAW_ENTRIES = 1000;
const MAX_PREVIEW_CHARS = 200;

export const ACTIVITY_KINDS = ['mention', 'reaction', 'reply', 'meeting', 'call', 'other'] as const;
export type ActivityKind = (typeof ACTIVITY_KINDS)[number];

/** `activityType` → the bucket the `kinds` filter and each row's `kind` use. */
const KIND_OF_TYPE: Record<string, ActivityKind> = {
  mention: 'mention',
  mentionInChat: 'mention',
  reaction: 'reaction',
  reactionInChat: 'reaction',
  reply: 'reply',
  replyToReply: 'reply',
  // Every msGraph entry seen live was a private-meeting notice
  // (privateMeetingCreated / Updated / Forwarded / …).
  msGraph: 'meeting',
  call: 'call',
};

export function activityKind(type: string): ActivityKind {
  return Object.hasOwn(KIND_OF_TYPE, type) ? KIND_OF_TYPE[type] : 'other';
}

/** The fields of `properties.activity` this tool reads. Every one is optional on the wire. */
interface ActivityPayload {
  activityType?: unknown;
  activitySubtype?: unknown;
  activityTimestamp?: unknown;
  sourceThreadId?: unknown;
  sourceMessageId?: unknown;
  sourceReplyChainId?: unknown;
  sourceUserImDisplayName?: unknown;
  messagePreview?: unknown;
  sourceThreadTopic?: unknown;
  sourceThreadType?: unknown;
}

export interface ActivityRow {
  id: string;
  time?: string;
  kind: ActivityKind;
  type: string;
  subtype: string;
  read: boolean;
  from?: string;
  /**
   * The conversation: its id, its title — as the listing tools name it (an
   * untitled group chat by its participants, a 1:1 by the other person) when
   * the listing carries it, else the feed's own topic — the feed's thread type,
   * and a channel's team and an untitled chat's participants when known.
   */
  where: { conversation_id?: string; title?: string; type?: string; team?: string; participants?: string[] };
  message_id?: string;
  reply_chain_id?: string;
  preview: string;
}

/** `properties.activity` as an object — it arrives as a JSON string or already parsed. */
function activityOf(m: RawMessage): ActivityPayload | undefined {
  const raw = m.properties?.activity;
  let value: unknown = raw;
  if (typeof raw === 'string') {
    try {
      value = JSON.parse(raw);
    } catch {
      return undefined;
    }
  }
  return value !== null && typeof value === 'object' && !Array.isArray(value) ? (value as ActivityPayload) : undefined;
}

/** A wire value as a non-empty string (ids can arrive as numbers), else undefined. */
function str(v: unknown): string | undefined {
  if (typeof v === 'string') return v === '' ? undefined : v;
  if (typeof v === 'number' && Number.isFinite(v)) return String(v);
  return undefined;
}

function isoOrUndefined(v: unknown): string | undefined {
  const s = str(v);
  if (s === undefined) return undefined;
  const ms = /^\d+$/.test(s) ? Number(s) : Date.parse(s);
  return Number.isFinite(ms) ? new Date(ms).toISOString() : undefined;
}

function clip(text: string): string {
  return text.length <= MAX_PREVIEW_CHARS ? text : `${text.slice(0, MAX_PREVIEW_CHARS - 1).trimEnd()}…`;
}

/** One feed entry as a row, or `undefined` for a deleted entry or one with no activity payload. */
export function activityRow(m: RawMessage): ActivityRow | undefined {
  if (m.properties?.deletetime) return undefined;
  const a = activityOf(m);
  if (a === undefined) return undefined;
  const type = str(a.activityType) ?? '';
  const replyChain = str(a.sourceReplyChainId);
  const isread = m.properties?.isread;
  return {
    id: m.id,
    time: isoOrUndefined(a.activityTimestamp) ?? isoOrUndefined(m.originalarrivaltime),
    kind: activityKind(type),
    type,
    subtype: str(a.activitySubtype) ?? '',
    read: isread === true || isread === 'true',
    from: str(a.sourceUserImDisplayName),
    where: {
      conversation_id: str(a.sourceThreadId),
      title: str(a.sourceThreadTopic),
      type: str(a.sourceThreadType),
    },
    message_id: str(a.sourceMessageId),
    ...(replyChain !== undefined ? { reply_chain_id: replyChain } : {}),
    preview: clip(htmlToText(str(a.messagePreview) ?? '')),
  };
}

/**
 * `rows` with each conversation named as the listing tools name it
 * (`conversationNamer`): only the conversations of the rows returned, and a
 * conversation the listing does not carry — or a listing that cannot be read —
 * keeps the feed's own topic.
 */
async function nameRows(name: ConversationNamer, rows: ActivityRow[]): Promise<ActivityRow[]> {
  const names = await name(
    rows.flatMap((r) => (r.where.conversation_id !== undefined ? [r.where.conversation_id] : [])),
  );
  return rows.map((r) => {
    const n = r.where.conversation_id !== undefined ? names.get(r.where.conversation_id) : undefined;
    if (n === undefined) return r;
    return {
      ...r,
      where: {
        ...r.where,
        title: n.title,
        ...(n.team !== undefined ? { team: n.team } : {}),
        ...(n.participants !== undefined ? { participants: n.participants } : {}),
      },
    };
  });
}

const newestFirst = (a: ActivityRow, b: ActivityRow): number => (b.time ?? '').localeCompare(a.time ?? '');

export function registerActivityTools(server: McpServer, { chatsvc }: Pick<TeamsDeps, 'chatsvc'>): void {
  server.registerTool(
    'teams_get_activity',
    {
      title: 'Read the Activity feed',
      description:
        'Read the Teams Activity feed (the bell icon) — mentions, reactions, replies, meeting ' +
        'notices and missed calls across every chat and channel — newest first, through the ' +
        'Teams chat-service API, so it works without the Activity view (or any particular view) ' +
        'open in the browser. Each row has id, time (ISO 8601), kind (mention, reaction, reply, ' +
        'meeting, call or other) with the raw type/subtype, read, from (who did it), where ' +
        '(conversation_id; title — named as teams_list_chats names it, so an untitled group chat is ' +
        '"Group chat: A; B; C +N" with its participants and a 1:1 is the other person; type; team for a channel), ' +
        'message_id, reply_chain_id (a channel thread\'s root, when there is one), ' +
        'and a plain-text preview (at most 200 characters). Pass where.conversation_id to ' +
        'teams_read_conversation to read the message in context. unread_only keeps rows the user ' +
        'has not seen; kinds narrows to some buckets; since_days (default 14) sets how far back ' +
        'to look. This does NOT mark anything read — the user\'s Activity badge is unaffected.' +
        UNTRUSTED_DESCRIPTION_SUFFIX,
      annotations: toolAnnotations({ readOnly: true, openWorld: true }),
      inputSchema: {
        unread_only: z
          .boolean()
          .optional()
          .describe('Only return activity the user has not seen yet (default false). Never marks it read.'),
        kinds: z
          .array(z.enum(ACTIVITY_KINDS))
          .min(1)
          .optional()
          .describe(
            'Only these kinds of activity (default all): mention (incl. @team / @everyone), reaction, ' +
              'reply, meeting (meeting invitations and updates), call (missed calls and meeting nudges), other.',
          ),
        since_days: z
          .number()
          .int()
          .min(1)
          .max(MAX_SINCE_DAYS)
          .optional()
          .describe(`Only activity from the last since_days days (default ${DEFAULT_SINCE_DAYS}, max ${MAX_SINCE_DAYS}).`),
        limit: limitArg(MAX_ROWS, 'activity rows'),
      },
    },
    async ({ unread_only, kinds, since_days, limit }) => {
      try {
        // The listing that names the rows is read alongside the feed scan, not after it.
        const name = conversationNamer(chatsvc);
        const cutoff = sinceCutoff(since_days);
        let scanCapped = false;
        const entries = await chatsvc.getMessages(ACTIVITY_FEED_ID, {
          max: MAX_RAW_ENTRIES,
          stopAt: (m) => {
            const at = arrivalMs(m);
            return Number.isFinite(at) && at < cutoff;
          },
          onCapped: () => {
            scanCapped = true;
          },
        });

        const wanted = kinds !== undefined ? new Set<ActivityKind>(kinds) : undefined;
        const rows = entries
          .map(activityRow)
          .filter((r): r is ActivityRow => r !== undefined)
          .filter((r) => r.time === undefined || Date.parse(r.time) >= cutoff)
          .filter((r) => !unread_only || !r.read)
          .filter((r) => wanted === undefined || wanted.has(r.kind))
          .sort(newestFirst);

        const trimmed = trimRows(rows, limit, 'start', MAX_ROWS);
        return untrustedResult({
          ...trimmed,
          rows: await nameRows(name, trimmed.rows),
          ...(scanCapped
            ? {
                scan_capped: {
                  entries_read: MAX_RAW_ENTRIES,
                  hint:
                    `Only the newest ${MAX_RAW_ENTRIES} feed entries were read, so older activity inside ` +
                    'since_days may be missing; pass a smaller since_days to look at a shorter window.',
                },
              }
            : {}),
        });
      } catch (err) {
        throw wrapApiError(err, 'read the Activity feed');
      }
    },
  );
}
