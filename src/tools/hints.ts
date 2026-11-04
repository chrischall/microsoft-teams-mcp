/**
 * Context-group hints, published on each tool as `_meta.mint` (mcp-host's
 * `mintHintSchema`, spec 2026-11-03 D4). Inert data: a host that does not read
 * them sees an extra `_meta` key and nothing else. The shape is fixed by the
 * host's contract, so it is restated here rather than imported — this server
 * takes no dependency on mcp-host.
 *
 * Every tool here is in the `work` group. `subject` and `when` name THIS
 * tool's own argument names, never values; none is declared, because Teams
 * serves one signed-in account (no listable subjects) and the one canonical
 * read takes a cursor, not a start and end. `aliases` are the words a person
 * says — lowercase, at most 8, each at most 40 chars.
 */
export type MintIntent =
  | 'whats-new' | 'whats-due' | 'whats-on' | 'lookup' | 'turn-on-off' | 'send' | 'change' | 'other';

export interface MintHint {
  domain: 'work';
  intent: MintIntent;
  canonical?: true;
  subject?: string;
  when?: { start: string; end: string };
  aliases?: string[];
}

/** `_meta` for a Teams tool. */
export function workMeta(
  intent: MintIntent,
  extra: Omit<MintHint, 'domain' | 'intent'> = {},
): { mint: MintHint } {
  return { mint: { domain: 'work', intent, ...extra } };
}
