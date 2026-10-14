import { describe, it, expect } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import { decodeFrom, encodeFrom, takePage } from '../../src/tools/from.js';

const T = 1_760_000_000_000;

describe('from positions', () => {
  it('reads an ISO 8601 time as a position with nothing seen', () => {
    expect(decodeFrom('2025-10-09T08:53:20Z')).toEqual({ t: Date.parse('2025-10-09T08:53:20Z'), seen: [] });
    expect(decodeFrom('2025-10-09T10:53:20+02:00')).toEqual({ t: Date.parse('2025-10-09T08:53:20Z'), seen: [] });
  });

  it('reads a time with no zone as UTC, not local time', () => {
    expect(decodeFrom('2025-10-09T08:53:20').t).toBe(Date.parse('2025-10-09T08:53:20Z'));
    expect(decodeFrom('2025-10-09T08:53').t).toBe(Date.parse('2025-10-09T08:53:00Z'));
    expect(decodeFrom('2025-10-09').t).toBe(Date.parse('2025-10-09T00:00:00Z'));
  });

  it('round-trips a next_from token, seen keys included', () => {
    const pos = { t: T, seen: ['19:a@thread.v2/1760000000000'] };
    expect(decodeFrom(encodeFrom(pos))).toEqual(pos);
    expect(decodeFrom(encodeFrom({ t: T, seen: [] }))).toEqual({ t: T, seen: [] });
  });

  it('refuses what is neither a token nor a date, with a hint', () => {
    for (const bad of ['yesterday', '1760000000000', Buffer.from('{"t":"x"}').toString('base64url'), '2025-13-45']) {
      expect(() => decodeFrom(bad), bad).toThrow(McpToolError);
    }
    try {
      decodeFrom('yesterday');
    } catch (err) {
      expect((err as McpToolError).hint).toMatch(/next_from/);
    }
  });
});

describe('takePage', () => {
  const item = (ms: number, key = String(ms)) => ({ ms, key, item: key });

  it('returns the oldest first, at or after the position, up to limit', () => {
    const items = [item(T + 3), item(T - 1), item(T), item(T + 1), item(T + 2)];
    const out = takePage(items, { t: T, seen: [] }, 2);
    expect(out.page).toEqual([String(T), String(T + 1)]);
    expect(out.has_more).toBe(true);
    expect(out.remaining).toBe(2);
    expect(out.next).toEqual({ t: T + 1, seen: [String(T + 1)] });
  });

  it('leaves out what the position already saw at its own time, and nothing else there', () => {
    const items = [item(T, 'a'), item(T, 'b'), item(T, 'c'), item(T + 1, 'd')];
    const out = takePage(items, { t: T, seen: ['a'] }, 10);
    expect(out.page).toEqual(['b', 'c', 'd']);
    expect(out.has_more).toBe(false);
  });

  it('carries every key seen at a time that did not move, so a tie spanning pages is never repeated', () => {
    const items = [item(T, 'a'), item(T, 'b'), item(T, 'c')];
    const first = takePage(items, { t: T, seen: [] }, 2);
    expect(first.page).toEqual(['a', 'b']);
    expect(first.next).toEqual({ t: T, seen: ['a', 'b'] });
    const second = takePage(items, first.next, 2);
    expect(second.page).toEqual(['c']);
    expect(second.has_more).toBe(false);
    expect(second.next).toEqual({ t: T, seen: ['a', 'b', 'c'] });
  });

  it('keeps the position when nothing is new, so a poller can call again', () => {
    const pos = { t: T, seen: ['x'] };
    const out = takePage([item(T - 5)], pos, 10);
    expect(out.page).toEqual([]);
    expect(out.has_more).toBe(false);
    expect(out.next).toEqual(pos);
  });

  it('never moves past ceiling: what arrived after it is read again next time rather than risked', () => {
    const items = [item(T + 1), item(T + 5)];
    const out = takePage(items, { t: T, seen: [] }, 10, T + 3);
    expect(out.page).toEqual([String(T + 1), String(T + 5)]);
    expect(out.next).toEqual({ t: T + 3, seen: [] });
  });

  it('never moves backwards when the ceiling is behind the position', () => {
    const out = takePage([item(T + 1)], { t: T, seen: [] }, 10, T - 100);
    expect(out.next.t).toBe(T);
  });
});
