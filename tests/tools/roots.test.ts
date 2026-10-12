import { describe, it, expect, vi } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import { fetchOutsideRoots, MAX_ROOT_FETCHES, ROOT_FETCH_CONCURRENCY, rootBudget } from '../../src/tools/roots.js';
import { normaliseMessage } from '../../src/model.js';
import type { RawMessage } from '../../src/chatsvc.js';
import { fakeChatsvc, ME, rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;
const MIN = 60_000;
const CHANNEL = '19:chan0001@thread.tacv2';

const post = (at: number, rootId?: number, text?: string) => rawMsg(at, { conv: CHANNEL, rootId: rootId ?? at, text });
const normalised = (raws: RawMessage[]) => raws.map((m) => normaliseMessage(m, ME));

/** A root store: `getMessage` answers with these roots by id, and undefined for any other. */
function rootStore(roots: RawMessage[]) {
  return vi.fn(async (_conv: string, id: string) => roots.find((r) => r.id === id));
}

describe('fetchOutsideRoots', () => {
  it('fetches each reply\'s root that the window does not hold, once, and only those', async () => {
    const oldRoot = post(T0 - 10 * MIN, undefined, 'old root');
    const getMessage = rootStore([oldRoot]);
    const messages = normalised([
      post(T0, T0 - 10 * MIN, 'reply 1'),
      post(T0 + MIN, T0 - 10 * MIN, 'reply 2'),
      post(T0 + 2 * MIN, undefined, 'in-window root'),
      post(T0 + 3 * MIN, T0 + 2 * MIN, 'reply to in-window root'),
    ]);

    const roots = await fetchOutsideRoots(fakeChatsvc([], { getMessage }), CHANNEL, messages, ME, { budget: rootBudget() });

    expect(getMessage).toHaveBeenCalledTimes(1);
    expect(getMessage).toHaveBeenCalledWith(CHANNEL, String(T0 - 10 * MIN));
    expect([...roots.keys()]).toEqual([String(T0 - 10 * MIN)]);
    expect(roots.get(String(T0 - 10 * MIN))).toMatchObject({ id: String(T0 - 10 * MIN), text: 'old root' });
  });

  it('normalises a fetched root with the caller\'s time zone', async () => {
    const getMessage = rootStore([post(T0 - 10 * MIN)]);
    const messages = normalised([post(T0, T0 - 10 * MIN)]);

    const roots = await fetchOutsideRoots(fakeChatsvc([], { getMessage }), CHANNEL, messages, ME, {
      budget: rootBudget(),
      timeZone: 'America/New_York',
    });

    expect(roots.get(String(T0 - 10 * MIN))?.time_local).toMatch(/-04:00$/);
  });

  it('leaves out a root that is gone, unreadable, or not a root after all', async () => {
    const getMessage = vi.fn(async (_conv: string, id: string) => {
      if (id === String(T0 - 3 * MIN)) throw new McpToolError('Teams chat service returned HTTP 403.', { status: 403 });
      // Not a root: the service answered with a reply (it names a root of its own).
      if (id === String(T0 - 2 * MIN)) return post(T0 - 2 * MIN, T0 - 9 * MIN);
      // An answer for some other id.
      if (id === String(T0 - 1 * MIN)) return post(T0 - 7 * MIN);
      return undefined;
    });
    const messages = normalised([
      post(T0, T0 - 4 * MIN),
      post(T0 + 1, T0 - 3 * MIN),
      post(T0 + 2, T0 - 2 * MIN),
      post(T0 + 3, T0 - 1 * MIN),
    ]);

    const roots = await fetchOutsideRoots(fakeChatsvc([], { getMessage }), CHANNEL, messages, ME, { budget: rootBudget() });

    expect(getMessage).toHaveBeenCalledTimes(4);
    expect(roots.size).toBe(0);
  });

  it('spends at most the budget, newest threads first, and shares it across calls', async () => {
    const rootIds = Array.from({ length: MAX_ROOT_FETCHES + 5 }, (_, i) => T0 - (i + 1) * MIN);
    const getMessage = rootStore(rootIds.map((id) => post(id)));
    // One reply per root; the reply to the newest root (rootIds[0]) is the newest message.
    const messages = normalised(rootIds.map((id, i) => post(T0 + (rootIds.length - i) * 1000, id))).reverse();
    const budget = rootBudget();

    const roots = await fetchOutsideRoots(fakeChatsvc([], { getMessage }), CHANNEL, messages, ME, { budget });

    expect(getMessage).toHaveBeenCalledTimes(MAX_ROOT_FETCHES);
    expect([...roots.keys()].sort()).toEqual(rootIds.slice(0, MAX_ROOT_FETCHES).map(String).sort());
    expect(budget.remaining).toBe(0);

    const again = await fetchOutsideRoots(fakeChatsvc([], { getMessage }), CHANNEL, messages, ME, { budget });
    expect(again.size).toBe(0);
    expect(getMessage).toHaveBeenCalledTimes(MAX_ROOT_FETCHES);
  });

  it('runs at most a few root fetches at once', async () => {
    let inFlight = 0;
    let peak = 0;
    const getMessage = vi.fn(async (_conv: string, id: string) => {
      inFlight++;
      peak = Math.max(peak, inFlight);
      await new Promise((r) => setTimeout(r, 2));
      inFlight--;
      return post(Number(id));
    });
    const messages = normalised(Array.from({ length: 12 }, (_, i) => post(T0 + i, T0 - (i + 1) * MIN)));

    const roots = await fetchOutsideRoots(fakeChatsvc([], { getMessage }), CHANNEL, messages, ME, { budget: rootBudget() });

    expect(roots.size).toBe(12);
    expect(peak).toBeLessThanOrEqual(ROOT_FETCH_CONCURRENCY);
  });

  it('fetches nothing when every root is in the window', async () => {
    const getMessage = vi.fn();
    const messages = normalised([post(T0), post(T0 + MIN, T0)]);

    const roots = await fetchOutsideRoots(fakeChatsvc([], { getMessage }), CHANNEL, messages, ME, { budget: rootBudget() });

    expect(roots.size).toBe(0);
    expect(getMessage).not.toHaveBeenCalled();
  });
});
