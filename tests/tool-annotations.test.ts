import { describe, it, expect } from 'vitest';
import { servedTools } from './served-tools.js';

/**
 * The fleet annotation invariants, read off the SERVED `tools/list` rather
 * than a hand-kept list, so a new tool cannot slip past them.
 *
 * `destructiveHint` DEFAULTS TO TRUE whenever `readOnlyHint` is not true, so
 * a write that forgets to declare it is published as destructive and nothing
 * fails — a considered `false` and a forgotten one look identical. Every
 * write must therefore choose. This server is read-only today (the bridge can
 * only read the DOM); this pins that the day a write arrives it declares
 * itself, and that no read claims otherwise.
 */
describe('every tool declares what it does', () => {
  it('covers the full served surface (guards against a registrar being dropped)', async () => {
    expect(await servedTools()).toHaveLength(6);
  });

  it('sets an explicit boolean readOnlyHint and openWorldHint on all of them', async () => {
    const missing = (await servedTools())
      .filter((t) => typeof t.annotations?.readOnlyHint !== 'boolean' || typeof t.annotations?.openWorldHint !== 'boolean')
      .map((t) => t.name);
    expect(missing).toEqual([]);
  });

  it('marks every tool open-world (each one reads the live Teams tab through the bridge)', async () => {
    const closed = (await servedTools()).filter((t) => t.annotations?.openWorldHint !== true).map((t) => t.name);
    expect(closed).toEqual([]);
  });

  it('sets an explicit boolean destructiveHint on every write', async () => {
    const undeclared = (await servedTools())
      .filter((t) => t.annotations?.readOnlyHint !== true && typeof t.annotations?.destructiveHint !== 'boolean')
      .map((t) => t.name);
    expect(undeclared).toEqual([]);
  });

  it('never lets a read claim to be destructive', async () => {
    const contradictory = (await servedTools())
      .filter((t) => t.annotations?.readOnlyHint === true && t.annotations?.destructiveHint === true)
      .map((t) => t.name);
    expect(contradictory).toEqual([]);
  });
});
