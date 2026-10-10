import { describe, it, expect } from 'vitest';
import { servedTools } from './served-tools.js';

/**
 * The fleet annotation invariants, read off the SERVED `tools/list` rather
 * than a hand-kept list, so a new tool cannot slip past them.
 *
 * `destructiveHint` DEFAULTS TO TRUE whenever `readOnlyHint` is not true, so
 * a write that forgets to declare it is published as destructive and nothing
 * fails — a considered `false` and a forgotten one look identical. Every
 * write must therefore choose. The two writes (send, react) are pinned below
 * to the exact hints the plan settled on; everything else must stay a read.
 */
describe('every tool declares what it does', () => {
  it('covers the full served surface (guards against a registrar being dropped)', async () => {
    expect(await servedTools()).toHaveLength(9);
  });

  it('sets an explicit boolean readOnlyHint and openWorldHint on all of them', async () => {
    const missing = (await servedTools())
      .filter((t) => typeof t.annotations?.readOnlyHint !== 'boolean' || typeof t.annotations?.openWorldHint !== 'boolean')
      .map((t) => t.name);
    expect(missing).toEqual([]);
  });

  it('marks every tool open-world (each one reads live Teams data)', async () => {
    const closed = (await servedTools()).filter((t) => t.annotations?.openWorldHint !== true).map((t) => t.name);
    expect(closed).toEqual([]);
  });

  it('sets an explicit boolean destructiveHint on every write', async () => {
    const undeclared = (await servedTools())
      .filter((t) => t.annotations?.readOnlyHint !== true && typeof t.annotations?.destructiveHint !== 'boolean')
      .map((t) => t.name);
    expect(undeclared).toEqual([]);
  });

  it('keeps every tool but the two writes read-only', async () => {
    const writes = (await servedTools())
      .filter((t) => t.annotations?.readOnlyHint !== true)
      .map((t) => t.name)
      .sort();
    expect(writes).toEqual(['teams_react', 'teams_send_message']);
  });

  it('publishes a send as destructive and non-idempotent (a repeat posts twice)', async () => {
    const send = (await servedTools()).find((t) => t.name === 'teams_send_message');
    expect(send?.annotations).toMatchObject({
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: true,
    });
  });

  it('publishes a reaction as non-destructive and idempotent (setting it twice is one reaction)', async () => {
    const react = (await servedTools()).find((t) => t.name === 'teams_react');
    expect(react?.annotations).toMatchObject({
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    });
  });

  it('gates both writes behind a confirmToken', async () => {
    const ungated = (await servedTools())
      .filter((t) => t.annotations?.readOnlyHint !== true)
      .filter((t) => !Object.hasOwn((t.inputSchema.properties ?? {}) as object, 'confirmToken'))
      .map((t) => t.name);
    expect(ungated).toEqual([]);
  });

  it('never lets a read claim to be destructive', async () => {
    const contradictory = (await servedTools())
      .filter((t) => t.annotations?.readOnlyHint === true && t.annotations?.destructiveHint === true)
      .map((t) => t.name);
    expect(contradictory).toEqual([]);
  });
});
