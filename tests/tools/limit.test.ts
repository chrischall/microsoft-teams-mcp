import { describe, it, expect } from 'vitest';
import { makeLimitedStopAt } from '../../src/tools/limit.js';
import { isUserMessage } from '../../src/model.js';
import { rawMsg } from './fake-chatsvc.js';

const T0 = 1_760_000_000_000;

describe('makeLimitedStopAt', () => {
  it('stops at the read horizon, excluding it', () => {
    const limited = makeLimitedStopAt(T0 + 2, isUserMessage, 10);

    expect(limited.stopAt(rawMsg(T0 + 3))).toBe(false);
    expect(limited.stopAt(rawMsg(T0 + 2))).toBe(true);
    expect(limited.truncated).toBe(false);
  });

  it('never stops at a horizon when none is given', () => {
    const limited = makeLimitedStopAt(undefined, isUserMessage, 10);

    expect(limited.stopAt(rawMsg(1))).toBe(false);
  });

  it('counts only wanted messages and stops at the first one past the limit, marking truncated', () => {
    const limited = makeLimitedStopAt(undefined, isUserMessage, 2);

    expect(limited.stopAt(rawMsg(T0 + 9))).toBe(false);
    expect(limited.stopAt(rawMsg(T0 + 8, { type: 'Event/Call' }))).toBe(false);
    expect(limited.stopAt(rawMsg(T0 + 7))).toBe(false);
    expect(limited.truncated).toBe(false);
    expect(limited.stopAt(rawMsg(T0 + 6))).toBe(true);
    expect(limited.truncated).toBe(true);
  });

  it('lets the horizon win over the limit: a boundary at the horizon is not a truncation', () => {
    const limited = makeLimitedStopAt(T0 + 5, isUserMessage, 1);

    expect(limited.stopAt(rawMsg(T0 + 6))).toBe(false);
    expect(limited.stopAt(rawMsg(T0 + 5))).toBe(true);
    expect(limited.truncated).toBe(false);
  });
});
