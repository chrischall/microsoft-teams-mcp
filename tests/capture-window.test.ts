import { afterEach, describe, expect, it, vi } from 'vitest';

const { bridgeDeadlineMs, captureWindowMs, DEFAULT_CAPTURE_WINDOW_MS } = await import('../src/capture-window.js');

afterEach(() => {
  vi.unstubAllEnvs();
});

describe('captureWindowMs', () => {
  it('defaults to 30s', () => {
    expect(captureWindowMs({})).toBe(30_000);
    expect(DEFAULT_CAPTURE_WINDOW_MS).toBe(30_000);
  });

  it('reads TEAMS_CAPTURE_TIMEOUT as seconds', () => {
    expect(captureWindowMs({ TEAMS_CAPTURE_TIMEOUT: '120' })).toBe(120_000);
  });

  it('falls back to the default for junk or 0', () => {
    expect(captureWindowMs({ TEAMS_CAPTURE_TIMEOUT: 'soon' })).toBe(30_000);
    expect(captureWindowMs({ TEAMS_CAPTURE_TIMEOUT: '0' })).toBe(30_000);
  });

  it('reads process.env at call time, not import time', () => {
    vi.stubEnv('TEAMS_CAPTURE_TIMEOUT', '90');
    expect(captureWindowMs()).toBe(90_000);
  });
});

describe('bridgeDeadlineMs', () => {
  it('outlasts the capture window by 15s, and is never under the library 30s', () => {
    expect(bridgeDeadlineMs(30_000)).toBe(45_000);
    expect(bridgeDeadlineMs(120_000)).toBe(135_000);
    expect(bridgeDeadlineMs(1_000)).toBe(30_000);
  });
});
