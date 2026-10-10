// The fpx skill's shell helpers, run for real under bash with a stub `fpx` on
// PATH — no browser, no network, an invented token.
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, statSync, writeFileSync, chmodSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const ENV_SH = new URL('../skills/teams-fpx/references/teams-env.sh', import.meta.url).pathname;

let dir: string;
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'teams-env-test-'));
  // A stub fpx that "captures" an invented, non-JWT header.
  writeFileSync(join(dir, 'fpx'), '#!/bin/sh\necho \'{"authorization": "Bearer not-a-real-token"}\'\n');
  chmodSync(join(dir, 'fpx'), 0o755);
});
afterEach(() => rmSync(dir, { recursive: true, force: true }));

describe('teams_token_refresh', () => {
  it('writes the token file 0600 without changing the caller\'s umask', () => {
    const out = execFileSync(
      'bash',
      ['-c', `umask 022; source "${ENV_SH}"; teams_token_refresh 5 2>/dev/null; umask`],
      { env: { ...process.env, PATH: `${dir}:${process.env.PATH}`, TEAMS_FPX_DIR: join(dir, 'state') }, encoding: 'utf8' },
    );

    expect(out.trim()).toBe('0022');
    expect(statSync(join(dir, 'state', 'curlrc')).mode & 0o777).toBe(0o600);
  });
});
