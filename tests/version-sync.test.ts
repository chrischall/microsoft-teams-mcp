import { describe, it, expect } from 'vitest';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import { versionSyncTest } from '@chrischall/mcp-utils/test';

const here = dirname(fileURLToPath(import.meta.url));

describe('version sync', () => {
  it('every x-release-please-version marker matches package.json', () => {
    const mismatches = versionSyncTest({
      srcDir: resolve(here, '../src'),
      pkgPath: resolve(here, '../package.json'),
    });
    expect(mismatches).toEqual([]);
  });
});
