/**
 * Every test runs with a throwaway HOME, so nothing that falls back to the
 * user's home directory — the token caches, `bridge-scope.json`, the
 * fetchproxy identity — can read or write the real one. A capture test with no
 * env override once wrote a fake SharePoint host into the developer's
 * `~/.microsoft-teams-mcp/bridge-scope.json`.
 */
import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

process.env.HOME = mkdtempSync(join(tmpdir(), 'teams-mcp-home-'));
