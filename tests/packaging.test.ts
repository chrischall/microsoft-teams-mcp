import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { versionSyncTest } from '@chrischall/mcp-utils/test';
import { servedTools } from './served-tools.js';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const read = (p: string) => JSON.parse(readFileSync(join(root, p), 'utf8'));
const pkg = read('package.json');

describe('version sync', () => {
  it('keeps every version marker equal to package.json', () => {
    expect(versionSyncTest({ srcDir: join(root, 'src'), pkgPath: join(root, 'package.json') })).toEqual(
      [],
    );
  });

  it('keeps the manifests on the same version', () => {
    const v = pkg.version;
    expect(read('manifest.json').version).toBe(v);
    expect(read('server.json').version).toBe(v);
    expect(read('server.json').packages[0].version).toBe(v);
    expect(read('.claude-plugin/plugin.json').version).toBe(v);
    expect(read('.claude-plugin/marketplace.json').metadata.version).toBe(v);
    expect(read('.claude-plugin/marketplace.json').plugins[0].version).toBe(v);
  });

  it('registers every version-bearing file with release-please', () => {
    const extras = read('release-please-config.json').packages['.']['extra-files'];
    const paths = extras.map((e: unknown) => (typeof e === 'string' ? e : (e as { path: string }).path));
    for (const p of [
      'manifest.json',
      'server.json',
      '.claude-plugin/plugin.json',
      '.claude-plugin/marketplace.json',
      'src/version.ts',
    ]) {
      expect(paths).toContain(p);
    }
  });
});

describe('publish scaffold', () => {
  it('declares the repository url npm provenance validates against', () => {
    expect(pkg.repository?.url).toBe(
      'git+https://github.com/chrischall/microsoft-teams-mcp.git',
    );
  });

  it('publishes under the chrischall scope with public access', () => {
    expect(pkg.name).toBe('@chrischall/microsoft-teams-mcp');
    expect(pkg.publishConfig?.access).toBe('public');
  });

  it('points the plugin at a config that resolves under a plugin install', () => {
    const pluginMcp = read('.claude-plugin/plugin.json').mcp as string;
    const cfg = read(join('.claude-plugin', pluginMcp.replace(/^\.\//, '')));
    const server = cfg.mcpServers.teams;
    expect(server.args.join(' ')).toContain('${CLAUDE_PLUGIN_ROOT}');
    expect(server.args.join(' ')).toContain('dist/bundle.js');
  });

  it('ships the files an install and a registration need', () => {
    for (const f of ['dist', 'skills', 'mint.yaml', 'server.json', '.claude-plugin']) {
      expect(pkg.files).toContain(f);
    }
  });

  it('keeps the registry description within the 100-char schema limit', () => {
    expect(read('server.json').description.length).toBeLessThanOrEqual(100);
  });

  it('keeps the mcpb runtime floor on an LTS Node so LTS users can install', () => {
    expect(read('manifest.json').compatibility.runtimes.node).toBe('>=22.5.0');
  });

  it('carries no key the mcpb schema would reject', () => {
    const allowed = new Set([
      '$schema', 'manifest_version', 'name', 'display_name', 'version',
      'description', 'author', 'repository', 'homepage', 'support', 'license',
      'keywords', 'server', 'user_config', 'tools', 'compatibility', 'icon',
      'screenshots', 'long_description', 'documentation', 'privacy_policies',
      'tools_generated', 'prompts', 'prompts_generated',
    ]);
    const unknown = Object.keys(read('manifest.json')).filter((k) => !allowed.has(k));
    expect(unknown).toEqual([]);
  });

  it('names the same author in the .mcpb manifest as in the plugin manifest', () => {
    expect(read('manifest.json').author.name).toBe('Chris Hall');
    expect(read('manifest.json').author.name).toBe(read('.claude-plugin/plugin.json').author.name);
  });

  it('does not ship mint.yaml inside the .mcpb bundle', () => {
    const ignore = readFileSync(join(root, '.mcpbignore'), 'utf8');
    expect(ignore).toMatch(/^\*\.ya?ml\s*$/m);
  });
});

/**
 * The env keys the server honours. `TEAMS_WS_PORT` is this server's own
 * (.env.example / README); `FETCHPROXY_WS_HOST` and `FETCHPROXY_IDENTITY_DIR`
 * are read by the bundled `@fetchproxy/server` because this server never
 * passes `host` / `identityDir`. It always passes `port`, so
 * `FETCHPROXY_WS_PORT` is NOT honoured and is not declared.
 */
const ENV = ['FETCHPROXY_IDENTITY_DIR', 'FETCHPROXY_WS_HOST', 'TEAMS_WS_PORT'];

describe('manifest tool roster', () => {
  it('matches the registered tools in BOTH directions', async () => {
    const registered = (await servedTools()).map((t) => t.name).sort();
    const declared = read('manifest.json').tools.map((t: { name: string }) => t.name).sort();
    expect(declared).toEqual(registered);

    for (const t of read('manifest.json').tools) {
      expect(t.description, `${t.name} needs a description`).toBeTruthy();
    }
  });
});

describe('install surfaces pass every honoured env key', () => {
  it('manifest.json passes each one from an optional user_config entry', () => {
    const m = read('manifest.json');
    const env = (m.server.mcp_config.env ?? {}) as Record<string, string>;
    expect(Object.keys(env).sort()).toEqual(ENV);
    for (const [key, value] of Object.entries(env)) {
      const ref = /^\$\{user_config\.([^}]+)\}$/.exec(value)?.[1];
      expect(ref, key).toBeDefined();
      expect(m.user_config?.[ref!]?.required, key).toBe(false);
    }
  });

  it('server.json declares each one as optional', () => {
    const vars = read('server.json').packages[0].environmentVariables as { name: string; isRequired: boolean }[];
    expect(vars.map((v) => v.name).sort()).toEqual(ENV);
    expect(vars.every((v) => v.isRequired === false)).toBe(true);
  });
});
