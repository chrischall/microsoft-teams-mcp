import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
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

  /**
   * Claude Code reads a plugin's MCP config from `mcpServers` (an unknown key
   * such as `mcp` is silently ignored) and resolves the path against the
   * PLUGIN ROOT -- the folder holding .claude-plugin/ -- not .claude-plugin/.
   * The root .mcp.json always loads too, with a relative path meant for
   * project-scoped launches; the plugin config must reuse its server name so
   * it replaces that entry in a plugin install.
   */
  it('points the plugin at a config that resolves under a plugin install', () => {
    const plugin = read('.claude-plugin/plugin.json');
    expect(plugin).not.toHaveProperty('mcp');
    expect(typeof plugin.mcpServers).toBe('string');

    const rel = (plugin.mcpServers as string).replace(/^\.\//, '');
    expect(existsSync(join(root, rel)), `${rel} must exist under the plugin root`).toBe(true);

    const cfg = read(rel);
    const names = Object.keys(cfg.mcpServers);
    expect(names).toEqual(Object.keys(read('.mcp.json').mcpServers));
    for (const name of names) {
      const args = cfg.mcpServers[name].args.join(' ');
      expect(args).toContain('${CLAUDE_PLUGIN_ROOT}/dist/bundle.js');
    }
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
 * The env keys the server honours. The `TEAMS_*` keys are this server's own
 * (.env.example / README): the bridge port, and the chat-service token's
 * source, cache, capture deadline and API base. `FETCHPROXY_WS_HOST` and
 * `FETCHPROXY_IDENTITY_DIR` are read by the bundled `@fetchproxy/server`
 * because this server never passes `host` / `identityDir`. It always passes
 * `port`, so `FETCHPROXY_WS_PORT` is NOT honoured and is not declared. The
 * `MCP_CONFIRM_*` keys are mcp-utils' `confirmWrite` settings for the two
 * write tools, as the README's confirmation table documents them.
 */
const ENV = [
  'FETCHPROXY_IDENTITY_DIR',
  'FETCHPROXY_WS_HOST',
  'MCP_CONFIRM_ELICITATION',
  'MCP_CONFIRM_MODE',
  'MCP_CONFIRM_SECRET',
  'MCP_CONFIRM_TTL_SECONDS',
  'TEAMS_ACCESS_TOKEN',
  'TEAMS_CAPTURE_TIMEOUT',
  'TEAMS_CHATSVC_BASE',
  'TEAMS_CHATSVC_REGION',
  'TEAMS_DISABLE_FETCHPROXY',
  'TEAMS_TIME_ZONE',
  'TEAMS_TOKEN_CACHE',
  'TEAMS_TOKEN_FILE',
  'TEAMS_WS_PORT',
];

/**
 * Chats and channels are read through the chat-service API now, by id or
 * name, so no install surface may still promise only "the open" conversation.
 */
describe('install surface descriptions', () => {
  it('no longer describe reads as limited to the currently open conversation', () => {
    const texts = [
      read('manifest.json').description,
      read('server.json').description,
      read('.claude-plugin/plugin.json').description,
      read('.claude-plugin/marketplace.json').metadata.description,
      read('.claude-plugin/marketplace.json').plugins[0].description,
      ...read('manifest.json').tools.map((t: { description: string }) => t.description),
    ];
    for (const t of texts) expect(t).not.toMatch(/open (messages|chat|channel)|currently open/i);
  });
});

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

/**
 * teams_send_message and teams_react post as the user, visibly to everyone in
 * the conversation, so every surface that describes them must say they ask
 * first — and nothing may still claim the server only reads.
 */
describe('write tools are documented as confirm-gated', () => {
  const readme = readFileSync(join(root, 'README.md'), 'utf8');
  const skill = readFileSync(join(root, 'skills/teams-fpx/SKILL.md'), 'utf8');

  it('says in each write tool\'s manifest description that it asks to confirm first', () => {
    const tools = read('manifest.json').tools as { name: string; description: string }[];
    for (const name of ['teams_send_message', 'teams_react']) {
      expect(tools.find((t) => t.name === name)?.description, name).toMatch(/confirm/i);
    }
  });

  it('README describes both writes, the confirmation flow and the self-chat to test in', () => {
    for (const s of ['teams_send_message', 'teams_react', 'confirmToken', 'MCP_CONFIRM_MODE', '48:notes']) {
      expect(readme, s).toContain(s);
    }
    expect(readme).not.toMatch(/all tools are read-only/i);
  });

  it('the skill points writes at the confirm-gated tools and demands approval for a shell send', () => {
    expect(skill).toContain('teams_send_message');
    expect(skill).toContain('teams_react');
    expect(skill).toMatch(/explicit approval/i);
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

  it('marks the access token and the confirm signing key secret on both surfaces, and nothing else', () => {
    const SECRET = ['MCP_CONFIRM_SECRET', 'TEAMS_ACCESS_TOKEN'];
    const vars = read('server.json').packages[0].environmentVariables as { name: string; isSecret: boolean }[];
    expect(vars.filter((v) => v.isSecret).map((v) => v.name).sort()).toEqual(SECRET);

    const m = read('manifest.json');
    const refs = SECRET.map((k) => /^\$\{user_config\.([^}]+)\}$/.exec(m.server.mcp_config.env[k])![1]).sort();
    const sensitive = Object.entries(m.user_config as Record<string, { sensitive?: boolean }>)
      .filter(([, c]) => c.sensitive === true)
      .map(([k]) => k)
      .sort();
    expect(sensitive).toEqual(refs);
  });

  it('documents in the README every MCP_CONFIRM_* key it passes', () => {
    const readme = readFileSync(join(root, 'README.md'), 'utf8');
    for (const k of ENV.filter((e) => e.startsWith('MCP_CONFIRM_'))) expect(readme, k).toContain(`\`${k}\``);
  });
});
