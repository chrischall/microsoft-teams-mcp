import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { versionSyncTest, createTestHarness } from '@chrischall/mcp-utils/test';
import type { McpServer } from '@modelcontextprotocol/server';
import { registerChatTools } from '../src/tools/chat.js';
import { registerChannelTools } from '../src/tools/channels.js';
import { registerHealthcheckTool } from '../src/tools/healthcheck.js';
import type { TeamsClient } from '../src/client.js';

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

  it('does not ship mint.yaml inside the .mcpb bundle', () => {
    const ignore = readFileSync(join(root, '.mcpbignore'), 'utf8');
    expect(ignore).toMatch(/^\*\.ya?ml\s*$/m);
  });
});

describe('manifest tool roster', () => {
  it('matches the registered tools in BOTH directions', async () => {
    const client = {
      listChats: async () => [],
      getOpenChatMessages: async () => [],
      listTeamsAndChannels: async () => [],
      getOpenChannelPosts: async () => [],
      transport: { status: () => ({}), runProbe: async () => ({ ok: true, elapsed_ms: 0, bridge: {} }) },
    } as unknown as TeamsClient;

    const h = await createTestHarness((server: McpServer) => {
      registerChatTools(server, client);
      registerChannelTools(server, client);
      registerHealthcheckTool(server, client);
    });
    const registered = (await h.listTools()).map((t) => t.name).sort();
    await h.close();

    const declared = read('manifest.json').tools.map((t: { name: string }) => t.name).sort();
    expect(declared).toEqual(registered);

    for (const t of read('manifest.json').tools) {
      expect(t.description, `${t.name} needs a description`).toBeTruthy();
    }
  });
});
