#!/usr/bin/env node
import { runMcp, loadDotenvSafely } from '@chrischall/mcp-utils';
import { VERSION } from './version.js';
import { TeamsClient } from './client.js';
import { registerChatTools } from './tools/chat.js';
import { registerHealthcheckTool } from './tools/healthcheck.js';

loadDotenvSafely();

/**
 * Built HERE, in the caller, not inside a registrar — see office-outlook-mcp
 * for the full rationale (`runMcp` builds a server instance per served
 * connection, so anything built inside a registrar would be rebuilt). For
 * this MCP it matters even more: a second `TeamsClient` would mean a second
 * fetchproxy identity/session for the same process.
 */
const client = new TeamsClient();

await runMcp({
  name: 'microsoft-teams-mcp',
  version: VERSION,
  banner:
    '[microsoft-teams-mcp] This project was developed and is maintained by AI. Use at your own discretion.',
  deps: client,
  tools: [registerChatTools, registerHealthcheckTool],
});
