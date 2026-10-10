#!/usr/bin/env node
import { runMcp, loadDotenvSafely } from '@chrischall/mcp-utils';
import { VERSION } from './version.js';
import { TeamsClient } from './client.js';
import { ChatServiceClient } from './chatsvc.js';
import type { TeamsDeps } from './deps.js';
import { registerChatTools } from './tools/chat.js';
import { registerChannelTools } from './tools/channels.js';
import { registerConversationTools } from './tools/conversation.js';
import { registerUnreadTools } from './tools/unread.js';
import { registerAttachmentTools } from './tools/attachment.js';
import { registerActivityTools } from './tools/activity.js';
import { registerHealthcheckTool } from './tools/healthcheck.js';

loadDotenvSafely();

/**
 * Built HERE, in the caller, not inside a registrar — see office-outlook-mcp
 * for the full rationale (`runMcp` builds a server instance per served
 * connection, so anything built inside a registrar would be rebuilt). For
 * this MCP it matters even more: a second `TeamsClient` would mean a second
 * fetchproxy identity/session for the same process, and a second
 * `ChatServiceClient` a second token capture.
 */
const deps: TeamsDeps = {
  teams: new TeamsClient(),
  chatsvc: new ChatServiceClient(),
};

await runMcp({
  name: 'microsoft-teams-mcp',
  version: VERSION,
  banner:
    '[microsoft-teams-mcp] This project was developed and is maintained by AI. Use at your own discretion.',
  deps,
  tools: [
    registerChatTools,
    registerChannelTools,
    registerConversationTools,
    registerUnreadTools,
    registerAttachmentTools,
    registerActivityTools,
    registerHealthcheckTool,
  ],
});
