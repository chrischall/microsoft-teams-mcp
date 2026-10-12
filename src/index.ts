#!/usr/bin/env node
import { runMcp, loadDotenvSafely } from '@chrischall/mcp-utils';
import { VERSION } from './version.js';
import { ChatServiceClient } from './chatsvc.js';
import { SharePointClient } from './sharepoint.js';
import type { TeamsDeps } from './deps.js';
import { registerChatTools } from './tools/chat.js';
import { registerChannelTools } from './tools/channels.js';
import { registerConversationTools } from './tools/conversation.js';
import { registerUnreadTools } from './tools/unread.js';
import { registerChangesTools } from './tools/changes.js';
import { registerAttachmentTools } from './tools/attachment.js';
import { registerSendTools } from './tools/send.js';
import { registerReactTools } from './tools/react.js';
import { registerActivityTools } from './tools/activity.js';
import { registerHealthcheckTool } from './tools/healthcheck.js';

loadDotenvSafely();

/**
 * Built HERE, in the caller, not inside a registrar — see office-outlook-mcp
 * for the full rationale (`runMcp` builds a server instance per served
 * connection, so anything built inside a registrar would be rebuilt). For
 * this MCP a second `ChatServiceClient` would mean a second token capture.
 * `SharePointClient` holds its own per-host token managers, so the same reason
 * applies — build it once, here, beside the chat-service client.
 */
const deps: TeamsDeps = {
  chatsvc: new ChatServiceClient(),
  sharepoint: new SharePointClient(),
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
    registerChangesTools,
    registerAttachmentTools,
    registerSendTools,
    registerReactTools,
    registerActivityTools,
    registerHealthcheckTool,
  ],
});
