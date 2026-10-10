import { describe, it, expect } from 'vitest';
import { FetchproxyCapabilityUnavailableError } from '@fetchproxy/server';
import { McpToolError } from '@chrischall/mcp-utils';
import { wrapApiError, wrapBridgeError } from '../../src/tools/errors.js';

describe('wrapBridgeError', () => {
  it('points a pairing failure at ContextMint Bridge', () => {
    const err = wrapBridgeError(new Error('pair code 123456 awaiting approval'), 'read chats');
    expect(err.message).toMatch(/not yet approved/i);
    expect(JSON.stringify(err)).not.toMatch(/Transporter/i);
    expect(String((err as { hint?: string }).hint)).toMatch(/ContextMint Bridge/);
  });

  it('reports an unavailable capability as a browser limitation, not a pairing or MCP bug', () => {
    const raw = new FetchproxyCapabilityUnavailableError(
      'capability "read_dom_list" is not available in this browser (safari)',
      { capability: 'read_dom_list', platform: 'safari' },
    );
    const err = wrapBridgeError(raw, 'read chats');
    expect(err.message).toMatch(/this browser/i);
    expect(err.message).toMatch(/safari/i);
    expect(err.message).not.toMatch(/not yet approved/i);
    expect(String((err as { hint?: string }).hint)).toMatch(/Chrome/);
  });
});

describe('wrapApiError', () => {
  it('turns a token-capture failure into the open-a-tab-and-click-a-chat hint', () => {
    const err = wrapApiError(
      new Error('could not capture a Teams token from the browser (timed out after 30000ms). Open …'),
      'list chats',
    );
    expect(err.message).toMatch(/could not get a Teams access token to list chats/i);
    expect(String((err as { hint?: string }).hint)).toMatch(/teams\.cloud\.microsoft/);
    expect(String((err as { hint?: string }).hint)).toMatch(/click into any chat/i);
  });

  it('points a pairing failure during capture at ContextMint Bridge', () => {
    const err = wrapApiError(
      new Error('could not capture a Teams token from the browser (pair code 123456 awaiting approval). Open …'),
      'list chats',
    );
    expect(String((err as { hint?: string }).hint)).toMatch(/ContextMint Bridge/);
  });

  it('passes an McpToolError from the API client through untouched', () => {
    const original = new McpToolError('Teams chat service returned HTTP 503.', { status: 503 });
    expect(wrapApiError(original, 'list chats')).toBe(original);
  });

  it('wraps anything else with the action that failed', () => {
    const err = wrapApiError(new Error('socket hang up'), 'list channels');
    expect(err).toBeInstanceOf(McpToolError);
    expect(err.message).toBe('Failed to list channels: socket hang up');
  });
});
