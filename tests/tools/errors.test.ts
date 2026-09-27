import { describe, it, expect } from 'vitest';
import { FetchproxyCapabilityUnavailableError } from '@fetchproxy/server';
import { wrapBridgeError } from '../../src/tools/errors.js';

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
