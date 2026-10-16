import { describe, expect, it } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import { describeTransportFailure, transportCauseCode } from '../src/network-error.js';

/** What Node's fetch (undici) throws for a request that never got a response. */
function fetchFailed(cause: unknown): TypeError {
  return new TypeError('fetch failed', { cause });
}

function sysError(code: string, message: string): Error {
  return Object.assign(new Error(message), { code });
}

const IMG = 'https://in-api.asm.skype.com/v1/objects/0-sin-d3-abc/views/imgo?secret=1';

describe('transportCauseCode', () => {
  it('reads the code off the cause', () => {
    expect(transportCauseCode(fetchFailed(sysError('ENOTFOUND', 'getaddrinfo ENOTFOUND x')))).toBe('ENOTFOUND');
  });

  it('reads the first code inside an AggregateError (both address families refused)', () => {
    const agg = new AggregateError([sysError('ECONNREFUSED', 'a'), sysError('ECONNREFUSED', 'b')], '');
    expect(transportCauseCode(fetchFailed(agg))).toBe('ECONNREFUSED');
  });

  it('follows a nested cause chain', () => {
    const inner = sysError('UNABLE_TO_GET_ISSUER_CERT_LOCALLY', 'unable to get local issuer certificate');
    expect(transportCauseCode(fetchFailed(new Error('wrapped', { cause: inner })))).toBe(
      'UNABLE_TO_GET_ISSUER_CERT_LOCALLY',
    );
  });

  it('is undefined when no code is anywhere in the chain', () => {
    expect(transportCauseCode(fetchFailed(new Error('other side closed')))).toBeUndefined();
  });
});

describe('describeTransportFailure', () => {
  it('names the host and the code, never the path or query', () => {
    const err = describeTransportFailure(
      fetchFailed(sysError('ENOTFOUND', 'getaddrinfo ENOTFOUND in-api.asm.skype.com')),
      IMG,
      'Teams image store',
    );
    expect(err).toBeInstanceOf(McpToolError);
    expect(err!.message).toContain('in-api.asm.skype.com');
    expect(err!.message).toContain('ENOTFOUND');
    expect(err!.message).not.toContain('/v1/objects');
    expect(err!.message).not.toContain('secret');
    expect(err!.kind).toBe('transport');
    expect(err!.hint).toMatch(/DNS|VPN/);
  });

  it.each([
    ['ECONNREFUSED', /proxy|firewall/i],
    ['ECONNRESET', /proxy|firewall/i],
    ['ETIMEDOUT', /proxy|firewall/i],
    ['UND_ERR_CONNECT_TIMEOUT', /proxy|firewall/i],
    ['ENETUNREACH', /proxy|firewall/i],
    ['EAI_AGAIN', /DNS|VPN/],
    ['UNABLE_TO_GET_ISSUER_CERT_LOCALLY', /NODE_EXTRA_CA_CERTS/],
    ['SELF_SIGNED_CERT_IN_CHAIN', /NODE_EXTRA_CA_CERTS/],
    ['UNABLE_TO_VERIFY_LEAF_SIGNATURE', /NODE_EXTRA_CA_CERTS/],
    ['CERT_HAS_EXPIRED', /NODE_EXTRA_CA_CERTS|clock/],
    ['ERR_TLS_CERT_ALTNAME_INVALID', /NODE_EXTRA_CA_CERTS|intercept/i],
  ])('gives an actionable hint for %s', (code, hint) => {
    const err = describeTransportFailure(fetchFailed(sysError(code, `x ${code}`)), IMG, 'Teams image store');
    expect(err!.message).toContain(code);
    expect(err!.hint).toMatch(hint);
  });

  it('points a refused connection at the proxy settings Node actually reads', () => {
    const err = describeTransportFailure(fetchFailed(sysError('ECONNREFUSED', 'connect ECONNREFUSED')), IMG, 'x');
    expect(err!.hint).toContain('NODE_USE_ENV_PROXY=1');
    expect(err!.hint).toContain('HTTPS_PROXY');
  });

  it('still reports host and cause message when the code is unknown', () => {
    const err = describeTransportFailure(fetchFailed(new Error('other side closed')), IMG, 'Teams image store');
    expect(err!.message).toContain('in-api.asm.skype.com');
    expect(err!.message).toContain('other side closed');
  });

  it('leaves anything that is not a transport failure to the caller', () => {
    expect(describeTransportFailure(new McpToolError('HTTP 404'), IMG, 'x')).toBeUndefined();
    expect(describeTransportFailure(new Error('boom'), IMG, 'x')).toBeUndefined();
    expect(describeTransportFailure(new TypeError('Invalid URL'), IMG, 'x')).toBeUndefined();
  });
});

describe('describeTransportFailure — an egress proxy that refused the tunnel', () => {
  /** What Node's fetch throws when HTTPS_PROXY answers the CONNECT with a non-200 (captured on Node 26). */
  function proxyRefused(status: number): TypeError {
    const aborted = Object.assign(new Error(`Proxy response (${status}) !== 200 when HTTP Tunneling`), {
      name: 'RequestAbortedError',
      code: 'UND_ERR_ABORTED',
    });
    return fetchFailed(Object.assign(new Error('Request was cancelled.', { cause: aborted }), { code: 0 }));
  }

  const FILE = 'https://contoso.sharepoint.com/_api/v2.1/shares/u!abc/driveItem';

  it('says the network policy refused the host, not "check DNS/VPN"', () => {
    const err = describeTransportFailure(proxyRefused(403), FILE, 'SharePoint');
    expect(err).toBeInstanceOf(McpToolError);
    expect(err!.message).toContain('contoso.sharepoint.com');
    expect(err!.message).not.toContain('/_api');
    expect(err!.hint).toMatch(/egress/i);
    expect(err!.hint).toMatch(/allow/i);
    expect(err!.hint).not.toMatch(/VPN|DNS|nslookup|curl/);
  });

  it('treats Node http\'s ERR_PROXY_TUNNEL the same way', () => {
    const err = describeTransportFailure(
      fetchFailed(Object.assign(new Error('Failed to establish tunnel: proxy returned 403'), { code: 'ERR_PROXY_TUNNEL' })),
      FILE,
      'SharePoint',
    );
    expect(err!.hint).toMatch(/egress/i);
    expect(err!.hint).not.toMatch(/VPN|DNS/);
  });

  it('a proxy that answered 502 (could not reach the host) is not called a policy refusal', () => {
    const err = describeTransportFailure(proxyRefused(502), FILE, 'SharePoint');
    expect(err!.hint).not.toMatch(/egress allow/i);
    expect(err!.hint).toMatch(/proxy/i);
  });
});
