/**
 * A request that never got a response, said in a way someone can act on.
 *
 * Node's `fetch` (undici) reports every transport failure — DNS, a refused or
 * reset connection, a TLS certificate it will not trust — as the same
 * `TypeError('fetch failed')`, with the real reason on its `cause` chain.
 * Surfaced as is, that reads "Failed to get the attachment: fetch failed",
 * which names neither the host nor the reason (2026-10-12: two inline images
 * on `in-api`/`us-api.asm.skype.com` failed that way on a machine whose
 * chat-service calls to `teams.cloud.microsoft` were fine, while the same refs
 * downloaded on another machine). This turns it into an `McpToolError` naming
 * the HOST — never the path or query, which may carry ids — the cause's code,
 * and a hint keyed to that code.
 */
import { McpToolError } from '@chrischall/mcp-utils';

/** How far down a cause chain (and into an AggregateError) to look for a code. */
const MAX_CAUSE_DEPTH = 5;

const DNS_CODES = new Set(['ENOTFOUND', 'EAI_AGAIN', 'EAI_FAIL', 'EAI_NONAME']);

const CONNECT_CODES = new Set([
  'ECONNREFUSED',
  'ECONNRESET',
  'ETIMEDOUT',
  'ENETUNREACH',
  'EHOSTUNREACH',
  'EPIPE',
  'UND_ERR_CONNECT_TIMEOUT',
  'UND_ERR_SOCKET',
  'UND_ERR_HEADERS_TIMEOUT',
]);

const CERT_CODES = new Set([
  'UNABLE_TO_GET_ISSUER_CERT_LOCALLY',
  'UNABLE_TO_GET_ISSUER_CERT',
  'UNABLE_TO_VERIFY_LEAF_SIGNATURE',
  'SELF_SIGNED_CERT_IN_CHAIN',
  'DEPTH_ZERO_SELF_SIGNED_CERT',
  'CERT_HAS_EXPIRED',
  'CERT_NOT_YET_VALID',
  'ERR_TLS_CERT_ALTNAME_INVALID',
]);

function codeOf(e: unknown): string | undefined {
  if (e === null || typeof e !== 'object') return undefined;
  const code = (e as { code?: unknown }).code;
  return typeof code === 'string' && code !== '' ? code : undefined;
}

/** The causes to look at, nearest first: the cause chain, and an AggregateError's members. */
function causes(err: unknown): unknown[] {
  const out: unknown[] = [];
  let at: unknown = err;
  for (let depth = 0; depth < MAX_CAUSE_DEPTH && at !== undefined && at !== null; depth++) {
    out.push(at);
    if (at instanceof AggregateError) out.push(...at.errors);
    at = (at as { cause?: unknown }).cause;
  }
  return out;
}

/** The first system/undici error code on `err`'s cause chain. */
export function transportCauseCode(err: unknown): string | undefined {
  for (const e of causes(err)) {
    const code = codeOf(e);
    if (code !== undefined) return code;
  }
  return undefined;
}

/** The innermost cause's message, which is where undici puts the real reason. */
function causeMessage(err: unknown): string | undefined {
  const chain = causes(err).slice(1);
  for (let i = chain.length - 1; i >= 0; i--) {
    const e = chain[i];
    if (e instanceof Error && e.message !== '') return e.message;
  }
  return undefined;
}

/** Whether `err` is `fetch`'s "never got a response" failure. */
function isFetchFailed(err: unknown): boolean {
  return err instanceof TypeError && err.message === 'fetch failed';
}

/**
 * The status a configured HTTPS proxy answered the CONNECT with, when that is
 * why the request failed. undici reports it only in the message of an
 * `UND_ERR_ABORTED` cause ("Proxy response (403) !== 200 when HTTP
 * Tunneling") — there is no structured field to branch on — and Node's own
 * http proxy support uses code `ERR_PROXY_TUNNEL`.
 */
function proxyRefusalStatus(err: unknown): number | 'unknown' | undefined {
  for (const e of causes(err)) {
    if (codeOf(e) === 'ERR_PROXY_TUNNEL') {
      const m = /\b(\d{3})\b/.exec(e instanceof Error ? e.message : '');
      return m ? Number(m[1]) : 'unknown';
    }
    if (e instanceof Error) {
      const m = /^Proxy response \((\d{3})\) !== 200 when HTTP Tunneling/.exec(e.message);
      if (m) return Number(m[1]);
    }
  }
  return undefined;
}

/**
 * A proxy that answered 403/407 refused the destination by policy: on mcp-host
 * that is the registration's egress allow-list (mint.yaml `egress.allow`), not
 * this machine's DNS or VPN, so say that instead of the network-route hints.
 */
function proxyRefusalHint(status: number | 'unknown', host: string): string | undefined {
  if (status !== 403 && status !== 407 && status !== 'unknown') return undefined;
  return (
    `The outbound proxy this server runs behind refused the connection to ${host}: its egress ` +
    'policy does not allow that host. On a hosted (mcp-host) registration, the host must be on the ' +
    "registration's egress allow-list — the owner accepts the manifest's egress proposal (or adds " +
    'the host) on the registration, then retries. Elsewhere, ask whoever runs the proxy to allow it.'
  );
}

function hintFor(code: string | undefined, host: string, err?: unknown): string {
  const refused = proxyRefusalStatus(err);
  if (refused !== undefined) {
    const policy = proxyRefusalHint(refused, host);
    if (policy !== undefined) return policy;
    return (
      `The outbound proxy this server runs behind could not open a connection to ${host} ` +
      `(it answered ${refused}). Retry; if it persists, check the proxy can reach the host.`
    );
  }
  if (code !== undefined && DNS_CODES.has(code)) {
    return (
      `This machine could not resolve ${host}. Check DNS and VPN: a split-tunnel VPN that resolves ` +
      'teams.cloud.microsoft but not the regional *.asm.skype.com and *.sharepoint.com hosts ' +
      `fails exactly like this. \`nslookup ${host}\` on the machine running this server tells you which.`
    );
  }
  if (code !== undefined && CERT_CODES.has(code)) {
    return (
      `Node did not trust the TLS certificate ${host} presented. On a corporate network this is ` +
      'usually TLS inspection: export your proxy\'s root CA as PEM and start this server with ' +
      'NODE_EXTRA_CA_CERTS=/path/to/ca.pem (Node does not read the system keychain). If the ' +
      'certificate is expired or not yet valid, also check the machine\'s clock.'
    );
  }
  if (code !== undefined && CONNECT_CODES.has(code)) {
    return (
      `The connection to ${host} was refused, reset or timed out before Teams answered — usually a ` +
      'firewall, or a network that only lets traffic out through a proxy. Node\'s fetch ignores ' +
      'proxy settings unless told: start this server with HTTPS_PROXY=<your proxy URL> and ' +
      'NODE_USE_ENV_PROXY=1 (Node 22.21+ / 24.5+), or allow the host through the firewall.'
    );
  }
  return (
    `No response from ${host}. Check this machine can reach it (VPN, proxy, firewall); ` +
    `\`curl -sI https://${host}/\` from the same machine is a quick test.`
  );
}

/**
 * An `McpToolError` describing `err` when it is `fetch`'s transport failure
 * for `url`, else `undefined` (the caller rethrows `err` as it was).
 */
export function describeTransportFailure(err: unknown, url: string, service: string): McpToolError | undefined {
  if (!isFetchFailed(err)) return undefined;
  let host: string;
  try {
    host = new URL(url).hostname;
  } catch {
    host = '(unknown host)';
  }
  const code = transportCauseCode(err);
  const detail = [code, causeMessage(err)].filter((s): s is string => s !== undefined);
  const why = detail.length > 0 ? ` (${[...new Set(detail)].join(': ')})` : '';
  return new McpToolError(`Could not reach ${service} at ${host}${why}.`, {
    kind: 'transport',
    hint: hintFor(code, host, err),
    cause: err,
  });
}
