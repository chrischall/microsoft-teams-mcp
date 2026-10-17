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
 * downloaded on another machine).
 *
 * The classifier is mcp-utils' `describeTransportFailure` (lifted from this
 * file): an `McpToolError` naming the HOST — never the path or query, which
 * may carry ids — the cause's code, and a hint keyed to that code. Only two
 * hints are Teams-specific, below.
 */
import { describeTransportFailure as describe, type TransportFailureHints } from '@chrischall/mcp-utils';
import type { McpToolError } from '@chrischall/mcp-utils';

export { transportCauseCode } from '@chrischall/mcp-utils';

const TEAMS_HINTS: TransportFailureHints = {
  // Name the hosts a split-tunnel VPN actually fails on for Teams.
  dns: (host) =>
    `This machine could not resolve ${host}. Check DNS and VPN: a split-tunnel VPN that resolves ` +
    'teams.cloud.microsoft but not the regional *.asm.skype.com and *.sharepoint.com hosts ' +
    `fails exactly like this. \`nslookup ${host}\` on the machine running this server tells you which.`,
  // "before Teams answered", whichever Teams host (image store, gateway, SharePoint) it was.
  connect: (host) =>
    `The connection to ${host} was refused, reset or timed out before Teams answered — usually a ` +
    "firewall, or a network that only lets traffic out through a proxy. Node's fetch ignores " +
    'proxy settings unless told: start this server with HTTPS_PROXY=<your proxy URL> and ' +
    'NODE_USE_ENV_PROXY=1 (Node 22.21+ / 24.5+), or allow the host through the firewall.',
};

/**
 * An `McpToolError` describing `err` when it is `fetch`'s transport failure
 * for `url`, else `undefined` (the caller rethrows `err` as it was).
 */
export function describeTransportFailure(err: unknown, url: string, service: string): McpToolError | undefined {
  return describe(err, url, service, { hints: TEAMS_HINTS });
}
