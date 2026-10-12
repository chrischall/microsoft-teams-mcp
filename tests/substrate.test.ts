import { describe, expect, it, vi } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import {
  SubstrateClient,
  SUBSTRATE_QUERY_URL,
  buildQueryBody,
  parseQueryResponse,
  CAPTURE_RETRY_COOLDOWN_MS,
} from '../src/substrate.js';

/** A JWT carrying `exp`; opaque filler, never a real token. */
function jwt(secondsFromNow = 3000): string {
  const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'none' })}.${b64({ exp: Math.floor(Date.now() / 1000) + secondsFromNow })}.x`;
}

/** Cache off, so no test touches the developer's real `~/.microsoft-teams-mcp`. */
const ENV = { TEAMS_TOKEN_CACHE: 'false' } as NodeJS.ProcessEnv;

type FetchMock = ReturnType<typeof vi.fn<(input: RequestInfo | URL, init?: RequestInit) => Promise<Response>>>;

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

const CHAT = '19:group0001@thread.v2';
const CHANNEL = '19:chan0001@thread.tacv2';

/** One Substrate result, in the shape the client parses (invented values only). */
function result(source: Record<string, unknown>, summary = 'the <c0>deploy</c0> is done'): Record<string, unknown> {
  return { Id: 'opaque-id', ReferenceId: 'opaque-ref', HitHighlightedSummary: summary, Source: source };
}

function responseOf(results: unknown[], extra: Record<string, unknown> = {}): unknown {
  return { EntitySets: [{ ResultSets: [{ Total: results.length, MoreResultsAvailable: false, Results: results, ...extra }] }] };
}

function client(fetchImpl: FetchMock, capture: () => Promise<string> = async () => jwt(), env = ENV) {
  return new SubstrateClient({ env, captureToken: capture, fetchImpl: fetchImpl as unknown as typeof fetch });
}

describe('buildQueryBody', () => {
  it('asks for Teams messages, newest first, one page at from/size', () => {
    const body = buildQueryBody({ query: 'deploy plan', from: 50, size: 25 }) as {
      entityRequests: Record<string, unknown>[];
    };
    expect(body.entityRequests).toHaveLength(1);
    expect(body.entityRequests[0]).toMatchObject({
      entityType: 'Message',
      contentSources: ['Teams'],
      query: { queryString: 'deploy plan' },
      from: 50,
      size: 25,
      sort: [{ field: 'DateTime', sortDirection: 'Desc' }],
    });
  });
});

describe('parseQueryResponse', () => {
  it('maps a chat hit: conversation, message id, time, sender, body', () => {
    const page = parseQueryResponse(
      responseOf([
        result({
          ClientConversationId: CHAT,
          ClientMessageId: '1760000000000',
          DateTimeSent: '2025-10-09T08:53:20Z',
          DateTimeReceived: '2025-10-09T08:53:21Z',
          From: { EmailAddress: { Name: 'Ada Example', Address: 'ada@example.test' } },
          Preview: 'The deploy is done',
        }),
      ]),
    );
    expect(page.results).toEqual([
      {
        conversationId: CHAT,
        messageId: '1760000000000',
        time: '2025-10-09T08:53:20.000Z',
        sender: 'Ada Example',
        senderEmail: 'ada@example.test',
        html: 'The deploy is done',
        excerpt: true,
      },
    ]);
    expect(page.moreAvailable).toBe(false);
    expect(page.total).toBe(1);
  });

  it('prefers the full body over the preview, and the preview over the highlighted summary', () => {
    const [withBody, withSummary] = parseQueryResponse(
      responseOf([
        result({ ClientConversationId: CHAT, Body: { Content: '<p>full <b>body</b></p>' }, Preview: 'full body' }),
        result({ ClientConversationId: CHAT }, 'only <c0>summary</c0>'),
      ]),
    ).results;
    expect(withBody).toMatchObject({ html: '<p>full <b>body</b></p>' });
    expect(withBody).not.toHaveProperty('excerpt');
    expect(withSummary).toMatchObject({ html: 'only <c0>summary</c0>', excerpt: true });
  });

  it("reads a channel reply's root off the conversation id's ;messageid= suffix", () => {
    const [reply, root] = parseQueryResponse(
      responseOf([
        result({ ClientConversationId: `${CHANNEL};messageid=1759000000000`, ClientMessageId: '1759000000500', Subject: 'Release' }),
        result({ ClientConversationId: `${CHANNEL};messageid=1759000000000`, ClientMessageId: '1759000000000' }),
      ]),
    ).results;
    expect(reply).toMatchObject({ conversationId: CHANNEL, replyTo: '1759000000000', subject: 'Release' });
    expect(root.conversationId).toBe(CHANNEL);
    expect(root).not.toHaveProperty('replyTo');
  });

  it('carries the sender MRI when the hit has one, so fromMe can be decided', () => {
    const [hit] = parseQueryResponse(
      responseOf([
        result({
          ClientConversationId: CHAT,
          Extensions: { SkypeSpaces_ConversationPost_Extension_FromSkypeInternalId: '8:orgid:abc' },
        }),
      ]),
    ).results;
    expect(hit.senderMri).toBe('8:orgid:abc');
  });

  it('counts a hit with no Teams conversation id as unmapped rather than inventing one', () => {
    const page = parseQueryResponse(responseOf([result({ Subject: 'an email' }), result({ ClientConversationId: CHAT })]));
    expect(page.results).toHaveLength(1);
    expect(page.unmapped).toBe(1);
  });

  it('reads MoreResultsAvailable and Total', () => {
    const page = parseQueryResponse(responseOf([], { Total: 340, MoreResultsAvailable: true }));
    expect(page).toMatchObject({ total: 340, moreAvailable: true, results: [] });
  });

  it('treats a response with no result sets as an empty page', () => {
    expect(parseQueryResponse({})).toEqual({ results: [], moreAvailable: false, unmapped: 0 });
  });
});

describe('SubstrateClient.searchMessages', () => {
  it('POSTs the query to the search service with the captured bearer, never following a redirect', async () => {
    const token = jwt();
    const fetchImpl: FetchMock = vi.fn(async () => json(responseOf([])));
    await client(fetchImpl, async () => token).searchMessages({ query: 'deploy', from: 0, size: 25 });

    expect(fetchImpl).toHaveBeenCalledTimes(1);
    const [url, init] = fetchImpl.mock.calls[0];
    expect(String(url)).toBe(SUBSTRATE_QUERY_URL);
    expect(SUBSTRATE_QUERY_URL).toBe('https://substrate.office.com/searchservice/api/v2/query');
    expect(init?.method).toBe('POST');
    expect(init?.redirect).toBe('error');
    const headers = new Headers(init?.headers);
    expect(headers.get('authorization')).toBe(`Bearer ${token}`);
    expect(headers.get('content-type')).toBe('application/json');
    expect(JSON.parse(String(init?.body))).toMatchObject({ entityRequests: [{ query: { queryString: 'deploy' } }] });
  });

  it('re-captures once and replays on a 401', async () => {
    const capture = vi.fn(async () => jwt());
    const fetchImpl: FetchMock = vi.fn();
    fetchImpl.mockResolvedValueOnce(new Response('', { status: 401 }));
    fetchImpl.mockResolvedValueOnce(json(responseOf([])));
    await client(fetchImpl, capture).searchMessages({ query: 'deploy', from: 0, size: 25 });
    expect(capture).toHaveBeenCalledTimes(2);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('maps a non-2xx to an McpToolError with its status and no body', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => new Response('secret-ish body', { status: 500 }));
    const err = (await client(fetchImpl).searchMessages({ query: 'x', from: 0, size: 25 }).catch((e) => e)) as McpToolError;
    expect(err).toBeInstanceOf(McpToolError);
    expect(err.status).toBe(500);
    expect(err.message).not.toContain('secret-ish');
  });

  it('is not configured, and never captures, with the bridge disabled', async () => {
    const capture = vi.fn(async () => jwt());
    const c = client(vi.fn(), capture, { ...ENV, TEAMS_DISABLE_FETCHPROXY: 'true' });
    expect(c.isConfigured()).toBe(false);
    await expect(c.searchMessages({ query: 'x', from: 0, size: 25 })).rejects.toMatchObject({ kind: 'no_credential' });
    expect(capture).not.toHaveBeenCalled();
  });

  it('after a failed capture, skips capturing again for a cooldown unless forced', async () => {
    vi.useFakeTimers({ toFake: ['Date'] });
    try {
      const capture = vi.fn(async (): Promise<string> => {
        throw new Error('could not capture a Substrate search token from the browser (timed out)');
      });
      const c = client(vi.fn(), capture);
      await expect(c.searchMessages({ query: 'x', from: 0, size: 25 })).rejects.toThrow(/timed out/);
      expect(capture).toHaveBeenCalledTimes(1);

      // Inside the cooldown: no second capture (and no second wait on the bridge).
      await expect(c.searchMessages({ query: 'x', from: 0, size: 25 })).rejects.toThrow(/timed out/);
      expect(capture).toHaveBeenCalledTimes(1);

      // Forced (mode "substrate"): it tries again.
      await expect(c.searchMessages({ query: 'x', from: 0, size: 25 }, { force: true })).rejects.toThrow();
      expect(capture).toHaveBeenCalledTimes(2);

      // After the cooldown, an unforced call tries again too.
      vi.setSystemTime(Date.now() + CAPTURE_RETRY_COOLDOWN_MS + 1);
      await expect(c.searchMessages({ query: 'x', from: 0, size: 25 })).rejects.toThrow();
      expect(capture).toHaveBeenCalledTimes(3);
    } finally {
      vi.useRealTimers();
    }
  });
});
