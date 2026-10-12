import { readFileSync } from 'node:fs';
import { describe, expect, it, vi } from 'vitest';
import { McpToolError } from '@chrischall/mcp-utils';
import {
  SubstrateClient,
  SUBSTRATE_QUERY_URL,
  buildQueryBody,
  buildQueryString,
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

/** A real response (2026-10-12), redacted: names, addresses, text and ids are placeholders. */
const FIXTURE = JSON.parse(
  readFileSync(new URL('./fixtures/substrate-query-response.json', import.meta.url), 'utf8'),
) as { EntitySets: { ResultSets: { Results: { Source: Record<string, unknown> }[] }[] }[] };

const CHAT = '19:group0001@thread.v2';

function responseOf(results: unknown[], extra: Record<string, unknown> = {}): unknown {
  return { EntitySets: [{ ResultSets: [{ Total: results.length, MoreResultsAvailable: false, Results: results, ...extra }] }] };
}

function client(fetchImpl: FetchMock, capture: () => Promise<string> = async () => jwt(), env = ENV) {
  return new SubstrateClient({ env, captureToken: capture, fetchImpl: fetchImpl as unknown as typeof fetch });
}

describe('buildQueryBody', () => {
  it('asks for Teams messages, newest first by Time, one page at from/size, with the required cvid and scenario', () => {
    const body = buildQueryBody({ query: 'deploy plan', from: 50, size: 25 }) as {
      entityRequests: Record<string, unknown>[];
      cvid: string;
      logicalId: string;
      scenario: { Name: string };
    };
    expect(body.entityRequests).toHaveLength(1);
    expect(body.entityRequests[0]).toEqual({
      entityType: 'Message',
      contentSources: ['Teams'],
      query: { queryString: 'deploy plan', displayQueryString: 'deploy plan' },
      from: 50,
      size: 25,
      sort: [{ Field: 'Time', SortDirection: 'Desc' }],
    });
    expect(body.cvid).toMatch(/^[0-9a-f-]{36}$/);
    expect(body.logicalId).toBe(body.cvid);
    expect(body.scenario.Name).toBe('powerbar');
  });

  it('shows the user their own query, not the filtered one', () => {
    const body = buildQueryBody({ query: 'deploy', from: 0, size: 5, conversationId: CHAT }) as {
      entityRequests: { query: { queryString: string; displayQueryString: string } }[];
    };
    expect(body.entityRequests[0].query.displayQueryString).toBe('deploy');
    expect(body.entityRequests[0].query.queryString).toContain('ClientThreadId:');
  });
});

describe('buildQueryString', () => {
  it('passes an unrestricted query through as typed', () => {
    expect(buildQueryString({ query: 'deploy OR "release plan"' })).toBe('deploy OR "release plan"');
  });

  it('restricts to one conversation with a quoted ClientThreadId, parenthesising the query', () => {
    expect(buildQueryString({ query: 'deploy OR ship', conversationId: '48:notes' })).toBe('(deploy OR ship) ClientThreadId:"48:notes"');
  });

  it('restricts by time with an UNQUOTED sent>= in whole seconds, rounded down', () => {
    expect(buildQueryString({ query: 'deploy', sentAfter: Date.parse('2026-10-12T13:40:00.900Z') })).toBe(
      '(deploy) sent>=2026-10-12T13:40:00Z',
    );
  });

  it('combines both', () => {
    expect(buildQueryString({ query: 'deploy', conversationId: CHAT, sentAfter: Date.parse('2026-10-01T00:00:00Z') })).toBe(
      `(deploy) ClientThreadId:"${CHAT}" sent>=2026-10-01T00:00:00Z`,
    );
  });

  it('refuses a conversation id that would break out of the quoted filter', () => {
    expect(() => buildQueryString({ query: 'x', conversationId: '19:a" OR "b' })).toThrow(McpToolError);
  });
});

describe('parseQueryResponse (real, redacted response)', () => {
  const page = parseQueryResponse(FIXTURE);

  it('reads Total and MoreResultsAvailable', () => {
    expect(page).toMatchObject({ total: 4821, moreAvailable: true, unmapped: 0 });
    expect(page.results).toHaveLength(4);
  });

  it('maps a channel root post: thread id, message id, subject, no replyTo', () => {
    expect(page.results[0]).toEqual({
      conversationId: '19:channel0001@thread.tacv2',
      messageId: '1791813781971',
      time: '2026-10-12T14:03:01.971Z',
      sender: 'Person 1',
      senderEmail: 'person1@example.test',
      text: 'Has anyone got the staging deploy checklist handy?',
      subject: 'Staging deploy',
    });
  });

  it("maps a channel reply, reading its root off ClientConversationId's ;messageid=", () => {
    expect(page.results[1]).toMatchObject({
      conversationId: '19:channel0002@thread.tacv2',
      messageId: '1791812658351',
      replyTo: '1791790117331',
      text: 'Please deploy the service to QA after the review.',
    });
    expect(page.results[1]).not.toHaveProperty('subject');
  });

  it('maps a group chat hit and a self-chat hit by their bare conversation id', () => {
    expect(page.results[2]).toMatchObject({ conversationId: '19:group0003@thread.v2', messageId: '1791811870502', sender: 'Person 3' });
    expect(page.results[2]).not.toHaveProperty('replyTo');
    expect(page.results[3]).toMatchObject({ conversationId: '48:notes', sender: 'Me Myself' });
  });

  it('keeps the preview as plain text, never treating it as HTML', () => {
    const raw = structuredClone(FIXTURE);
    raw.EntitySets[0].ResultSets[0].Results[0].Source.Preview = 'use <b> for bold & keep it';
    expect(parseQueryResponse(raw).results[0].text).toBe('use <b> for bold & keep it');
  });

  it('counts a hit with no Teams conversation id as unmapped rather than inventing one', () => {
    const raw = structuredClone(FIXTURE);
    delete raw.EntitySets[0].ResultSets[0].Results[0].Source.ClientConversationId;
    delete raw.EntitySets[0].ResultSets[0].Results[0].Source.ClientThreadId;
    const p = parseQueryResponse(raw);
    expect(p.results).toHaveLength(3);
    expect(p.unmapped).toBe(1);
  });

  it('treats a result set with no Results key (a query with no hits, as seen live) as empty', () => {
    expect(parseQueryResponse({ EntitySets: [{ ResultSets: [{ Total: 0, MoreResultsAvailable: false }] }] })).toEqual({
      results: [],
      total: 0,
      moreAvailable: false,
      unmapped: 0,
    });
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

  it('names substrate.office.com and the cause when the request never got a response', async () => {
    const cause = Object.assign(new Error('getaddrinfo EAI_AGAIN substrate.office.com'), { code: 'EAI_AGAIN' });
    const fetchImpl: FetchMock = vi.fn(async () => {
      throw new TypeError('fetch failed', { cause });
    });
    const err = await client(fetchImpl)
      .searchMessages({ query: 'deploy', from: 0, size: 25 })
      .catch((e: unknown) => e);
    expect(err).toBeInstanceOf(McpToolError);
    expect((err as McpToolError).kind).toBe('transport');
    expect((err as McpToolError).message).toMatch(/substrate\.office\.com.*EAI_AGAIN/);
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

  it('after a 401 on a freshly captured token, does not re-capture again for a cooldown unless forced', async () => {
    vi.useFakeTimers({ toFake: ['Date'] });
    try {
      const capture = vi.fn(async () => jwt());
      const fetchImpl: FetchMock = vi.fn(async () => new Response('', { status: 401 }));
      const c = client(fetchImpl, capture);
      await expect(c.searchMessages({ query: 'x', from: 0, size: 25 })).rejects.toMatchObject({ status: 401, kind: 'credential_rejected' });
      expect(capture).toHaveBeenCalledTimes(2); // the initial capture and the one re-capture on 401
      expect(fetchImpl).toHaveBeenCalledTimes(2);

      // Inside the cooldown: fails at once — no capture, no request.
      await expect(c.searchMessages({ query: 'x', from: 0, size: 25 })).rejects.toMatchObject({ status: 401 });
      expect(capture).toHaveBeenCalledTimes(2);
      expect(fetchImpl).toHaveBeenCalledTimes(2);

      // Forced (mode "substrate"): tries again, re-capturing on the 401.
      await expect(c.searchMessages({ query: 'x', from: 0, size: 25 }, { force: true })).rejects.toMatchObject({ status: 401 });
      expect(capture).toHaveBeenCalledTimes(3);

      // After the cooldown, an unforced call tries again too.
      vi.setSystemTime(Date.now() + CAPTURE_RETRY_COOLDOWN_MS + 1);
      await expect(c.searchMessages({ query: 'x', from: 0, size: 25 })).rejects.toMatchObject({ status: 401 });
      expect(capture).toHaveBeenCalledTimes(4);
    } finally {
      vi.useRealTimers();
    }
  });

  it('a 401 that a re-capture cures sets no cooldown', async () => {
    const capture = vi.fn(async () => jwt());
    const fetchImpl: FetchMock = vi.fn();
    fetchImpl.mockResolvedValueOnce(new Response('', { status: 401 }));
    fetchImpl.mockImplementation(async () => json(responseOf([])));
    const c = client(fetchImpl, capture);
    await c.searchMessages({ query: 'x', from: 0, size: 25 });
    await c.searchMessages({ query: 'x', from: 0, size: 25 });
    expect(fetchImpl).toHaveBeenCalledTimes(3);
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
