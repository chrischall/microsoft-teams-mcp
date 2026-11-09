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
  SUBSTRATE_SUGGESTIONS_URL,
  buildSuggestionsBody,
  parseSuggestionsResponse,
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

describe('buildSuggestionsBody', () => {
  it('asks the powerbar scenario for People matching the text, with the required cvid and logical id', () => {
    const body = buildSuggestionsBody('Gary Solomon', 7) as {
      EntityRequests: Record<string, unknown>[];
      Cvid: string;
      LogicalId: string;
      Scenario: { Name: string };
    };
    expect(body.EntityRequests).toEqual([{ EntityType: 'People', Query: { QueryString: 'Gary Solomon' }, Size: 7 }]);
    expect(body.Cvid).toMatch(/^[0-9a-f-]{36}$/);
    expect(body.LogicalId).toMatch(/^[0-9a-f-]{36}$/);
    expect(body.Scenario).toEqual({ Name: 'powerbar' });
  });
});

const OID_G = '0000AAAA-0000-0000-0000-0000000000F1';
const OID_H = '00000000-0000-0000-0000-0000000000f2';

/** A suggestion as the live response carries one (default fields); placeholders only. */
function suggestion(over: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    Id: 'opaque-id',
    DisplayName: 'Solomon, Gary',
    EmailAddresses: ['gary.solomon@example.test'],
    ProxyEmailAddresses: ['SMTP:gary.solomon@example.test'],
    Department: 'Engineering',
    OfficeLocation: 'Remote',
    JobTitle: 'Engineer',
    ImAddress: 'sip:gary.solomon@example.test',
    PeopleSubtype: 'OrganizationUser',
    UserPrincipalName: 'gsolomon@example.test',
    ExternalDirectoryObjectId: OID_G,
    Confidence: 0.9,
    ...over,
  };
}

function suggestionsOf(items: unknown[]): unknown {
  return { Groups: [{ Type: 'People', Suggestions: items }], Instrumentation: {} };
}

describe('parseSuggestionsResponse', () => {
  it('maps an organisation user, deriving the MRI from the lower-cased directory object id', () => {
    expect(parseSuggestionsResponse(suggestionsOf([suggestion()]))).toEqual([
      {
        mri: `8:orgid:${OID_G.toLowerCase()}`,
        name: 'Solomon, Gary',
        email: 'gary.solomon@example.test',
        upn: 'gsolomon@example.test',
        jobTitle: 'Engineer',
        department: 'Engineering',
      },
    ]);
  });

  it('prefers an 8:orgid: MRI field over the object id', () => {
    const [p] = parseSuggestionsResponse(suggestionsOf([suggestion({ MRI: `8:orgid:${OID_H}` })]));
    expect(p.mri).toBe(`8:orgid:${OID_H}`);
  });

  it('ignores an MRI field that is not 8:orgid: and falls back to the object id', () => {
    const [p] = parseSuggestionsResponse(suggestionsOf([suggestion({ MRI: '28:bot' })]));
    expect(p.mri).toBe(`8:orgid:${OID_G.toLowerCase()}`);
  });

  it('keeps an entry with no subtype, and leaves out absent optional fields', () => {
    const raw = suggestion({ PeopleSubtype: undefined, JobTitle: undefined, Department: '', EmailAddresses: [] });
    expect(parseSuggestionsResponse(suggestionsOf([raw]))).toEqual([
      { mri: `8:orgid:${OID_G.toLowerCase()}`, name: 'Solomon, Gary', upn: 'gsolomon@example.test' },
    ]);
  });

  it('drops distribution lists, personal contacts, unknowns and entries with no usable id', () => {
    const people = parseSuggestionsResponse(
      suggestionsOf([
        suggestion({ PeopleSubtype: 'PublicDistributionList', DisplayName: 'All staff' }),
        suggestion({ PeopleSubtype: 'PersonalContact', DisplayName: 'Plumber' }),
        suggestion({ PeopleSubtype: 'Unknown', DisplayName: 'Someone' }),
        suggestion({ ExternalDirectoryObjectId: undefined, DisplayName: 'No id' }),
        suggestion({ ExternalDirectoryObjectId: 'not-a-guid', DisplayName: 'Bad id' }),
        suggestion({ DisplayName: 'Kept', ExternalDirectoryObjectId: OID_H }),
      ]),
    );
    expect(people.map((p) => p.name)).toEqual(['Kept']);
  });

  it('keeps one entry per MRI', () => {
    expect(parseSuggestionsResponse(suggestionsOf([suggestion(), suggestion({ DisplayName: 'Gary S' })]))).toHaveLength(1);
  });

  it('reads only the People group and never throws on shape', () => {
    expect(parseSuggestionsResponse({ Groups: [{ Type: 'Files', Suggestions: [suggestion()] }] })).toEqual([]);
    expect(parseSuggestionsResponse({})).toEqual([]);
    expect(parseSuggestionsResponse(null)).toEqual([]);
    expect(parseSuggestionsResponse({ Groups: [{ Type: 'People', Suggestions: [null, 3, suggestion()] }] })).toHaveLength(1);
  });
});

describe('SubstrateClient.suggestPeople', () => {
  it('POSTs to the suggestions endpoint with the captured bearer, never following a redirect', async () => {
    const token = jwt();
    const fetchImpl: FetchMock = vi.fn(async () => json(suggestionsOf([suggestion()])));
    const people = await client(fetchImpl, async () => token).suggestPeople('  Solomon ', { size: 5 });
    expect(people).toHaveLength(1);

    const [url, init] = fetchImpl.mock.calls[0];
    expect(String(url)).toBe(SUBSTRATE_SUGGESTIONS_URL);
    expect(SUBSTRATE_SUGGESTIONS_URL).toBe('https://substrate.office.com/search/api/v1/suggestions?scenario=powerbar');
    expect(init?.method).toBe('POST');
    expect(init?.redirect).toBe('error');
    const headers = new Headers(init?.headers);
    expect(headers.get('authorization')).toBe(`Bearer ${token}`);
    expect(headers.get('content-type')).toBe('application/json');
    expect(headers.get('accept')).toBe('application/json');
    expect(JSON.parse(String(init?.body))).toMatchObject({
      EntityRequests: [{ EntityType: 'People', Query: { QueryString: 'Solomon' }, Size: 5 }],
    });
  });

  it('asks for 10 by default', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => json(suggestionsOf([])));
    await client(fetchImpl).suggestPeople('x');
    expect(JSON.parse(String(fetchImpl.mock.calls[0][1]?.body))).toMatchObject({ EntityRequests: [{ Size: 10 }] });
  });

  it('asks nothing for blank text', async () => {
    const fetchImpl: FetchMock = vi.fn();
    await expect(client(fetchImpl).suggestPeople('   ')).resolves.toEqual([]);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('re-captures once and replays on a 401', async () => {
    const capture = vi.fn(async () => jwt());
    const fetchImpl: FetchMock = vi.fn();
    fetchImpl.mockResolvedValueOnce(new Response('', { status: 401 }));
    fetchImpl.mockResolvedValueOnce(json(suggestionsOf([suggestion()])));
    await expect(client(fetchImpl, capture).suggestPeople('gary')).resolves.toHaveLength(1);
    expect(capture).toHaveBeenCalledTimes(2);
  });

  it('maps a non-2xx to an McpToolError with its status, naming the people lookup and never the body', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => new Response('secret-ish body', { status: 400 }));
    const err = (await client(fetchImpl).suggestPeople('x').catch((e) => e)) as McpToolError;
    expect(err).toBeInstanceOf(McpToolError);
    expect(err.status).toBe(400);
    expect(err.message).toMatch(/people/i);
    expect(err.message).not.toContain('secret-ish');
  });

  it('refuses a body that is not JSON without echoing it', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => new Response('<html>secret-ish</html>', { status: 200 }));
    const err = (await client(fetchImpl).suggestPeople('x').catch((e) => e)) as McpToolError;
    expect(err).toBeInstanceOf(McpToolError);
    expect(err.message).not.toContain('secret-ish');
  });

  it('names substrate.office.com and the cause when the request never got a response', async () => {
    const cause = Object.assign(new Error('getaddrinfo EAI_AGAIN substrate.office.com'), { code: 'EAI_AGAIN' });
    const fetchImpl: FetchMock = vi.fn(async () => {
      throw new TypeError('fetch failed', { cause });
    });
    const err = (await client(fetchImpl).suggestPeople('x').catch((e) => e)) as McpToolError;
    expect(err.kind).toBe('transport');
    expect(err.message).toMatch(/substrate\.office\.com.*EAI_AGAIN/);
  });

  it('is refused, never capturing, with the bridge disabled', async () => {
    const capture = vi.fn(async () => jwt());
    const c = client(vi.fn(), capture, { ...ENV, TEAMS_DISABLE_FETCHPROXY: 'true' });
    await expect(c.suggestPeople('x')).rejects.toMatchObject({ kind: 'no_credential' });
    expect(capture).not.toHaveBeenCalled();
  });

  it('shares the capture cooldown with search: a failed capture is not retried unless forced', async () => {
    vi.useFakeTimers({ toFake: ['Date'] });
    try {
      const capture = vi.fn(async (): Promise<string> => {
        throw new Error('could not capture a Substrate search token from the browser (timed out)');
      });
      const c = client(vi.fn(), capture);
      await expect(c.searchMessages({ query: 'x', from: 0, size: 25 })).rejects.toThrow(/timed out/);
      await expect(c.suggestPeople('x')).rejects.toThrow(/timed out/);
      expect(capture).toHaveBeenCalledTimes(1);
      await expect(c.suggestPeople('x', { force: true })).rejects.toThrow(/timed out/);
      expect(capture).toHaveBeenCalledTimes(2);
    } finally {
      vi.useRealTimers();
    }
  });
});
