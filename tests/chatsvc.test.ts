import { describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { McpToolError, RequestTimeoutError, WriteOutcomeUnknownError } from '@chrischall/mcp-utils';
import { ChatServiceClient, DEFAULT_REGION, chatServiceBase, outcomeUnknownStatus } from '../src/chatsvc.js';

/** A hand-written fixture from `tests/fixtures/` — invented names and text only. */
function fixture(name: string): unknown {
  return JSON.parse(readFileSync(new URL(`./fixtures/${name}`, import.meta.url), 'utf8'));
}

const ME_OID = '00000000-0000-0000-0000-0000000000a1';

/** A JWT carrying `exp` and `oid`. The signature is irrelevant here. */
function jwt(secondsFromNow = 3600, claims: Record<string, unknown> = { oid: ME_OID }): string {
  const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ alg: 'none' })}.${b64({
    exp: Math.floor(Date.now() / 1000) + secondsFromNow,
    ...claims,
  })}.x`;
}

/**
 * Test env. Every client gets an EXPLICIT env, so the cache must be turned off
 * here — otherwise it defaults on and resolves to the developer's real
 * `~/.microsoft-teams-mcp`.
 */
function env(overrides: Record<string, string> = {}): NodeJS.ProcessEnv {
  return { TEAMS_TOKEN_CACHE: 'false', ...overrides };
}

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

const BASE = `https://teams.cloud.microsoft/api/chatsvc/${DEFAULT_REGION}/v1`;

type FetchMock = ReturnType<typeof vi.fn<(input: RequestInfo | URL, init?: RequestInit) => Promise<Response>>>;

function client(fetchImpl: FetchMock, capture: () => Promise<string> = async () => jwt()) {
  return new ChatServiceClient({
    env: env(),
    captureToken: capture,
    fetchImpl: fetchImpl as unknown as typeof fetch,
  });
}

function urlOf(mock: FetchMock, n: number): string {
  return String(mock.mock.calls[n]?.[0]);
}

function headersOf(mock: FetchMock, n: number): Headers {
  return new Headers(mock.mock.calls[n]?.[1]?.headers);
}

/** A message whose id (arrival ms) is `id`. */
function msg(id: number): Record<string, unknown> {
  return { id: String(id), messagetype: 'Text', content: `m${id}` };
}

describe('configuration', () => {
  it('defaults to the amer region and honours TEAMS_CHATSVC_REGION / TEAMS_CHATSVC_BASE', () => {
    expect(chatServiceBase(env())).toBe(BASE);
    expect(chatServiceBase(env({ TEAMS_CHATSVC_REGION: 'emea' }))).toBe(
      'https://teams.cloud.microsoft/api/chatsvc/emea/v1',
    );
    expect(chatServiceBase(env({ TEAMS_CHATSVC_BASE: 'https://example.test/x/v1/' }))).toBe(
      'https://example.test/x/v1',
    );
  });

  it('does not throw with no credential, and defers the error to the first call', async () => {
    const c = new ChatServiceClient({ env: env({ TEAMS_DISABLE_FETCHPROXY: '1' }) });
    await expect(c.listConversations()).rejects.toThrow(/No Teams credential is configured/);
  });

  it('reports a non-secret token source', () => {
    expect(new ChatServiceClient({ env: env({ TEAMS_ACCESS_TOKEN: 'x' }) }).tokenSource).toBe(
      'TEAMS_ACCESS_TOKEN',
    );
    expect(new ChatServiceClient({ env: env() }).tokenSource).toBe('browser capture');
  });
});

describe('requests', () => {
  it('sends the bearer, the redirect override and a JSON accept on every request', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({ conversations: [] }));
    const token = jwt();
    await client(fetchImpl, async () => token).listConversations();

    const h = headersOf(fetchImpl, 0);
    expect(h.get('authorization')).toBe(`Bearer ${token}`);
    expect(h.get('behavioroverride')).toBe('redirectAs404');
    expect(h.get('accept')).toBe('application/json');
  });

  it('derives the signed-in user MRI from the token oid', async () => {
    const c = client(vi.fn(async () => jsonResponse({})));
    await expect(c.meMri()).resolves.toBe(`8:orgid:${ME_OID}`);
  });

  it('refuses a token with no oid claim', async () => {
    const c = client(vi.fn(async () => jsonResponse({})), async () => jwt(3600, {}));
    await expect(c.meMri()).rejects.toBeInstanceOf(McpToolError);
  });

  it('derives token expiry from the JWT exp', async () => {
    const c = client(vi.fn(async () => jsonResponse({ conversations: [] })), async () => jwt(7200));
    await c.listConversations();
    const left = (c.tokenExpiresAt() ?? 0) - Date.now();
    expect(left).toBeGreaterThan(7000_000);
    expect(left).toBeLessThan(7300_000);
  });
});

describe('401 handling', () => {
  it('re-captures once and replays once on a 401', async () => {
    const capture = vi.fn(async () => jwt());
    let n = 0;
    const fetchImpl: FetchMock = vi.fn(async () =>
      n++ === 0 ? jsonResponse({ message: 'expired' }, 401) : jsonResponse({ conversations: [] }),
    );
    await expect(client(fetchImpl, capture).listConversations()).resolves.toEqual([]);
    expect(capture).toHaveBeenCalledTimes(2);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('fails with a re-open-Teams hint on a second 401', async () => {
    const capture = vi.fn(async () => jwt());
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({}, 401));
    const p = client(fetchImpl, capture).listConversations();
    await expect(p).rejects.toBeInstanceOf(McpToolError);
    await expect(p).rejects.toMatchObject({ status: 401, hint: expect.stringMatching(/Teams/) });
    expect(capture).toHaveBeenCalledTimes(2);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('never reaches for the browser to refresh a directly-supplied token', async () => {
    const capture = vi.fn(async () => jwt());
    const c = new ChatServiceClient({
      env: env({ TEAMS_ACCESS_TOKEN: jwt() }),
      captureToken: capture,
      fetchImpl: (async () => jsonResponse({}, 401)) as unknown as typeof fetch,
    });
    await expect(c.listConversations()).rejects.toThrow(/TEAMS_ACCESS_TOKEN/);
    expect(capture).not.toHaveBeenCalled();
  });
});

describe('other failures', () => {
  it.each([429, 500, 503])('turns a %i into an McpToolError carrying the status, no retry', async (status) => {
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({ message: 'nope' }, status));
    const p = client(fetchImpl).listConversations();
    await expect(p).rejects.toBeInstanceOf(McpToolError);
    await expect(p).rejects.toMatchObject({ status });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
});

describe('listConversations', () => {
  it('follows backwardLink pages and drops 48: system ids', async () => {
    const fetchImpl: FetchMock = vi.fn(async (input) =>
      String(input).includes('syncState=page2')
        ? jsonResponse(fixture('conversations-page2.json'))
        : jsonResponse(fixture('conversations-page1.json')),
    );
    const convs = await client(fetchImpl).listConversations();

    expect(urlOf(fetchImpl, 0)).toBe(
      `${BASE}/users/ME/conversations?view=msnp24Equivalent&pageSize=200`,
    );
    expect(urlOf(fetchImpl, 1)).toBe(
      `${BASE}/users/ME/conversations?view=msnp24Equivalent&pageSize=200&syncState=page2`,
    );
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect(convs.map((c) => c.id)).toEqual([
      '19:aaaa1111@thread.v2',
      '19:bbbb2222@thread.tacv2',
      '19:00000000-0000-0000-0000-0000000000a1_00000000-0000-0000-0000-0000000000b2@unq.gbl.spaces',
    ]);
  });

  it('stops on an empty page even when it carries a link', async () => {
    let n = 0;
    const fetchImpl: FetchMock = vi.fn(async () =>
      jsonResponse({
        conversations: n++ === 0 ? [{ id: '19:x@thread.v2' }] : [],
        _metadata: { backwardLink: `${BASE}/users/ME/conversations?syncState=more${n}` },
      }),
    );
    await expect(client(fetchImpl).listConversations()).resolves.toHaveLength(1);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('caps paging at 10 pages', async () => {
    let n = 0;
    const fetchImpl: FetchMock = vi.fn(async () => {
      n++;
      return jsonResponse({
        conversations: [{ id: `19:c${n}@thread.v2` }],
        _metadata: { backwardLink: `${BASE}/users/ME/conversations?syncState=p${n}` },
      });
    });
    await expect(client(fetchImpl).listConversations()).resolves.toHaveLength(10);
    expect(fetchImpl).toHaveBeenCalledTimes(10);
  });

  it('re-roots a backwardLink on another host onto the configured base, never sending the token there', async () => {
    let n = 0;
    const fetchImpl: FetchMock = vi.fn(async () =>
      jsonResponse({
        conversations: n++ === 0 ? [{ id: '19:x@thread.v2' }] : [],
        _metadata: {
          backwardLink: 'https://amer.ng.msg.teams.microsoft.com/v1/users/ME/conversations?syncState=s2',
        },
      }),
    );
    await client(fetchImpl).listConversations();
    expect(urlOf(fetchImpl, 1)).toBe(`${BASE}/users/ME/conversations?syncState=s2`);
  });

  it('refuses a backwardLink that is not a chat-service path', async () => {
    const fetchImpl: FetchMock = vi.fn(async () =>
      jsonResponse({
        conversations: [{ id: '19:x@thread.v2' }],
        _metadata: { backwardLink: 'https://evil.example.com/steal' },
      }),
    );
    await expect(client(fetchImpl).listConversations()).rejects.toThrow(/Refusing to follow/);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
});

describe('getMessages', () => {
  const CONV = '19:aaaa1111@thread.v2';

  function pagedFetch(): FetchMock {
    return vi.fn(async (input) =>
      String(input).includes('syncState=older')
        ? jsonResponse(fixture('messages-page2.json'))
        : jsonResponse(fixture('messages-page1.json')),
    );
  }

  it('requests the encoded conversation and returns every page newest-first', async () => {
    const fetchImpl = pagedFetch();
    const msgs = await client(fetchImpl).getMessages(CONV, { max: 100 });
    expect(urlOf(fetchImpl, 0)).toBe(
      `${BASE}/users/ME/conversations/${encodeURIComponent(CONV)}/messages` +
        '?view=msnp24Equivalent|supportsMessageProperties&pageSize=200&startTime=1',
    );
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect(msgs.map((m) => m.id)).toEqual([
      '1760000300000',
      '1760000200000',
      '1760000100000',
      '1760000000000',
    ]);
  });

  it('stops at the first message stopAt accepts, excluding it, without fetching further', async () => {
    const fetchImpl = pagedFetch();
    const msgs = await client(fetchImpl).getMessages(CONV, {
      max: 100,
      stopAt: (m) => Number(m.id) <= 1760000200000,
    });
    expect(msgs.map((m) => m.id)).toEqual(['1760000300000']);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('stops when max is reached', async () => {
    const fetchImpl = pagedFetch();
    const msgs = await client(fetchImpl).getMessages(CONV, { max: 3 });
    expect(msgs.map((m) => m.id)).toEqual(['1760000300000', '1760000200000', '1760000100000']);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('stops when there are no more pages', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({ messages: [msg(5)] }));
    await expect(client(fetchImpl).getMessages(CONV, { max: 100 })).resolves.toHaveLength(1);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('caps paging at 20 pages', async () => {
    let n = 1000;
    const fetchImpl: FetchMock = vi.fn(async () => {
      n--;
      return jsonResponse({
        messages: [msg(n)],
        _metadata: { backwardLink: `${BASE}/users/ME/conversations/x/messages?syncState=${n}` },
      });
    });
    const onCapped = vi.fn();
    await expect(client(fetchImpl).getMessages(CONV, { max: 1000, onCapped })).resolves.toHaveLength(20);
    expect(fetchImpl).toHaveBeenCalledTimes(20);
    expect(onCapped).toHaveBeenCalledTimes(1);
  });

  it('reports hitting max through onCapped', async () => {
    const onCapped = vi.fn();
    await client(pagedFetch()).getMessages(CONV, { max: 3, onCapped });
    expect(onCapped).toHaveBeenCalledTimes(1);
  });

  it('does not report a cap when stopAt or the end of the history ended the read', async () => {
    const onCapped = vi.fn();
    await client(pagedFetch()).getMessages(CONV, { max: 100, onCapped });
    await client(pagedFetch()).getMessages(CONV, { max: 100, onCapped, stopAt: (m) => Number(m.id) <= 1760000200000 });
    expect(onCapped).not.toHaveBeenCalled();
  });

  it('returns nothing, without a request, for a max of 0 or less', async () => {
    const fetchImpl = pagedFetch();
    await expect(client(fetchImpl).getMessages(CONV, { max: 0 })).resolves.toEqual([]);
    await expect(client(fetchImpl).getMessages(CONV, { max: -5 })).resolves.toEqual([]);
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});

describe('fetchImage', () => {
  const IMG = 'https://us-api.asm.skype.com/v1/objects/0-abc-123/views/imgo';

  it('fetches an asm.skype.com image with the bearer and returns its bytes', async () => {
    const fetchImpl: FetchMock = vi.fn(
      async () =>
        new Response(new Uint8Array([0xff, 0xd8, 0xff]), {
          status: 200,
          headers: { 'content-type': 'image/jpeg' },
        }),
    );
    const img = await client(fetchImpl).fetchImage(IMG);
    expect(img.mimeType).toBe('image/jpeg');
    expect(Buffer.isBuffer(img.data)).toBe(true);
    expect([...img.data]).toEqual([0xff, 0xd8, 0xff]);
    expect(urlOf(fetchImpl, 0)).toBe(IMG);
    expect(headersOf(fetchImpl, 0).get('authorization')).toMatch(/^Bearer /);
  });

  it.each([
    'https://evil.com/v1/objects/x/views/imgo',
    'https://asm.skype.com.evil.com/v1/objects/x/views/imgo',
    'http://us-api.asm.skype.com/v1/objects/x/views/imgo',
    'https://asm.skype.com/v1/objects/x/views/imgo',
    'not a url',
  ])('refuses %s without sending anything', async (url) => {
    const capture = vi.fn(async () => jwt());
    const fetchImpl: FetchMock = vi.fn(async () => new Response('x'));
    await expect(client(fetchImpl, capture).fetchImage(url)).rejects.toBeInstanceOf(McpToolError);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('refuses an image over 5 MB', async () => {
    const big = new Uint8Array(5 * 1024 * 1024 + 1);
    const fetchImpl: FetchMock = vi.fn(
      async () => new Response(big, { status: 200, headers: { 'content-type': 'image/png' } }),
    );
    await expect(client(fetchImpl).fetchImage(IMG)).rejects.toBeInstanceOf(McpToolError);
  });

  it('re-captures once on a 401 for an image too', async () => {
    const capture = vi.fn(async () => jwt());
    let n = 0;
    const fetchImpl: FetchMock = vi.fn(async () =>
      n++ === 0
        ? new Response('no', { status: 401 })
        : new Response(new Uint8Array([1]), { headers: { 'content-type': 'image/png' } }),
    );
    await expect(client(fetchImpl, capture).fetchImage(IMG)).resolves.toMatchObject({
      mimeType: 'image/png',
    });
    expect(capture).toHaveBeenCalledTimes(2);
  });
});

describe('meDisplayName', () => {
  it('reads the signed-in user name from the token name claim', async () => {
    const c = client(vi.fn(async () => jsonResponse({})), async () => jwt(3600, { oid: ME_OID, name: 'Me Myself' }));
    await expect(c.meDisplayName()).resolves.toBe('Me Myself');
  });

  it('is an empty string when the token carries no name', async () => {
    const c = client(vi.fn(async () => jsonResponse({})));
    await expect(c.meDisplayName()).resolves.toBe('');
  });
});

describe('getMessage', () => {
  it('GETs one message of the encoded conversation', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse(msg(1760000000000)));
    const m = await client(fetchImpl).getMessage('19:a b@thread.v2', '1760000000000');
    expect(m?.content).toBe('m1760000000000');
    expect(urlOf(fetchImpl, 0)).toBe(
      `${BASE}/users/ME/conversations/${encodeURIComponent('19:a b@thread.v2')}/messages/1760000000000`,
    );
    expect(fetchImpl.mock.calls[0]?.[1]?.method).toBe('GET');
  });

  it('is undefined for a 404', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({}, 404));
    await expect(client(fetchImpl).getMessage('19:a@thread.v2', '1')).resolves.toBeUndefined();
  });

  it('still throws for other failures', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({}, 500));
    await expect(client(fetchImpl).getMessage('19:a@thread.v2', '1')).rejects.toMatchObject({ status: 500 });
  });
});

describe('sendMessage', () => {
  const BODY = { content: '<p>hi</p>', clientmessageid: '1234567890123456789' };

  function created(id = '1760000999999'): Response {
    return new Response('', {
      status: 201,
      headers: { location: `${BASE}/users/ME/conversations/19%3Aa%40thread.v2/messages/${id}` },
    });
  }

  it('POSTs the JSON body to the encoded conversation and returns the new id from Location', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => created());
    const res = await client(fetchImpl).sendMessage('19:a@thread.v2', BODY);
    expect(res).toEqual({ id: '1760000999999' });
    expect(urlOf(fetchImpl, 0)).toBe(`${BASE}/users/ME/conversations/${encodeURIComponent('19:a@thread.v2')}/messages`);
    const init = fetchImpl.mock.calls[0]?.[1];
    expect(init?.method).toBe('POST');
    expect(JSON.parse(String(init?.body))).toEqual(BODY);
    const h = headersOf(fetchImpl, 0);
    expect(h.get('content-type')).toBe('application/json');
    expect(h.get('authorization')).toMatch(/^Bearer /);
    expect(h.get('behavioroverride')).toBe('redirectAs404');
  });

  it('encodes a channel thread target (id;messageid=root) as one path segment', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => created());
    await client(fetchImpl).sendMessage('19:c@thread.tacv2;messageid=1760000000000', BODY);
    expect(urlOf(fetchImpl, 0)).toBe(
      `${BASE}/users/ME/conversations/${encodeURIComponent('19:c@thread.tacv2;messageid=1760000000000')}/messages`,
    );
  });

  it('has no id when the response carries no Location', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => new Response('', { status: 201 }));
    await expect(client(fetchImpl).sendMessage('c', BODY)).resolves.toEqual({});
  });

  it('re-captures once and replays once on a 401 (the first POST was refused, not accepted)', async () => {
    const capture = vi.fn(async () => jwt());
    let n = 0;
    const fetchImpl: FetchMock = vi.fn(async () => (n++ === 0 ? jsonResponse({}, 401) : created()));
    await expect(client(fetchImpl, capture).sendMessage('c', BODY)).resolves.toEqual({ id: '1760000999999' });
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('reports a dropped connection after sending as an unknown outcome, never a plain error', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => {
      throw new TypeError('fetch failed');
    });
    const p = client(fetchImpl).sendMessage('c', BODY);
    await expect(p).rejects.toBeInstanceOf(WriteOutcomeUnknownError);
    await expect(p).rejects.toMatchObject({ outcomeUnknown: true, timedOut: false });
  });

  it('reports a timeout as an unknown outcome', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => {
      throw new RequestTimeoutError('Teams', 60_000);
    });
    await expect(client(fetchImpl).sendMessage('c', BODY)).rejects.toMatchObject({
      outcomeUnknown: true,
      timedOut: true,
    });
  });

  it.each([400, 403, 404, 429])('maps a %i refusal to a plain McpToolError with its status', async (status) => {
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({}, status));
    const p = client(fetchImpl).sendMessage('c', BODY);
    await expect(p).rejects.toMatchObject({ status });
    await expect(p).rejects.not.toBeInstanceOf(WriteOutcomeUnknownError);
  });

  it.each([500, 502, 503, 504])('reports a %i after sending as an unknown outcome, never a plain error', async (status) => {
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({}, status));
    let err: unknown;
    try {
      await client(fetchImpl).sendMessage('c', BODY);
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(WriteOutcomeUnknownError);
    const e = err as WriteOutcomeUnknownError;
    expect(e).toMatchObject({ outcomeUnknown: true, timedOut: false, method: 'POST' });
    expect(outcomeUnknownStatus(e)).toBe(status);
    expect(e.hint).toMatch(/check the conversation/i);
    // Exactly one POST: a 5xx is never replayed.
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('reports a 5xx on the post-401 replay as an unknown outcome too', async () => {
    let n = 0;
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({}, n++ === 0 ? 401 : 503));
    const p = client(fetchImpl).sendMessage('c', BODY);
    await expect(p).rejects.toBeInstanceOf(WriteOutcomeUnknownError);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it('has no status on an unknown outcome caused by a dropped connection', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => {
      throw new TypeError('fetch failed');
    });
    const err = await client(fetchImpl)
      .sendMessage('c', BODY)
      .catch((e: unknown) => e);
    expect(outcomeUnknownStatus(err as WriteOutcomeUnknownError)).toBeUndefined();
  });
});

describe('setReaction / removeReaction', () => {
  const CONV = '19:a@thread.v2';
  const PROPS_URL = `${BASE}/users/ME/conversations/${encodeURIComponent(CONV)}/messages/1760000000000/properties?name=emotions`;

  it('PUTs the emotions property with the key and the reaction time', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => new Response('', { status: 200 }));
    await client(fetchImpl).setReaction(CONV, '1760000000000', 'like', 1_760_000_123_456);
    expect(urlOf(fetchImpl, 0)).toBe(PROPS_URL);
    const init = fetchImpl.mock.calls[0]?.[1];
    expect(init?.method).toBe('PUT');
    expect(JSON.parse(String(init?.body))).toEqual({ emotions: { key: 'like', value: 1_760_000_123_456 } });
    expect(headersOf(fetchImpl, 0).get('content-type')).toBe('application/json');
  });

  it('DELETEs the same property with just the key to remove', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => new Response(null, { status: 204 }));
    await client(fetchImpl).removeReaction(CONV, '1760000000000', 'heart');
    expect(urlOf(fetchImpl, 0)).toBe(PROPS_URL);
    const init = fetchImpl.mock.calls[0]?.[1];
    expect(init?.method).toBe('DELETE');
    expect(JSON.parse(String(init?.body))).toEqual({ emotions: { key: 'heart' } });
  });

  it('maps a 4xx to a plain McpToolError with its status', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({}, 404));
    const p = client(fetchImpl).setReaction(CONV, '1', 'like');
    await expect(p).rejects.toMatchObject({ status: 404 });
    await expect(p).rejects.not.toBeInstanceOf(WriteOutcomeUnknownError);
  });

  it.each([
    ['PUT', (c: ChatServiceClient) => c.setReaction(CONV, '1', 'like')],
    ['DELETE', (c: ChatServiceClient) => c.removeReaction(CONV, '1', 'like')],
  ] as const)('reports a 503 on the %s as an unknown outcome pointing at the message', async (method, call) => {
    const fetchImpl: FetchMock = vi.fn(async () => jsonResponse({}, 503));
    const err = await call(client(fetchImpl)).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(WriteOutcomeUnknownError);
    expect(err).toMatchObject({ method });
    expect(outcomeUnknownStatus(err as WriteOutcomeUnknownError)).toBe(503);
    expect((err as WriteOutcomeUnknownError).hint).toMatch(/check the message/i);
  });

  it('reports a timeout as an unknown outcome too — the change may have landed', async () => {
    const fetchImpl: FetchMock = vi.fn(async () => {
      throw new RequestTimeoutError('Teams', 60_000);
    });
    await expect(client(fetchImpl).setReaction(CONV, '1', 'like')).rejects.toMatchObject({
      outcomeUnknown: true,
      timedOut: true,
    });
  });
});
