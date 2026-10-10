import { describe, it, expect, vi } from 'vitest';
import { TeamsClient, DOMAINS, DOM_LIST_SELECTORS, ACTIVITY_FEED_SELECTOR } from '../src/client.js';
import type { FetchproxyTransport } from '@chrischall/mcp-utils/fetchproxy';

function mockTransport(opts: {
  readDomList: (args: { name: string; domain?: string }) => Promise<Record<string, string>[]>;
}): FetchproxyTransport {
  return {
    start: vi.fn().mockResolvedValue(undefined),
    close: vi.fn().mockResolvedValue(undefined),
    status: vi.fn().mockReturnValue({ role: 'host', port: 37149 }),
    role: 'host',
    server: {
      readDomList: vi.fn(opts.readDomList),
    },
  } as unknown as FetchproxyTransport;
}

describe('TeamsClient.getActivity', () => {
  it('reads the activityFeed selector off the first domain that has a tab', async () => {
    const readDomList = vi.fn().mockResolvedValue([
      {
        title: 'Example Person mentioned Example Channel',
        preview: 'Example Channel can someone take a look at this?',
        time: '1:19 PM',
        location: 'Example Team > Example Channel',
      },
    ]);
    const client = new TeamsClient({ transport: mockTransport({ readDomList }) });

    const rows = await client.getActivity();

    expect(rows).toEqual([
      {
        title: 'Example Person mentioned Example Channel',
        preview: 'Example Channel can someone take a look at this?',
        time: '1:19 PM',
        location: 'Example Team > Example Channel',
      },
    ]);
    expect(readDomList).toHaveBeenCalledWith({ name: 'activityFeed', domain: DOMAINS[0] });
  });

  it('starts the transport lazily, once', async () => {
    const readDomList = vi.fn().mockResolvedValue([]);
    const transport = mockTransport({ readDomList });
    const client = new TeamsClient({ transport });

    await client.getActivity();
    await client.getActivity();

    expect(transport.start).toHaveBeenCalledTimes(1);
  });

  it('starts the transport once when the first calls arrive in parallel', async () => {
    const readDomList = vi.fn().mockResolvedValue([]);
    const transport = mockTransport({ readDomList });
    let release!: () => void;
    (transport.start as ReturnType<typeof vi.fn>).mockImplementation(
      () => new Promise<void>((resolve) => { release = resolve; }),
    );
    const client = new TeamsClient({ transport });

    const calls = Promise.all([client.getActivity(), client.getActivity(), client.getActivity()]);
    await Promise.resolve();
    release();
    await calls;

    expect(transport.start).toHaveBeenCalledTimes(1);
    expect(readDomList).toHaveBeenCalledTimes(3);
  });

  it('retries the start on the next call after a failed start', async () => {
    const readDomList = vi.fn().mockResolvedValue([]);
    const transport = mockTransport({ readDomList });
    (transport.start as ReturnType<typeof vi.fn>)
      .mockRejectedValueOnce(new Error('identity load failed'))
      .mockResolvedValue(undefined);
    const client = new TeamsClient({ transport });

    await expect(client.getActivity()).rejects.toThrow(/identity load failed/);
    await client.getActivity();

    expect(transport.start).toHaveBeenCalledTimes(2);
    expect(readDomList).toHaveBeenCalledTimes(1);
  });
});

describe('TeamsClient error handling', () => {
  it('propagates a non-"no tab" failure unchanged', async () => {
    const readDomList = vi.fn().mockRejectedValue(new Error('read_dom_list name not in declared set: activityFeed'));
    const client = new TeamsClient({ transport: mockTransport({ readDomList }) });

    await expect(client.getActivity()).rejects.toThrow(/not in declared set/);
    expect(readDomList).toHaveBeenCalledTimes(1);
  });

  it('throws an actionable error when the tab is not reachable', async () => {
    const readDomList = vi.fn().mockRejectedValue(new Error('no tab matching https://x/'));
    const client = new TeamsClient({ transport: mockTransport({ readDomList }) });

    await expect(client.getActivity()).rejects.toThrow(
      /could not reach a signed-in Teams tab on teams\.cloud\.microsoft/,
    );
    expect(readDomList).toHaveBeenCalledTimes(1);
  });
});

describe('TeamsClient.close', () => {
  it('closes the transport and allows a fresh start on the next call', async () => {
    const readDomList = vi.fn().mockResolvedValue([]);
    const transport = mockTransport({ readDomList });
    const client = new TeamsClient({ transport });

    await client.getActivity();
    await client.close();
    await client.getActivity();

    expect(transport.close).toHaveBeenCalledTimes(1);
    expect(transport.start).toHaveBeenCalledTimes(2);
  });
});

describe('declared DOM-list selectors', () => {
  it('declares only the Activity feed — chats and channels are read through the chat-service API', () => {
    expect(DOM_LIST_SELECTORS).toEqual([ACTIVITY_FEED_SELECTOR]);
    expect(DOM_LIST_SELECTORS.map((s) => s.name)).toEqual(['activityFeed']);
  });

  it('no longer exposes the chat/channel DOM readers', () => {
    const client = new TeamsClient({ transport: mockTransport({ readDomList: vi.fn() }) }) as unknown as Record<string, unknown>;

    for (const gone of ['listChats', 'getOpenChatMessages', 'listTeamsAndChannels', 'getOpenChannelPosts', 'getSelectedConversations']) {
      expect(client[gone], gone).toBeUndefined();
    }
  });
});
