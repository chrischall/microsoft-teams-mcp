import { describe, it, expect, vi } from 'vitest';
import { TeamsClient, DOMAINS } from '../src/client.js';
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

describe('TeamsClient.listChats', () => {
  it('reads the chatList selector off the first domain that has a tab', async () => {
    const readDomList = vi.fn().mockResolvedValue([
      { title: 'Standup', preview: 'hi', time: '9:00 AM', conversationKey: 'k1', itemType: 'chat' },
    ]);
    const client = new TeamsClient({ transport: mockTransport({ readDomList }) });

    const rows = await client.listChats();

    expect(rows).toEqual([
      { title: 'Standup', preview: 'hi', time: '9:00 AM', conversationKey: 'k1', itemType: 'chat' },
    ]);
    expect(readDomList).toHaveBeenCalledWith({ name: 'chatList', domain: DOMAINS[0] });
  });

  it('starts the transport lazily, once', async () => {
    const readDomList = vi.fn().mockResolvedValue([]);
    const transport = mockTransport({ readDomList });
    const client = new TeamsClient({ transport });

    await client.listChats();
    await client.listChats();

    expect(transport.start).toHaveBeenCalledTimes(1);
  });
});

describe('TeamsClient.getOpenChatMessages', () => {
  it('returns rows from the chatMessages selector', async () => {
    const readDomList = vi.fn().mockResolvedValue([
      { sender: 'Alice', time: '2026-09-21T10:00:00.000Z', messageId: 'timestamp-1', text: 'hi' },
    ]);
    const client = new TeamsClient({ transport: mockTransport({ readDomList }) });

    const rows = await client.getOpenChatMessages();

    expect(rows).toEqual([
      { sender: 'Alice', time: '2026-09-21T10:00:00.000Z', messageId: 'timestamp-1', text: 'hi' },
    ]);
    expect(readDomList).toHaveBeenCalledWith({ name: 'chatMessages', domain: DOMAINS[0] });
  });
});

describe('TeamsClient.listTeamsAndChannels', () => {
  it('reads the teamsAndChannels selector off the first domain that has a tab', async () => {
    const readDomList = vi.fn().mockResolvedValue([
      { title: 'General', teamName: 'TruAudience', time: '9/18', conversationKey: 'k1', itemType: 'channel' },
      { title: 'TruAudience', time: '', conversationKey: 'k2', itemType: 'team' },
    ]);
    const client = new TeamsClient({ transport: mockTransport({ readDomList }) });

    const rows = await client.listTeamsAndChannels();

    expect(rows).toEqual([
      { title: 'General', teamName: 'TruAudience', time: '9/18', conversationKey: 'k1', itemType: 'channel' },
      { title: 'TruAudience', time: '', conversationKey: 'k2', itemType: 'team' },
    ]);
    expect(readDomList).toHaveBeenCalledWith({ name: 'teamsAndChannels', domain: DOMAINS[0] });
  });
});

describe('TeamsClient.getActivity', () => {
  it('reads the activityFeed selector off the first domain that has a tab', async () => {
    const readDomList = vi.fn().mockResolvedValue([
      {
        title: 'Sanderson, Xi mentioned MS DevOps',
        preview: 'MS DevOps can someone please help with this harness deployment job?',
        time: '1:19 PM',
        location: 'NS_MS Product CE > MS DevOps',
      },
    ]);
    const client = new TeamsClient({ transport: mockTransport({ readDomList }) });

    const rows = await client.getActivity();

    expect(rows).toEqual([
      {
        title: 'Sanderson, Xi mentioned MS DevOps',
        preview: 'MS DevOps can someone please help with this harness deployment job?',
        time: '1:19 PM',
        location: 'NS_MS Product CE > MS DevOps',
      },
    ]);
    expect(readDomList).toHaveBeenCalledWith({ name: 'activityFeed', domain: DOMAINS[0] });
  });
});

describe('TeamsClient.getOpenChannelPosts', () => {
  it('returns rows from the channelPosts selector', async () => {
    const readDomList = vi.fn().mockResolvedValue([
      {
        sender: 'Alice',
        subject: 'Deploy freeze',
        time: 'Tuesday, August 4, 2026 7:50 AM',
        messageId: 'timestamp-1',
        text: 'hi',
      },
    ]);
    const client = new TeamsClient({ transport: mockTransport({ readDomList }) });

    const rows = await client.getOpenChannelPosts();

    expect(rows).toEqual([
      {
        sender: 'Alice',
        subject: 'Deploy freeze',
        time: 'Tuesday, August 4, 2026 7:50 AM',
        messageId: 'timestamp-1',
        text: 'hi',
      },
    ]);
    expect(readDomList).toHaveBeenCalledWith({ name: 'channelPosts', domain: DOMAINS[0] });
  });
});

describe('TeamsClient error handling', () => {
  it('propagates a non-"no tab" failure unchanged', async () => {
    const readDomList = vi.fn().mockRejectedValue(new Error('read_dom_list name not in declared set: chatList'));
    const client = new TeamsClient({ transport: mockTransport({ readDomList }) });

    await expect(client.listChats()).rejects.toThrow(/not in declared set/);
    expect(readDomList).toHaveBeenCalledTimes(1);
  });

  it('throws an actionable error when the tab is not reachable', async () => {
    const readDomList = vi.fn().mockRejectedValue(new Error('no tab matching https://x/'));
    const client = new TeamsClient({ transport: mockTransport({ readDomList }) });

    await expect(client.listChats()).rejects.toThrow(
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

    await client.listChats();
    await client.close();
    await client.listChats();

    expect(transport.close).toHaveBeenCalledTimes(1);
    expect(transport.start).toHaveBeenCalledTimes(2);
  });
});
