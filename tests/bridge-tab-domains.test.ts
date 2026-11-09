import { describe, it, expect } from 'vitest';
import { BRIDGE_TRUST_DOMAINS, withBridgeScope } from '../src/bridge-scope.js';

describe('withBridgeScope tab domains', () => {
  it('asks for a tab on Teams only, whichever domains the capture trusts', () => {
    const scoped = withBridgeScope({}, {}, ['contoso-my.sharepoint.com']);
    expect(scoped.tabDomains).toEqual(['teams.cloud.microsoft']);
    expect(scoped.domains).toEqual([...BRIDGE_TRUST_DOMAINS]);
  });
});
