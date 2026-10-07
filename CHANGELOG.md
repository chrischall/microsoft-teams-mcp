# Changelog

## [0.2.7](https://github.com/chrischall/microsoft-teams-mcp/compare/v0.2.6...v0.2.7) (2026-10-07)


### Bug Fixes

* **deps:** bump the production-dependencies group with 3 updates ([#33](https://github.com/chrischall/microsoft-teams-mcp/issues/33)) ([f87936f](https://github.com/chrischall/microsoft-teams-mcp/commit/f87936fb326cb441d16ebd6c2ac481b3f299fd9d))
* **deps:** pick up mcp-utils 2.15.0 elicitation opt-out and fetchproxy 3.6.0 room-frame fix ([#35](https://github.com/chrischall/microsoft-teams-mcp/issues/35)) ([d99f34f](https://github.com/chrischall/microsoft-teams-mcp/commit/d99f34fabc98d925aff40f8ff4bc6ec7a851739c))

## [0.2.6](https://github.com/chrischall/microsoft-teams-mcp/compare/v0.2.5...v0.2.6) (2026-10-05)


### Bug Fixes

* **deps:** require @chrischall/mcp-utils 2.14.0 and MCP SDK 2.3.0 ([#29](https://github.com/chrischall/microsoft-teams-mcp/issues/29)) ([9738492](https://github.com/chrischall/microsoft-teams-mcp/commit/973849247e0d40ea35c7e143aa89ef34563c41ed))

## [0.2.5](https://github.com/chrischall/microsoft-teams-mcp/compare/v0.2.4...v0.2.5) (2026-10-03)


### Bug Fixes

* **deps:** adopt @chrischall/mcp-utils 2.12.0 untrusted-content framing ([#27](https://github.com/chrischall/microsoft-teams-mcp/issues/27)) ([19133c4](https://github.com/chrischall/microsoft-teams-mcp/commit/19133c4c2580d09f1ec05afb504eef0f4f538841))
* **deps:** bump @chrischall/mcp-utils to 2.13.0 ([#28](https://github.com/chrischall/microsoft-teams-mcp/issues/28)) ([2ba35d7](https://github.com/chrischall/microsoft-teams-mcp/commit/2ba35d7e50407ca98719115e79f725e4f97cd152))
* **deps:** bump dotenv from 18.0.2 to 18.0.3 in the production-dependencies group ([#23](https://github.com/chrischall/microsoft-teams-mcp/issues/23)) ([b382d2d](https://github.com/chrischall/microsoft-teams-mcp/commit/b382d2d61954b651d118d29d33455044d59badbb))
* keep credentials and report edge_blocked on CDN/WAF blocks (mcp-utils 2.10.0) ([#26](https://github.com/chrischall/microsoft-teams-mcp/issues/26)) ([0918b21](https://github.com/chrischall/microsoft-teams-mcp/commit/0918b21480b9bd9779e19f542d81f854f7e0c349))
* report CDN/WAF blocks as edge_blocked, not a rejected credential (mcp-utils 2.9.0) ([#25](https://github.com/chrischall/microsoft-teams-mcp/issues/25)) ([29b9f39](https://github.com/chrischall/microsoft-teams-mcp/commit/29b9f39636821c6281f79e06c4ed014b5611b76c))

## [0.2.4](https://github.com/chrischall/microsoft-teams-mcp/compare/v0.2.3...v0.2.4) (2026-09-27)


### Bug Fixes

* **deps:** move to [@fetchproxy](https://github.com/fetchproxy) 3.4 for ContextMint Bridge errors, capability subsets and managed pins ([#18](https://github.com/chrischall/microsoft-teams-mcp/issues/18)) ([5157712](https://github.com/chrischall/microsoft-teams-mcp/commit/51577127c4b87c98cff2bbe329a7a5c02e1a840c))
* **deps:** move to @chrischall/mcp-utils 2.8 and [@fetchproxy](https://github.com/fetchproxy) 3.4.1 for clearer browser-bridge errors ([#20](https://github.com/chrischall/microsoft-teams-mcp/issues/20)) ([299d297](https://github.com/chrischall/microsoft-teams-mcp/commit/299d29740aa7c6711d056d34ee162b732b64b623))

## [0.2.3](https://github.com/chrischall/microsoft-teams-mcp/compare/v0.2.2...v0.2.3) (2026-09-24)


### Bug Fixes

* **deps:** bump dotenv from 18.0.1 to 18.0.2 in the production-dependencies group ([#16](https://github.com/chrischall/microsoft-teams-mcp/issues/16)) ([8453fc1](https://github.com/chrischall/microsoft-teams-mcp/commit/8453fc136291fc5d57c90e9462e534bbb5ef5ae9))

## [0.2.2](https://github.com/chrischall/microsoft-teams-mcp/compare/v0.2.1...v0.2.2) (2026-09-23)


### Bug Fixes

* frame Teams message text as untrusted and say which chat/channel the open tools read ([#13](https://github.com/chrischall/microsoft-teams-mcp/issues/13)) ([e42cec0](https://github.com/chrischall/microsoft-teams-mcp/commit/e42cec08997e240091862e7afeb026f419099fc6))

## [0.2.1](https://github.com/chrischall/microsoft-teams-mcp/compare/v0.2.0...v0.2.1) (2026-09-23)


### Bug Fixes

* **deps:** require zod ^4.6.5 to match @chrischall/mcp-utils 2.4.0 ([#12](https://github.com/chrischall/microsoft-teams-mcp/issues/12)) ([edc9103](https://github.com/chrischall/microsoft-teams-mcp/commit/edc910397df9dec55710402bc84ea1bd4c40a14e))
* **deps:** upgrade @chrischall/mcp-utils to 2.4.0 and @fetchproxy/* to 3.2.0 ([#10](https://github.com/chrischall/microsoft-teams-mcp/issues/10)) ([5466e76](https://github.com/chrischall/microsoft-teams-mcp/commit/5466e766190fca651df9709be1502000acb51c06))

## [0.2.0](https://github.com/chrischall/microsoft-teams-mcp/compare/v0.1.0...v0.2.0) (2026-09-21)


### Features

* add Activity feed tool and restrict bridge to teams.cloud.microsoft ([#5](https://github.com/chrischall/microsoft-teams-mcp/issues/5)) ([e9d953e](https://github.com/chrischall/microsoft-teams-mcp/commit/e9d953e3be000a57b6f7e4d37b3a3b370f45bcf8))


### Bug Fixes

* add teams_get_activity to the packaged manifest's tool roster ([#8](https://github.com/chrischall/microsoft-teams-mcp/issues/8)) ([9f28af4](https://github.com/chrischall/microsoft-teams-mcp/commit/9f28af4a9f6cb5161d9f546ddeda11e1d61d1d4f))

## 0.1.0 (2026-09-21)


### Features

* list Teams/channels and read the currently open channel's posts ([#1](https://github.com/chrischall/microsoft-teams-mcp/issues/1)) ([b6c99b3](https://github.com/chrischall/microsoft-teams-mcp/commit/b6c99b323c7971ad441c28a4a661268b29feafd7))
* read the Teams chat list and the currently open chat's messages ([c8a1396](https://github.com/chrischall/microsoft-teams-mcp/commit/c8a1396f9a745803d133a08575c88089ed61bb6a))


### Bug Fixes

* widen packaging descriptions to cover teams and channels ([#4](https://github.com/chrischall/microsoft-teams-mcp/issues/4)) ([958207e](https://github.com/chrischall/microsoft-teams-mcp/commit/958207ec100ac6b86f9d599ff7e26edc04f3c048))
