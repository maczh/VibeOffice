# 15: Visual, interaction, file, and production acceptance

**What to build:** The production Vite build demonstrates that the React suite and editor packages meet the agreed visual, behavioral, file-compatibility, security, and offline requirements.

**Blocked by:** 00: Direct editor-engine mount foundation; 14: Host integration example and package delivery; 16: Shared UI module migration to React; 17: Wordx full module migration and naming; 18: Excelx full module migration and naming; 19: PPTX full module migration and naming; 20: Web-wide legacy brand name cleanup.

**Status:** ready-for-agent

**Acceptance note:** Production builds, package/provider tests, CSP checks, and the existing engine UI suites passed in this run. The full screenshot baseline across looks/window sizes, representative real Office/LibreOffice corpus batch, and live controlled WebDAV/CORS deployment are not available in this checkout/session; those acceptance checks remain open.

- [ ] Chromium screenshots of the Start Center and all three editors match the current baseline in the supported looks and representative window sizes, except for the explicitly removed Recent Files UI and added file-provider entries.
- [ ] Real keyboard and pointer workflows remain functional in standalone, suite, and embedded modes.
- [ ] Existing document engine unit, round-trip, corpus, and Office/LibreOffice compatibility checks pass without new format regressions.
- [ ] Local, WebDAV, custom provider, URL, and S3 host workflows pass using controlled test providers.
- [ ] Production assets load without console errors, policy violations, or unconfigured network requests.
- [ ] PWA installation and offline local editing work from the production build.
