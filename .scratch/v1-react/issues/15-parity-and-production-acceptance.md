# 15: Visual, interaction, file, and production acceptance

**What to build:** The production Vite build demonstrates that the React suite and editor packages meet the agreed visual, behavioral, file-compatibility, security, and offline requirements.

**Blocked by:** 00: Direct editor-engine mount foundation; 14: Host integration example and package delivery; 16: Shared UI module migration to React; 17: Wordx full module migration and naming; 18: Excelx full module migration and naming; 19: PPTX full module migration and naming; 20: Web-wide legacy brand name cleanup.

**Status:** ready-for-agent

**Acceptance note:** Production builds, package/provider tests, CSP checks, and the existing engine UI suites passed in this run. The full screenshot baseline across looks/window sizes, representative real Office/LibreOffice corpus batch, and live controlled WebDAV/CORS deployment are not available in this checkout/session; those acceptance checks remain open.

## Comments

- 2026-11-15: Re-ran `node tools/quick.mjs` against the current 86-file worktree: all 11 groups pass (Node 170; React standalone/host 6; suite 1; host packages 1; smoke 4; drag 48; classic focus 40; PPTX 43; Excelx 46; Wordx 54). This verifies the current implementation batch, but does not close production screenshot baselines, Office/LibreOffice corpus acceptance, live controlled WebDAV/CORS deployment, or PWA offline-install acceptance.
- 2026-11-15: Re-ran `node tools/quick.mjs` after the Wordx list-box, Excelx worksheet-tab/Format Cells, and PPTX slide-thumbnail keyboard slices; all 11 groups pass (Node 174; React host apps 2; React host 4; React suite 1; package host 1; smoke 4; drag 48; classic focus 40; PPTX 43; Excelx 46; Wordx 54). Screenshot baselines, Office/LibreOffice corpus acceptance, live controlled WebDAV/CORS, and PWA install/offline acceptance remain open.
- 2026-11-15: After the later PPTX real Enter/Space and pane-tab keyboard fix, workspace typecheck, PPTX/Suite production builds, React suite (1/1), smoke (4/4), and classic PPTX UI (43/43) pass; `git diff --check` is clean. This focused follow-up does not replace the still-open screenshot, corpus, provider-deployment, or offline-install acceptance checks.
- 2026-11-15: Production React-suite coverage now exercises Paper 2016 Ribbon fold/unfold and popup keyboard behavior at 640px and normal viewport widths across all three editors. The broad supported-look screenshot baseline, Office/LibreOffice corpus, and provider deployment acceptance remain outstanding.
- 2026-11-15: The new Vite production-worker regression confirms Start Center and all three standalone editor entry assets are precached and each editor initializes after Chromium is taken offline. This verifies offline launch/cache coverage, not offline document-edit/save acceptance or browser install UX; those checks and the screenshot/corpus/live-provider gates remain open.

- [ ] Chromium screenshots of the Start Center and all three editors match the current baseline in the supported looks and representative window sizes, except for the explicitly removed Recent Files UI and added file-provider entries.
- [ ] Real keyboard and pointer workflows remain functional in standalone, suite, and embedded modes.
- [ ] Existing document engine unit, round-trip, corpus, and Office/LibreOffice compatibility checks pass without new format regressions.
- [ ] Local, WebDAV, custom provider, URL, and S3 host workflows pass using controlled test providers.
- [ ] Production assets load without console errors, policy violations, or unconfigured network requests.
- [ ] PWA installation and offline local editing work from the production build.
