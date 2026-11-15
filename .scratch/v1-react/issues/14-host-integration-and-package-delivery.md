# 14: Host integration example and package delivery

**What to build:** A React application outside the suite can install and embed each complete editor package using its public types, styles, and storage-provider contract.

**Blocked by:** 00: Direct editor-engine mount foundation; 13: Suite Start Center, Recent Files removal, and PWA behavior; 16: Shared UI module migration to React; 17: Wordx full module migration and naming; 18: Excelx full module migration and naming; 19: PPTX full module migration and naming; 20: Web-wide legacy brand name cleanup.

**Status:** in-progress

The React packages now mount their current editor engines directly in the consumer's document without iframes. Their engines and much of the shared/app UI still come from the legacy public script bundles, so package delivery remains dependent on the full module migrations and final brand audit.

- 2026-11-15: React direct-host core identity is driven by the mounted package (`L.APP_ID`/`L.APP_NAME`/`L.APP` become `wordx`/Wordx, `excelx`/Excelx, or `pptx`/PPTX on package switches), independent of consumer `<html>` metadata. External-host production coverage also verifies qps language isolation. Targeted package/workspace/suite tests pass 26/26, typecheck, package/standalone/suite builds, full quick (all 11 groups) and diff check pass. The compatibility loader and incomplete shared/app UI migrations still block package delivery completion.
- 2026-11-15: Full `node tools/quick.mjs` after the title callback change passes all 11 groups (Node 170; React hosts 4; React suite 1; host package 1; smoke 4; drag 48; classic focus 40; PPTX 43; Excelx 46; Wordx 54). All package, standalone, suite builds and diff checks pass. Legacy runtime/module migration remains a prerequisite for ticket completion.
- 2026-11-15: Added optional `onTitleChange(title)` to each public editor component so host applications can update their own browser title without the embedded engine mutating `document.title`. The suite opts in; the consumer-style fixture verifies each package reports a title while preserving the host title. Production browser checks and workspace/package builds pass; full quick follows.

- 2026-11-15: Fixed the in-repo external-host Vite server so its legacy asset middleware forwards missing `/wordx/src/*`, `/excelx/src/*`, and `/pptx/src/*` requests to Vite's package resolver instead of returning an ENOENT 500; ignored `tmp/`, `.git/`, and `node_modules/` in the host fixture's file watcher to avoid Chromium-profile watcher exhaustion. The source alias returns HTTP 200 and workspace tests pass (24/24), including a regression for middleware fallthrough. Full host-driven Open/Save UI verification remains open, as do the underlying singletons and legacy UI modules.
- 2026-11-15: The consumer-style host fixture now injects a revision-aware in-memory provider into all exported Wordx/Excelx/PPTX components. Real Chromium opens a provider file and saves it through the host callback in all three independently mounted editors, confirming each current revision reaches `save`; the fixture keeps engine config/callbacks stable and advances the provider revision between reads. The run also caught and fixed a missing React child-key warning in PPTX's Getting Started pane. Host integration passes (1/1) with no console errors, and workspace typecheck passes; package readiness remains blocked by the broader direct-engine singleton and full UI-module migration.

- [x] Each editor package builds independently with its public React component, styles, and TypeScript declarations.
- [x] React and React DOM are declared as peer dependencies.
- [x] A consumer example mounts each editor without rendering the suite Start Center.
- [x] The example demonstrates local file operations and a host-provided remote storage adapter.
- [x] Public API documentation describes configuration, callbacks, styles, provider behavior, and deployment CSP/CORS requirements.
- [x] Package artifacts do not require VibeOffice backend services or runtime development-server dependencies.
