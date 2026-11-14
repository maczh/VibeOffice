# 13: Suite Start Center, Recent Files removal, and PWA behavior

**What to build:** Users can launch all three React editors from a suite Start Center that defaults to Create File and retains the existing installable/offline behavior without Recent Files.

**Blocked by:** 00: Direct editor-engine mount foundation; 05: Wordx full feature and format parity; 07: Excelx full feature and format parity; 09: PPTX full feature and format parity; 12: Remote document workflows in all editors; 20: Web-wide legacy brand name cleanup.

**Status:** in-progress

- [x] The Start Center composes the same Wordx, Excelx, and PPTX packages that can run independently.
- [x] The Start Center opens on the existing Create File gallery with blank and template choices.
- [x] Local and remote opening can route a file to its matching editor.
- [x] Recent Files cards, settings, clear controls, app menus, task-pane links, and recent-document navigation are absent.
- [x] Local unsaved draft recovery remains available independently of Recent Files.
- [x] Existing browser Recent Files data is left untouched and is not read by the new product.
- [x] The static production build retains installability, service-worker caching, and offline local editing.

**Remaining:** Depends on Wordx/Excelx/PPTX full-parity tickets, the all-editor remote workflow ticket, and final web-wide naming cleanup. The suite now composes the direct-mounted React editor packages in the same page with zero iframe elements; it is not acceptance-complete until those remaining feature and acceptance tickets pass.

2026-11-12: The React Vite build now derives PWA shortcut and screenshot routes/names for Wordx, Excelx, and PPTX while preserving the legacy public site's manifest until that entry point is migrated. The existing public service-worker suite passes (2/2), and the suite manifest is asserted through its production browser test.
- 2026-11-13: Removed the dormant Recent Files link injection from all three legacy Getting Started task-pane sources as well as the React task-pane surfaces. New browser regressions open each app's native pane and verify Open/Create remain available with no recent-file text or link; Wordx, Excelx and PPTX checks pass 3/3. Existing browser storage is not read or deleted. Suite acceptance remains gated by the app parity and brand-cleanup tickets listed above.
- 2026-11-13: Cross-app full quick after the pane cleanup and Excelx React task-pane migration passes Node 79, React suite 1, smoke 4, PPTX UI 41, Excelx UI 45, and Wordx UI 54. The legacy Recent Files recovery route remains a distinct compatibility surface to remove from the React host while keeping local draft recovery.
- 2026-11-13: Removed and explicitly blocked the legacy Recent Files recovery route in all three direct-mounted React editors; isolated local draft recovery remains independent. Production React coverage asserts that route activation is blocked and the legacy `vibeoffice` Recent Files database is not opened. Full quick passes Node 79, React suite 1, smoke 4, PPTX UI 41, Excelx UI 45, and Wordx UI 54; diff check clean. Ticket remains gated on remaining parity, remote-workflow and branding tickets.
- 2026-11-13: Removed the Recent Files view, Settings/clear controls, File-menu entries in all three classic app shells, the PPTX Options preference, and the old Start Center copy. The classic Start Center now opens directly on Create File; browser tests assert its UI and each app's menu omit Recent Files. Existing IndexedDB data is not deleted; React-host draft recovery remains isolated. Full quick validation pending.
- 2026-11-13: Recent Files removal regression passes full quick: Node 92, React suite 1, smoke 4, PPTX UI 41, Excelx UI 45, and Wordx UI 54. The English i18n catalog was gathered and checked (7/7); `git diff --check` is clean. Legacy stored browser entries are deliberately left untouched.
- 2026-11-14: Fixed the React Create File gallery's generated-template integration: the suite build/runtime remaps legacy catalog keys and preview URLs to the new editor routes. Production Chromium verifies Wordx/Excelx/PPTX gallery card counts and successful fetches for all 25 preview images. The workspace suite passes 21/21; React suite 1/1. Overall acceptance remains gated by the app parity and brand tickets.
