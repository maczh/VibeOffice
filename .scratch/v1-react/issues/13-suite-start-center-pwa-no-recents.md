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

**Remaining:** Depends on Wordx/Excelx/PPTX full-parity tickets and the all-editor remote workflow ticket. The suite must compose the same-page mounted React editor packages; iframe-hosted engines are not acceptable.
