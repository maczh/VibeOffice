# 04: Wordx core editor

**What to build:** Users can independently run or embed Wordx to create, open, edit, undo, and save a Word document with the current core editing experience.

**Blocked by:** 00: Direct editor-engine mount foundation; 01: Vite workspace and editor package contract; 02: Shared React interface foundation; 03: Local file access and independent draft recovery.

**Status:** in-progress

**Architecture constraint:** Mount the existing Quire engine directly into the Wordx React host. Iframes and other nested browsing contexts are not permitted. The engine must support an explicit mount/dispose lifecycle without relying on a page-global `#app` or document-wide singleton DOM.

- [ ] The Wordx package mounts a complete editor through React in both its standalone entry and the host fixture.
- [ ] Users can create a blank document or use the existing sample/start behavior.
- [ ] Text selection, caret placement, typing, keyboard commands, undo, and redo behave as they do in Quire.
- [ ] The editor uses the existing Quire document engine and retains the existing page layout behavior.
- [ ] Local file open, save, dirty-state prompts, and draft recovery work in Wordx.
- [ ] The core editor UI matches the corresponding current Quire screens in Chromium.
