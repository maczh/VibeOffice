# 08: PPTX core editor

**What to build:** Users can independently run or embed PPTX to create, open, edit, and save a presentation with the current core slide editing experience.

**Blocked by:** 00: Direct editor-engine mount foundation; 01: Vite workspace and editor package contract; 02: Shared React interface foundation; 03: Local file access and independent draft recovery.

**Status:** blocked

- [ ] The PPTX package mounts a complete editor through React in both its standalone entry and the host fixture.
- [ ] Users can create a blank presentation or use the existing sample/start behavior.
- [ ] Slide selection, text editing, selection, shape placement, keyboard commands, undo, and redo behave as they do in Lectern.
- [ ] Local file open, save, dirty-state prompts, and draft recovery work in PPTX.
- [ ] The core editor UI matches the corresponding current Lectern screens in Chromium.
