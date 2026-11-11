# 06: Excelx core editor

**What to build:** Users can independently run or embed Excelx to create, open, edit, calculate, and save a spreadsheet with the current core editing experience.

**Blocked by:** 00: Direct editor-engine mount foundation; 01: Vite workspace and editor package contract; 02: Shared React interface foundation; 03: Local file access and independent draft recovery.

**Status:** blocked

- [ ] The Excelx package mounts a complete editor through React in both its standalone entry and the host fixture.
- [ ] Users can create a blank workbook or use the existing sample/start behavior.
- [ ] Cell selection, editing, keyboard navigation, clipboard, undo, and redo behave as they do in Ledger.
- [ ] Existing formula evaluation and recalculation remain powered by the Ledger engine.
- [ ] Local file open, save, dirty-state prompts, and draft recovery work in Excelx.
- [ ] The core editor UI matches the corresponding current Ledger screens in Chromium.
