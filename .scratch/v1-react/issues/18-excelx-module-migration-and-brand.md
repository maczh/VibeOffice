# 18: Excelx full module migration and naming

**What to build:** Migrate every module in `public/ledger/js/` into `excelx/`; implement UI and interaction modules as typed React components/hooks, keep workbook, OOXML, and calculation algorithms as typed non-UI services, and remove runtime dependence on the legacy Ledger directory.

**Blocked by:** 00: Direct editor-engine mount foundation; 01: Vite workspace and editor package contract; 02: Shared React interface foundation; 16: Shared UI module migration to React.

**Status:** in-progress

## Acceptance

- [ ] Every source module under `public/ledger/js/` has an accounted-for Excelx package implementation or an explicit shared-service destination.
- [ ] Spreadsheet UI, formula bar, grid interactions, dialogs, panes, commands, menus, ribbon, and editors are React-owned.
- [ ] Excelx package, standalone entry, and suite run without loading `public/ledger/` or using an iframe.
- [ ] Web-owned Ledger names are migrated to Excelx in identifiers, storage keys, visible strings, docs, tests, paths, and generated outputs.
- [ ] XLSX/SpreadsheetML technical identifiers and Microsoft Excel interoperability terminology remain standards-correct.
- [ ] Existing Excelx visual, formula, interaction, local/remote file, and OOXML compatibility suites pass.

## Comments

- 2026-11-12: Added an Excelx-package React formula-bar layer. Its toolbar buttons and visible formula textarea are React-rendered; a hidden legacy textarea remains as a compatibility edit surface. Keyboard input, selection synchronization, cancel, commit, and formula evaluation continue through the existing editor contract. The production React-suite regression types and commits `=1+2` from the React textarea and verifies the resulting cell value and Enter-to-next-cell behavior. This advances the UI migration but does not complete the Excelx module, grid, command, pane, or engine migration.
- 2026-11-12: Extended the Excelx React editing chrome with a React Name Box input and names dropdown button. The hidden adapter preserves `goToRef`, named-range creation, menu contents, active-edit commit behavior, and selection refresh. The production suite verifies the Name Box routes `C3` through the existing selection command. Formula-bar keyboard Enter, Enter button, Cancel button, and calculation checks all pass in the same suite.
- 2026-11-12: Moved the Excelx worksheet tab rail items into `excelx/src/sheet-tabs.tsx`. React now renders visible tabs, grouped/active selection states, inline rename, context-menu handoff, and drag reorder/copy gestures; the legacy app controller remains the source of sheet snapshots and workbook operations. The production suite verifies model-to-tab state, sheet activation, rename, and drag-to-end reorder. Tab arrow/resize controls and the workbook/app controller are still classic, so Excelx migration remains in progress.
- 2026-11-12: React now also renders the sheet-scroll arrows. Added a classic-page fallback so the generated static Excelx shell continues to display its original formula bar, Name Box, sheet tabs, and navigation when the React package is not mounted; React host mounting alone activates the React-owned controls. The classic UI regression verifies those controls remain visible and all four navigation arrows are present. React production-suite and static smoke checks pass.
- 2026-11-12: Continued the legacy-engine lifecycle seam while the full React UI migration remains in progress. Ledger AutoFilter, comment editor, Alt+Down pick list, Print Preview, and touch long-press now bind document-level or delayed work to the active editor mount and clean it up on disposal; Print Preview also closes its overlay without restoring focus during unmount. Added browser regressions for all five cases; the complete Ledger UI + smoke suite passes 47/47. Browser document-title suffixes now use Excelx, Wordx, and PPTX (the excluded one-document desktop host keeps legacy title suffixes); Wordx/Excelx HTML generator metadata was updated. Rebranded Excelx's formula, search, repair, print, picture, and dependency messages and its default author label; `node tools/i18n.mjs gather` updated only the English catalog and `tests/i18n.test.mjs` passes 7/7. The production React-suite regression covers all three browser title suffixes, and builds pass. The actual app ribbon/dialog/pane UI remains classic JavaScript and the modules are not yet fully migrated.
