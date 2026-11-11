# 18: Excelx full module migration and naming

**What to build:** Migrate every module in `public/ledger/js/` into `excelx/`; implement UI and interaction modules as typed React components/hooks, keep workbook, OOXML, and calculation algorithms as typed non-UI services, and remove runtime dependence on the legacy Ledger directory.

**Blocked by:** 00: Direct editor-engine mount foundation; 01: Vite workspace and editor package contract; 02: Shared React interface foundation; 16: Shared UI module migration to React.

**Status:** ready-for-agent

## Acceptance

- [ ] Every source module under `public/ledger/js/` has an accounted-for Excelx package implementation or an explicit shared-service destination.
- [ ] Spreadsheet UI, formula bar, grid interactions, dialogs, panes, commands, menus, ribbon, and editors are React-owned.
- [ ] Excelx package, standalone entry, and suite run without loading `public/ledger/` or using an iframe.
- [ ] Web-owned Ledger names are migrated to Excelx in identifiers, storage keys, visible strings, docs, tests, paths, and generated outputs.
- [ ] XLSX/SpreadsheetML technical identifiers and Microsoft Excel interoperability terminology remain standards-correct.
- [ ] Existing Excelx visual, formula, interaction, local/remote file, and OOXML compatibility suites pass.
