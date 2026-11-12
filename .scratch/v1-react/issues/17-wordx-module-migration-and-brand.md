# 17: Wordx full module migration and naming

**What to build:** Migrate every module in `public/quire/js/` into `wordx/`; implement UI and interaction modules as typed React components/hooks, keep document and OOXML algorithms as typed non-UI services, and remove runtime dependence on the legacy Quire directory.

**Blocked by:** 00: Direct editor-engine mount foundation; 01: Vite workspace and editor package contract; 02: Shared React interface foundation; 16: Shared UI module migration to React.

**Status:** in-progress

## Acceptance

- [ ] Every source module under `public/quire/js/` has an accounted-for Wordx package implementation or an explicit shared-service destination.
- [ ] Editor UI, dialogs, panes, commands, menus, ribbon, proofing controls, and interaction handlers are React-owned.
- [ ] Wordx package, standalone entry, and suite run without loading `public/quire/` or using an iframe.
- [ ] Web-owned Quire names are migrated to Wordx in identifiers, storage keys, custom MIME types, visible strings, docs, tests, paths, and generated outputs.
- [ ] DOCX/OOXML technical identifiers and Microsoft Word interoperability terminology remain standards-correct.
- [ ] Existing Wordx visual, interaction, local/remote file, and OOXML compatibility suites pass.

## Comments

- 2026-11-12: Rebranded capitalized Quire product references throughout `public/quire/js/` to Wordx, including visible copy, code comments, metadata, About, the default sample, and defaults. Preserved lowercase legacy route/host API IDs and the one-document desktop title/edition-name compatibility. Added a regression for the Wordx sample. i18n gather + tests pass (7/7), full Wordx interaction UI passes 52/52, smoke and all workspace builds pass. Lowercase old identifiers/routes remain, the package still loads the classic engine runtime, and the app UI is not yet React-owned.
- 2026-11-12: Added the first app-specific React feature-control layer: Wordx `G.listBox()` now portals option rows, selection styling, and per-item presentation into a React listbox while preserving the helper's synchronous `get()`, `set()`, `setItems()`, filter-input, and callback behavior. The production suite checks selection, style, updates, and callback/value synchronization. Package typecheck/build and the three-app production browser suite pass. The input, surrounding dialog composition, and most of the Wordx modules remain classic DOM and the ticket stays in progress.
