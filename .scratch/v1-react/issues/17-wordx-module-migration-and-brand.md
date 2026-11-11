# 17: Wordx full module migration and naming

**What to build:** Migrate every module in `public/quire/js/` into `wordx/`; implement UI and interaction modules as typed React components/hooks, keep document and OOXML algorithms as typed non-UI services, and remove runtime dependence on the legacy Quire directory.

**Blocked by:** 00: Direct editor-engine mount foundation; 01: Vite workspace and editor package contract; 02: Shared React interface foundation; 16: Shared UI module migration to React.

**Status:** ready-for-agent

## Acceptance

- [ ] Every source module under `public/quire/js/` has an accounted-for Wordx package implementation or an explicit shared-service destination.
- [ ] Editor UI, dialogs, panes, commands, menus, ribbon, proofing controls, and interaction handlers are React-owned.
- [ ] Wordx package, standalone entry, and suite run without loading `public/quire/` or using an iframe.
- [ ] Web-owned Quire names are migrated to Wordx in identifiers, storage keys, custom MIME types, visible strings, docs, tests, paths, and generated outputs.
- [ ] DOCX/OOXML technical identifiers and Microsoft Word interoperability terminology remain standards-correct.
- [ ] Existing Wordx visual, interaction, local/remote file, and OOXML compatibility suites pass.
