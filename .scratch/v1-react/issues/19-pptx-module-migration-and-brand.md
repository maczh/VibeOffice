# 19: PPTX full module migration and naming

**What to build:** Migrate every module in `public/lectern/js/` into `pptx/`; implement UI and interaction modules as typed React components/hooks, keep presentation and OOXML algorithms as typed non-UI services, and remove runtime dependence on the legacy Lectern directory.

**Blocked by:** 00: Direct editor-engine mount foundation; 01: Vite workspace and editor package contract; 02: Shared UI interface foundation; 16: Shared UI module migration to React.

**Status:** ready-for-agent

## Acceptance

- [ ] Every source module under `public/lectern/js/` has an accounted-for PPTX package implementation or an explicit shared-service destination.
- [ ] Slide canvas, text editing, panes, dialogs, commands, menus, ribbon, slideshow controls, and interaction handlers are React-owned.
- [ ] PPTX package, standalone entry, and suite run without loading `public/lectern/` or using an iframe.
- [ ] Web-owned Lectern names are migrated to PPTX in identifiers, storage keys, custom MIME types, visible strings, docs, tests, paths, and generated outputs.
- [ ] PPTX/PresentationML technical identifiers and Microsoft PowerPoint interoperability terminology remain standards-correct.
- [ ] Existing PPTX visual, interaction, local/remote file, and OOXML compatibility suites pass.
