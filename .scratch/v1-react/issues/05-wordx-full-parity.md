# 05: Wordx full feature and format parity

**What to build:** Wordx supports the rest of Quire's editing, layout, review, proofing, drawing, and file workflows with existing Office compatibility.

**Architecture constraint:** Extend the direct host-container integration from ticket 04. Wordx must not use an iframe or another nested browsing context.

**Blocked by:** 00: Direct editor-engine mount foundation; 04: Wordx core editor; 17: Wordx full module migration and naming.

**Status:** in-progress

## Comments

- 2026-11-13: Started React-package workflow parity checks. Production Chromium opens a local Markdown `.md` file in the direct-host Wordx package and verifies the Markdown importer and rendered canvas content. This validates only the local Markdown path; remaining input/output formats, template flows, save/preservation behavior, and standalone/embedded workflows remain open.
- 2026-11-13: Expanded the local Markdown acceptance path with a production-package GitHub-flavour document containing a checked task, pipe table, strikethrough, and fenced code. Also opened a local Markdown `.zip`, verified its packaged GIF resolves, and round-tripped the package while preserving an unrelated file and the Markdown image reference. React suite passes; other format, template, save/preservation, and standalone/embedded flows remain open.

- [ ] Existing Quire formatting, page setup, tables, sections, fields, drawing, review, proofing, and other editing commands remain available.
- [ ] Existing supported input and output formats remain available with their current behavior.
- [ ] Existing templates and creation flows remain available.
- [ ] Saved files retain the existing preservation and loss-reporting behavior.
- [ ] Existing Wordx-visible workflows operate through the React package in standalone and embedded use.
