# 09: PPTX full feature and format parity

**What to build:** PPTX supports Lectern's remaining presentation, design, chart, diagram, slideshow, and file workflows with existing Office compatibility.

**Blocked by:** 00: Direct editor-engine mount foundation; 08: PPTX core editor; 19: PPTX full module migration and naming.

**Status:** ready-for-agent

- [ ] Existing presentation themes, designs, tables, charts, IntelliArt, drawings, comments, notes, and other features remain available.
- [ ] Existing slideshow behavior and slide navigation remain available.
- [ ] Existing supported input and output formats remain available with their current behavior.
- [ ] Existing templates and creation flows remain available.
- [ ] Saved files retain the existing preservation and loss-reporting behavior.
- [ ] Existing PPTX-visible workflows operate through the React package in standalone and embedded use.

2026-11-18 follow-up: The React package now provides the PPTX editor's foundational canvas/ruler DOM instead of asking the engine to create those UI containers; engine drawing/model services bind into the React surface. Production Chromium asserts the node structure and binding identity. This is an incremental React ownership improvement; feature/format parity acceptance remains open.
