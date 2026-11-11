# 05: Wordx full feature and format parity

**What to build:** Wordx supports the rest of Quire's editing, layout, review, proofing, drawing, and file workflows with existing Office compatibility.

**Architecture constraint:** Extend the direct host-container integration from ticket 04. Wordx must not use an iframe or another nested browsing context.

**Blocked by:** 00: Direct editor-engine mount foundation; 04: Wordx core editor; 17: Wordx full module migration and naming.

**Status:** ready-for-agent

- [ ] Existing Quire formatting, page setup, tables, sections, fields, drawing, review, proofing, and other editing commands remain available.
- [ ] Existing supported input and output formats remain available with their current behavior.
- [ ] Existing templates and creation flows remain available.
- [ ] Saved files retain the existing preservation and loss-reporting behavior.
- [ ] Existing Wordx-visible workflows operate through the React package in standalone and embedded use.
