# 01: Vite workspace and editor package contract

**What to build:** A React host can mount a typed VibeOffice editor package, and the Vite workspace can build the standalone and distributable package entry points.

**Blocked by:** None (can start immediately).

**Status:** in-progress

**Blocked by:** 00: Direct editor-engine mount foundation.

- [x] The workspace contains the requested Wordx, Excelx, and PPTX package boundaries plus one shared package.
- [x] New React-facing APIs use strict TypeScript, explicit public types, and generated declaration files.
- [x] React and React DOM are peer dependencies of the editor packages.
- [x] A minimal host fixture can import and mount each package without relying on the VibeOffice Start Center.
- [x] Vite can build the package artifacts and standalone entry points without adding runtime server requirements.
- [ ] The existing JavaScript document engines remain available through typed interfaces without an iframe or nested browsing context.
- [ ] Each editor engine can mount and dispose inside a React-owned host container without an iframe or nested browsing context.
