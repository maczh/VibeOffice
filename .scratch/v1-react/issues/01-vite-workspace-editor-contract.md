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

## Comments

- 2026-11-12: The packages now expose a typed same-document direct-engine mount bridge and browser coverage confirms clean unmounts for all three editors. The legacy engines are still loaded from their public HTML/script bundles and share mutable global state, so the acceptance items remain unchecked until those engine interfaces and lifecycle boundaries are genuinely isolated.
- 2026-11-12: Updated the independent host fixture to select and mount one editor at a time, matching the current single-engine-per-document boundary. Chromium verified Wordx → Excelx → PPTX → Wordx switching with exactly one mounted engine root at each step and no console errors. The fixture wraps its controls in the shared `VibeOfficeProvider`; full engine isolation and moving the legacy runtime into package source remain outstanding.
- 2026-11-12: The Vite React suite and standalone packages now expose new-brand engine routes. Added a production browser regression which serves each package's built `standalone/` directory independently and confirms its React shell plus corresponding document engine mount without iframe or external requests (3/3). Suite package/browser checks and standalone builds pass; legacy source modules still need migration into the package trees.
