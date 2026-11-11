# 03: Local file access and independent draft recovery

**What to build:** Users can open and save local documents in a React editor and recover unsaved work without depending on Recent Files.

**Blocked by:** 01: Vite workspace and editor package contract; 02: Shared React interface foundation.

**Status:** complete

- [x] A user can choose a local file, returning its File and (when available) writable handle for the editor adapter.
- [x] Browsers with writable file handles can save to the user-selected file after a user action.
- [x] Other supported browsers can save by downloading the generated file.
- [x] Unsaved document drafts have a local-only IndexedDB store separate from Recent Files.
- [x] Shared Save / Discard / Cancel guard is available for editor adapters; failed saves preserve the document and cancel prevents replacement.
- [x] Existing Recent Files data in browser storage is not automatically deleted or read by the new implementation.

App-specific engine routing and wiring the confirmation guard are performed by the editor-core tickets; this ticket provides their local file/draft boundary.
