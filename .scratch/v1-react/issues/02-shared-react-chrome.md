# 02: Shared React interface foundation

**What to build:** The three editors share React-based menus, dialogs, ribbon controls, looks, and language behavior that match the current VibeOffice interface.

**Blocked by:** 01: Vite workspace and editor package contract.

**Additional blocker:** 00: Direct editor-engine mount foundation.

**Status:** in-progress

**Progress:** Added a shared React command bar to all three editor packages. It exposes New/Open/Save plus look/language selectors. The editor runtime must mount in the host document directly; do not route commands through an iframe or cross-window bridge. The legacy app ribbon and menus remain until their individual React migration tickets are completed.

- [x] Shared React UI pieces can be used by all three editor packages without copying shared code.
- [x] Classic, Paper, and Paper 2016 looks retain their current chrome colors, typography, and layout behavior.
- [x] Existing command, menu, dialog, keyboard, and focus behavior has a React-facing integration path.
- [x] Existing language selection and translation lookup continue to work; new strings use the English catalog mechanism.
- [x] The shared interface renders in the host fixture and in the standalone Vite shell.
- [x] A shared React command bar mounts consistently for Wordx, Excelx, and PPTX, with commands routed into the hosted engine.
- [x] Look and language selectors are part of the React editor chrome and synchronize to the hosted engine.
- [ ] Legacy app menus, dialogs, and ribbon controls are migrated to shared React components while preserving commands, keyboard, and focus behavior.
