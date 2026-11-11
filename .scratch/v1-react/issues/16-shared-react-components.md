# 16: Shared UI module migration to React

**What to build:** Migrate every shared user-interface component in `public/common/` into the shared React workspace package, with typed APIs and one implementation shared by Wordx, Excelx, PPTX, and the suite.

**Scope boundary:** Keep non-UI algorithms and browser/document services as ordinary modules; shared UI (ribbon, menus, dialogs, tooltips, search, overlays, look/language controls, and shared React hooks/providers) must be React-owned. No app may load a duplicate legacy UI implementation.

**Blocked by:** 01: Vite workspace and editor package contract; 00: Direct editor-engine mount foundation.

**Status:** ready-for-agent

## Acceptance

- [ ] Every user-facing shared component currently in `public/common/` has a React implementation in the shared package.
- [ ] All three standalone package entries and the suite use the same shared components.
- [ ] Existing looks, language selection, keyboard/pointer behavior, accessibility, and CSP constraints remain valid.
- [ ] Shared services remain single-sourced and are not copied into editor packages.
- [ ] No migrated shared UI is loaded from `public/common/*.js` at runtime.
- [ ] Browser regression, shared package tests, typecheck, production build, and `git diff --check` pass.
