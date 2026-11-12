# 02: Shared React interface foundation

**What to build:** The three editors share React-based menus, dialogs, ribbon controls, looks, and language behavior that match the current VibeOffice interface.

**Blocked by:** 01: Vite workspace and editor package contract.

**Additional blocker:** 00: Direct editor-engine mount foundation.

**Status:** in-progress

**Progress:** Added a shared React command bar to all three editor packages. It exposes New/Open/Save plus look/language selectors. The editor runtime must mount in the host document directly; do not route commands through an iframe or cross-window bridge. The legacy app ribbon and menus remain until their individual React migration tickets are completed.

2026-11-12 update: Look and Language are now rendered by one exported `AppearanceControls` component shared between the Start Center and all editor packages. It retains existing translation/catalog choices and routes look/language changes into the direct-mounted engine. Workspace package tests (18/18), typecheck, all package/standalone/suite builds, the production React suite browser test (1/1), and `git diff --check` pass. Legacy app ribbon, menus, and dialogs remain outstanding.

2026-11-12 update: The shared React chrome now also owns the direct-host command-search combobox, backed by the existing engine's command matching/catalog API. Production Chromium verifies that searching `bold` yields the engine's Bold command and that the legacy duplicate is hidden in direct-host mode. Classic standalone pages retain their existing search control. Ribbon/menu/dialog migration remains incomplete.

2026-11-12 update: Direct-host About and message-box requests now render in shared React `AboutDialog` and `MessageDialog` components; both resolve the legacy command promises with their original result semantics. Production coverage exercises all three editor types. Feature-specific app dialogs, the Ribbon, and menus remain classic UI.

2026-11-12 update: Shared React chrome now also owns direct-host toast notices and busy-progress overlays. Toast action behavior and progress lifecycle are checked in production Chromium for Wordx, Excelx, and PPTX. Ribbon, menus, and feature-specific dialogs remain classic UI.

2026-11-12 update: Editor `data-tip` tooltips are now rendered by the shared React package in the mounted editor host, preserving the delay, translation exceptions, screen-edge placement, and legacy hide hook. The production suite test verifies all three editors and confirms no duplicate classic tooltip layer. Menus, Ribbon, and general feature dialogs remain classic UI.

2026-11-12 update: Common text prompts used by editor features now render through a shared React modal in direct-host mode, returning entered text or null on cancellation with the old Promise contract. Production Chromium verifies both paths across all three editors. Classic app pages continue using their existing dialogs.

2026-11-12 update: Protected-file password prompts and external hyperlink warning/copy dialogs now use shared React components in direct-host mode. They retain retry/cancel/input behavior and safe copy-only link handling; production Chromium exercises them in all three packages. Compatibility warnings, app feature dialogs, Ribbon, and menus are still to migrate.

2026-11-12 update: Save-time compatibility warnings now render in a shared React dialog. The existing loss service remains authoritative; Continue returns the original pending loss records and Cancel/native close prevents saving. Chromium coverage verifies grouping and outcomes in every app package. Feature-specific dialogs, Ribbon, and menus remain outstanding.

2026-11-12 update: The shared color palette popover now renders in React for direct-host editors, including standard/recent/theme swatches, keyboard movement, outside dismissal, effect action, and handoff to More Colors. Production tests verify selection and no duplicate classic palette in all editor packages; the custom-color dialog and Ribbon/menu migration remain outstanding.

2026-11-12 update: The More Colors dialog is now also React-owned: standard hue palette, grayscale row, native picker, RGB/hex editing, live previews, and OK/Cancel result semantics. Production tests exercise custom hex entry after opening from the palette. Ribbon, menus, and remaining feature dialogs remain to migrate.

2026-11-12 update: `ui.openMenu()` now routes direct-host standard menus through a shared React layer with nested command menus, checked/disabled states, access keys, outside/Escape close, and original command callbacks. Custom app menu DOM is mounted into a React-owned menu slot and remains migration work. Chromium exercises keyboard submenus and selection across all three packages; the menu bar was migrated to React in the subsequent slice, while Ribbon/toolbars remain to do.

2026-11-12 update: The menu bar now also renders via React portal into its existing app chrome location, keeping the legacy `ui.menuBar()` consumer API and access-key/menu switching. Production suite checks Wordx, Excelx, and PPTX mount and open this React bar. The app Ribbon/toolbars, custom menu bodies, and remaining feature dialogs are still not migrated.

2026-11-12 update: Generic direct-host `ui.dialog()` calls now use a React-owned common dialog shell while retaining legacy feature-body DOM, button/Promise/callback results, keyboard and focus semantics. React menus, tooltips, and this shell render under the engine's scoped style root, retaining each app's look without style leakage into a host document. Production suite assertions verify the shell/content contract and computed Office positioning across all apps; Ribbon/toolbars and feature-specific dialog bodies remain to migrate.

2026-11-12 update: Direct-host Paper 2016 Ribbon structure, tabs, tab selection, contextual tabs, groups, ordinary command buttons, and collapse/pin now render from each app's existing Ribbon spec in shared React. Split/drop/combo and custom controls still use legacy helpers; classic standalone pages keep their classic Ribbon implementation. Production Chromium switches Wordx, Excelx, and PPTX to Paper 2016 and verifies visible React tabs/commands without a duplicate. Root-level `data-layout` style scoping and mount-owned `L.onLook()` cleanup were fixed after the browser uncovered both. Builds and the suite/standalone/classic regressions pass.

2026-11-12 update: Shared toolbar rows and ordinary command buttons now also render in React for directly mounted editors, with live command state and the existing engine command callbacks; custom DOM controls remain in React-managed slots. Wordx, Excelx, and PPTX production coverage confirms React toolbar rows/commands are present. One first-run remote Save As timeout did not reproduce in the isolated suite rerun; continue watching this path.

2026-11-12 update: Status-bar message text is now rendered by a shared React status component in all three direct-host editors; the classic path stays unchanged. Engine status calculations remain in the legacy application modules and are bridged as text, so the status areas and the rest of each application UI are still not fully migrated.

2026-11-12 update: Ribbon split/drop command buttons now have React-owned markup and event handlers, including the split primary action and menu arrow. Existing app menu factories still feed the shared React menu layer. Production coverage opens the split menu in each app; app-specific custom controls and menu bodies are still compatibility slots.

2026-11-12 update: Common editable combo inputs (font, size, zoom, and feature list boxes) now mount as React input/dropdown controls for direct-host editors. Their existing typed-at-the-boundary option, value, note, preview, and change callbacks remain engine-backed; dropdown entries render through the shared React menu. Production Chromium verifies all three apps and font preview styling; classic pages remain unchanged.
2026-11-12 update: Direct-host shared select, spinner, checkbox, and radio controls now render as React elements while their established engine-facing APIs remain usable (`value`, `selectedIndex`, `.input`, `.get()`, `.set()`, `.setDisabled()`, and callbacks). Three-app production Chromium verifies the compatibility surface. DOM-built feature dialog bodies and `make()` controls remain unfinished.
2026-11-12 update: `ui.tabs()` now uses shared React for its tab bar, accessible selected state and keyboard movement, while React selection still invokes the existing body visibility and `onShow` callbacks. App-created panel DOM remains in place pending feature-dialog migration.
2026-11-12 update: `ui.button()` now portals translated label markup through the shared React layer into the original native button, preserving old references, attributes and event listeners. Its content setters update the React-managed label node instead of detaching it; app DOM composition remains legacy.
2026-11-12 update: `ui.field()` and `ui.group()` now delegate shared translated label/legend content to React while preserving their semantic native wrapper nodes and app-created content.

The production suite also verifies the message-box title/content/icon, sanitized access-key labels, primary button semantics, explicit selection indexes, and the legacy Escape/cancel result across Wordx, Excelx, and PPTX.

- [x] Shared React UI pieces can be used by all three editor packages without copying shared code.
- [x] Classic, Paper, and Paper 2016 looks retain their current chrome colors, typography, and layout behavior.
- [x] Existing command, menu, dialog, keyboard, and focus behavior has a React-facing integration path.
- [x] Existing language selection and translation lookup continue to work; new strings use the English catalog mechanism.
- [x] The shared interface renders in the host fixture and in the standalone Vite shell.
- [x] A shared React command bar mounts consistently for Wordx, Excelx, and PPTX, with commands routed into the hosted engine.
- [x] Look and language selectors are part of the React editor chrome and synchronize to the hosted engine.
- [ ] Legacy app menus, dialogs, and ribbon controls are migrated to shared React components while preserving commands, keyboard, and focus behavior.
