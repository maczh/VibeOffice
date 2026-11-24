# 02: Shared React interface foundation

**What to build:** The three editors share React-based menus, dialogs, ribbon controls, looks, and language behavior that match the current VibeOffice interface.

**Blocked by:** 01: Vite workspace and editor package contract.

**Additional blocker:** 00: Direct editor-engine mount foundation.

**Status:** in-progress

**Progress:** The shared React command bar, menus, dialogs, Ribbon, toolbars, color and appearance controls, search, status/toast/busy layers, and keyboard/focus behavior are shared by all three editor packages and the suite. React editor mounts install shared services without loading the legacy `common/ui.js`; all engines remain in the same host document with no iframe. Some direct-host feature-dialog bodies still use imperative engine DOM builders and remain part of the app-specific UI migration.

2026-11-15 update: Wordx and PPTX bottom view bars now route their command controls through the shared React toolbar layer, including Wordx's browse-object menu; classic launches keep their original DOM controls. The rest of the classic app Ribbon/menu/dialog runtime remains outstanding.

2026-11-15 update: The shared React toolbar now renders app-provided command, split, and dropdown buttons directly, including existing engine-backed enabled/checked state and menu callbacks. Production tests assert all three editor toolbars have no unowned legacy command buttons; the contextual task-pane Ribbon and non-toolbar view controls still use explicit DOM compatibility, and migration of the rest of the legacy Ribbon/menu/dialog runtime remains outstanding.

2026-11-15 update: Toolbar visibility for all three direct-host editors is now app-snapshot-driven and applied by the shared React toolbar lifecycle, including empty-row collapse. Production Chromium toggles Standard off/on in Wordx, Excelx, and PPTX; classic launches retain their existing imperative visibility path. Package/suite builds, typecheck, React suite, and full quick regression pass (Node 91, React 1, smoke 4, drag 48, focus 40, PPTX 41, Excelx 45, Wordx 54). Legacy shared chrome remains in progress.

2026-11-15 update: Restored Paper 2016 Ribbon's responsive group folding in React. Groups fold from the right as the host narrows, open their controls in an accessible React popup, close on outside click/Escape, restore focus on Escape, and unfold on resize. Production Chromium covers all three editors at narrow and normal widths; package builds, typecheck, React suite, and full quick (Node 91, React 1, smoke 4, drag 48, focus 40, PPTX 41, Excelx 45, Wordx 54) pass. Other legacy compatibility controls remain outstanding.

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

2026-11-24 follow-up: Moved the direct-host `ui.openMenu()`, `ui.closeMenus()`, `ui.menuBar()`, `ui.menuOpen()`, and `ui.menuKey()` bridge into `commonx/src/menu-service.ts`; the React host owns these functions while classic pages retain the DOM menu implementation. Typecheck, all package/standalone/Suite builds, React Suite (1/1), classic smoke (4/4), workspace/provider tests (32/32), and `git diff --check` pass. App-defined custom menu bodies and feature-specific dialog bodies still need React ownership.

2026-11-24 follow-up: Moved the synchronous direct-host `ui.dialog()` facade and dialog-stack/key state into `commonx/src/legacy-dialog-service.ts`; generic dialog chrome/content lifecycle remains rendered by `LegacyDialog`, and classic DOM dialogs remain the fallback. A React Suite run exposed a button-layer effect timing race when an engine API became ready before its listener; switching the listener to `useLayoutEffect` made the production suite pass. Typecheck, all package/standalone/Suite builds, React Suite (1/1), classic smoke (4/4), and `git diff --check` pass. Feature-specific dialog bodies and custom menu contents remain migration work.

2026-11-24 follow-up: Moved the direct-host `ui.button()` proxy/event adapter into `commonx/src/button-service.ts`, retaining `L.h` attribute semantics and the React portal label setters; `public/common/ui.js` now contains only the classic fallback. `ButtonLayer` registers its handler at layout-effect time to avoid losing the first engine button request. Typecheck, all package/standalone/Suite builds, React Suite (1/1, including the live setter assertion), classic smoke (4/4), and `git diff --check` pass.

2026-11-24 follow-up: Moved direct-host `ui.check()` and `ui.radio()` wrappers into `commonx/src/choice-service.ts`, preserving synchronous wrapper/input-reference and callback contracts. The shared ChoiceLayer now attaches host listeners in a layout effect so initial control requests cannot race post-paint effects. Typecheck, all builds, React Suite (1/1, covering the three editors' checkbox/radio behavior), classic smoke (4/4), and scoped `git diff --check` pass.

2026-11-24 follow-up: Moved the direct-host `ui.select()` wrapper into `commonx/src/select-service.ts`, retaining synchronous `value`, `disabled`, and `selectedIndex` proxies while React renders the native select. SelectLayer listener registration now occurs in a layout effect. Typecheck, all builds, React Suite (1/1), classic smoke (4/4), workspace/provider tests (32/32), and scoped `git diff --check` pass.

2026-11-24 follow-up: Moved direct-host `ui.spin()` wrapper/proxy into `commonx/src/spin-service.ts`, preserving synchronous input/get/set/setDisabled methods and event-backed React spinner updates. SpinLayer listeners now register during layout. Typecheck, all builds, React Suite (1/1), classic smoke (4/4), workspace/provider tests (32/32), and scoped `git diff --check` pass.

2026-11-24 follow-up: Moved direct-host `ui.tabs()`, `ui.field()`, and `ui.group()` request construction into `commonx/src/structure-service.ts`; React TabsLayer/UI content portals own the hosted labels, field legends, and tab chrome, while classic DOM fallbacks remain. Layout-phase subscriptions close the initial-request race. Typecheck, all builds, React Suite (1/1), classic smoke (4/4), workspace/provider tests (32/32), and scoped diff checks pass.

2026-11-24 follow-up: Moved direct-host `ui.combo()` bridge construction into `commonx/src/combo-service.ts`; React ComboLayer owns hosted input/menu controls, while `public/common/ui.js` retains the classic fallback. Combo requests and value refreshes now subscribe during React layout. Typecheck, all workspace/standalone/Suite builds, React Suite (1/1, including populated font options and previews), classic smoke (4/4), and `git diff --check` pass. App Ribbon custom controls, feature dialogs, and menus remain incomplete.

2026-11-24 follow-up: Moved direct-host toolbar-control and toolbar-row bridge construction into `commonx/src/toolbar-service.ts`; React toolbar rows render ordinary, split, and dropdown controls, while `forceDOM` remains an explicit compatibility route for app controls not yet migrated. Service API overrides restore safely on editor unmount. Full builds, React Suite (1/1), classic smoke (4/4), external-host test (1/1), workspace/provider tests (33/33), and `git diff --check` pass; remaining legacy Ribbon/app controls and feature dialog bodies are still outstanding.

2026-11-24 follow-up: Moved direct-host `ui.tbButton()`, `ui.tbSplit()`, `ui.tbDrop()`, and `ui.toolbar()` bridge construction into `commonx/src/toolbar-service.ts`; shared React toolbar rows and controls remain the renderer, with `forceDOM` preserved for legacy controls embedded in React toolbars. All builds, React Suite (1/1), classic smoke (4/4), workspace/provider tests (32/32), typecheck, and `git diff --check` pass.

2026-11-18 update: Paper 2016 Ribbon font, size, and zoom controls now render directly as shared React combo components in all three apps; Wordx's style selector is included. Existing engine option/value/change models are shared with classic toolbar fallbacks, while the React controls no longer call the legacy DOM combo builder. A production Chromium regression checks React ownership and a populated font value in each app. Typecheck, all app/standalone/suite builds, the React Suite, and the complete changed-file quick regression pass (Node 102; React host 4; host apps 2; host packages 1; suite 1; standalone 1; hyperlinks 3; smoke 4; painter 4; PPTX 43; textmap 9). Remaining Ribbon custom controls, menu bodies, and feature dialog bodies are still outstanding.

2026-11-18 follow-up: Removed the shared React Ribbon's generic `make()` DOM-control slot after verifying all three package Ribbon specifications use declarative React controls; classic pages retain their fallback Ribbon. The production React Suite now asserts no legacy Ribbon-control wrappers in all apps. Typecheck and all package/standalone/Suite builds pass. In the complete changed-file quick, the Node/host/standalone/hyperlink/smoke/painter/PPTX/textmap groups all passed; React Suite initially exposed a test-harness API mismatch (`page.locator` is unsupported), corrected to `page.evaluate`, after which the production React Suite passed 1/1. `git diff --check` passes. Other legacy Ribbon custom controls, menu bodies, and feature dialog bodies remain outstanding.

The production suite also verifies the message-box title/content/icon, sanitized access-key labels, primary button semantics, explicit selection indexes, and the legacy Escape/cancel result across Wordx, Excelx, and PPTX.

- [x] Shared React UI pieces can be used by all three editor packages without copying shared code.
- [x] Classic, Paper, and Paper 2016 looks retain their current chrome colors, typography, and layout behavior.
- [x] Existing command, menu, dialog, keyboard, and focus behavior has a React-facing integration path.
- [x] Existing language selection and translation lookup continue to work; new strings use the English catalog mechanism.
- [x] The shared interface renders in the host fixture and in the standalone Vite shell.
- [x] A shared React command bar mounts consistently for Wordx, Excelx, and PPTX, with commands routed into the hosted engine.
- [x] Look and language selectors are part of the React editor chrome and synchronize to the hosted engine.
- [ ] Legacy app menus, dialogs, and ribbon controls are migrated to shared React components while preserving commands, keyboard, and focus behavior.

2026-11-24 follow-up: Production React Suite (including all three editor menu/dialog/Ribbon workflows), standalone package launches, external host integration, shared package tests (42/42), classic smoke (4/4), all workspace/standalone/Suite production builds, and `git diff --check` pass. These validate the migrated shared layer; remaining imperative feature-dialog bodies and app-specific content keep this ticket open.
