# 20: Web-wide legacy brand name cleanup

**What to build:** Remove product-brand uses of Quire, Ledger, and Lectern from the web source, package metadata, routes, storage namespaces, clipboard protocols, user-facing strings, documentation, tests, and generated assets; use Wordx, Excelx, and PPTX consistently.

**Blocked by:** 17: Wordx full module migration and naming; 18: Excelx full module migration and naming; 19: PPTX full module migration and naming.

**Status:** claimed

**Progress:** React package and suite runtime routes now use `/wordx/`, `/excelx/`, and `/pptx/`; the Vite adapter maps those routes to the still-legacy engine source directories internally. The React build transforms its PWA shortcuts and screenshot paths to the new route/product names without changing the still-live classic public manifest. Standalone outputs are cleaned before emitting only their new routes. Legacy source folder names and many classic app identifiers/strings remain, so the repository-wide brand audit is not complete.

## Acceptance

- [ ] A case-insensitive repository audit finds no old product names in web-owned source, UI strings, paths, docs, tests, or generated web outputs.
- [ ] Old product-name references remain only where needed to identify historical source provenance or documented migration history, and those exceptions are narrowly listed.
- [ ] Microsoft product names, standards terminology, OOXML namespaces, and format-extension names are unchanged.
- [ ] Desktop/Tauri files remain untouched.
- [ ] Routes, persistence keys, template identifiers, MIME protocols, standalone app launch, and suite launch agree on the new names.
- [ ] Full tests, package builds, and `git diff --check` pass.

## Comments

- 2026-11-12: Started the React-package-facing rename: standalone document titles and exported React editor chrome now say Wordx, Excelx, and PPTX; package rendering expectations were updated. The package build passed; the first package test exposed its old Quire label assertion, which was updated, and all 21 package tests then passed. Legacy engine pages, runtime routes, suite labels, storage identifiers, docs, and clipboard protocols still need migration; routes were intentionally left unchanged until their runtime moves to avoid breaking standalone editors.
- 2026-11-12: The suite's app IDs and Start Center labels now use `wordx`, `excelx`, and `pptx`; legacy folder routes remain mapped explicitly until their runtime migration. Typecheck, all workspace builds, React-suite + smoke browser tests (5/5), and `git diff --check` pass. The embedded legacy pages still display old product names, so the brand audit remains open.
- 2026-11-12: Rebuilt all standalone apps and the suite with `npm run build:apps` (the root `npm run build` alone does not build the suite app), updated browser assertions for the new frame titles, and verified the production suite test passes. `npm run build:apps` still warns that the package HTML loads legacy `/common/look.js`, `/common/lang.js`, and `/common/i18n.js` as classic scripts and that shared CSS is resolved only at runtime; this is evidence that shared-module migration remains outstanding.
- 2026-11-12: Rebranded the app document-title suffixes to Wordx, Excelx, and PPTX and the Wordx/Excelx HTML generator metadata for browser runs. The production suite regression verifies title suffixes for all three directly mounted engines; it failed first on `Document2 - Quire`, then passed after the names changed. The existing desktop-host one-document title suffix remains unchanged because Rust/Tauri is out of scope. Other old identifiers, folders, UI strings, and generated assets remain to be migrated.
- 2026-11-12: Rebranded the default Wordx sample's document title, creator, heading, header, and default profile label. Replaced user-visible Excelx references in formula repair, file repair, search, print, picture/dependency errors, and About; updated the English catalog with `node tools/i18n.mjs gather` (no non-English catalog files were accessed). The focused Wordx sample test and i18n tests pass. Most legacy app source strings and identifiers are still present.
- 2026-11-12: Replaced capitalized Quire product mentions throughout `public/quire/js/` with Wordx, refreshed the English catalog, and passed the full Wordx UI suite (52/52), i18n checks (7/7), smoke, and workspace builds. Lowercase route/host IDs and the legacy runtime mapping remain intentionally for compatibility during migration; other old app names still need cleanup.
- 2026-11-12: Replaced capitalized Ledger and Lectern product mentions throughout their classic JS module trees with Excelx and PPTX, respectively, and refreshed the English catalog. Preserved lowercase runtime route/host IDs and desktop title/edition compatibility. Full Excelx UI + smoke passed 47/47; PPTX UI + smoke passed 43/43; i18n passed 7/7; workspace builds passed. Lowercase identifiers, old source directories, legacy public assets, localized catalogs, and React UI migration remain.
