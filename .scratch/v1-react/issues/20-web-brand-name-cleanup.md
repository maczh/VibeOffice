# 20: Web-wide legacy brand name cleanup

**What to build:** Remove product-brand uses of Quire, Ledger, and Lectern from the web source, package metadata, routes, storage namespaces, clipboard protocols, user-facing strings, documentation, tests, and generated assets; use Wordx, Excelx, and PPTX consistently.

**Blocked by:** 17: Wordx full module migration and naming; 18: Excelx full module migration and naming; 19: PPTX full module migration and naming.

**Status:** ready-for-agent

## Acceptance

- [ ] A case-insensitive repository audit finds no old product names in web-owned source, UI strings, paths, docs, tests, or generated web outputs.
- [ ] Old product-name references remain only where needed to identify historical source provenance or documented migration history, and those exceptions are narrowly listed.
- [ ] Microsoft product names, standards terminology, OOXML namespaces, and format-extension names are unchanged.
- [ ] Desktop/Tauri files remain untouched.
- [ ] Routes, persistence keys, template identifiers, MIME protocols, standalone app launch, and suite launch agree on the new names.
- [ ] Full tests, package builds, and `git diff --check` pass.

## Comments

- 2026-11-12: Started the React-package-facing rename: standalone document titles and exported React editor chrome now say Wordx, Excelx, and PPTX; package rendering expectations were updated. The package build passed; the first package test exposed its old Quire label assertion, which was updated, and all 21 package tests then passed. Legacy engine pages, runtime routes, suite labels, storage identifiers, docs, and clipboard protocols still need migration; routes were intentionally left unchanged until their runtime moves to avoid breaking standalone editors.
