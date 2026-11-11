Status: ready-for-agent

# Vite and React editor packages

## Problem Statement

VibeOffice currently runs as three browser editors and a Start Center assembled from ordered classic JavaScript files. The editors depend on shared globals, and there is no Vite workspace or distributable React component interface. This makes the editors difficult to embed independently in other applications and leaves no unified frontend build for running the editors separately or as a suite.

The web frontend also needs clear local and remote file workflows. The new version must retain the current browser editing and Office file compatibility while adding local file write-back and remote URL, WebDAV, custom server, and S3-provider workflows. Recent Files is not wanted in the new version.

## Solution

Build the browser product as a Vite workspace with three TypeScript/React editor packages in `wordx`, `excelx`, and `pptx`, plus a shared package for common browser and document services. Each editor exports a complete React editor that can run from its own Vite entry or be embedded in another React application. A suite entry composes the same editors behind the existing VibeOffice Start Center and look-and-feel.

Keep the mature document and shared JavaScript engines. Wrap their app lifecycle and shared services behind typed interfaces used by the React editors. Do not rewrite OOXML readers/writers, document models, or spreadsheet calculation as part of this migration.

### Expanded migration and product naming scope

The React migration now includes every user-interface and application module currently implemented as classic JavaScript under `public/quire/js/`, `public/ledger/js/`, and `public/lectern/js/`, plus the shared UI components in `public/common/`. These must be migrated into the `wordx/`, `excelx/`, `pptx/`, and shared React workspace packages. User-facing UI, dialogs, menus, panes, ribbons, editors, and interaction handlers become React components/hooks. OOXML codecs, document models, spreadsheet calculation, geometry, cryptography, and other non-UI algorithms remain ordinary typed service modules consumed by those components; wrapping pure algorithms in artificial React components is not required. No app package may load the legacy `public/<old-name>/js` runtime after its migration ticket is complete.

The former product names Quire, Ledger, and Lectern are replaced by Wordx, Excelx, and PPTX throughout the web product: source identifiers, app IDs, routes, package metadata, storage keys, clipboard MIME types, UI strings, docs, tests, and generated web assets. Preserve Office and Open XML standard names, namespaces, file-format identifiers, and references to Microsoft Word/Excel/PowerPoint where those are technical interoperability terminology rather than this product's brand. Desktop/Tauri remains out of scope.

Preserve existing editing features and the current appearance, including the three looks and supported languages. Add local and remote file choices inside the existing open/save flows. Remove Recent Files lists and entry points; make Create File the Start Center's default view. Keep local unsaved-document recovery as a separate draft service. Do not automatically delete existing browser Recent Files data.

The browser application supports local file selection and saving. When the browser permits a writable file handle, save to the selected file; otherwise save through the browser's download flow. URL opening is read-only unless a configured provider supports saving. WebDAV supports browsing and document operations. Custom server POST and S3 operations are supplied through a host-provided storage provider; S3 signing and service-specific listing remain the host's responsibility. No server is added to VibeOffice.

Rust/Tauri desktop code is outside this web migration and remains in the repository. The new web application does not depend on the desktop host interface.

## User Stories

1. As a VibeOffice user, I want Wordx, Excelx, and PPTX to retain the existing editors' tools and behavior, so that the migration does not disrupt my work.
2. As a VibeOffice user, I want the editors to retain their current visual appearance, so that the React version feels like the same product.
3. As a VibeOffice user, I want to choose the existing Classic, Paper, and Paper 2016 looks, so that the suite remains familiar in every look.
4. As a VibeOffice user, I want the existing supported UI languages and language selection to continue working, so that the migration does not reduce accessibility for current users.
5. As a VibeOffice user, I want the Start Center to open on Create File, so that I can start a blank document or choose a template immediately.
6. As a VibeOffice user, I want to open a document from my device, so that I can work with files already stored locally.
7. As a VibeOffice user, I want to save a locally opened document back to its chosen file when the browser supports it, so that I can continue editing the same file.
8. As a VibeOffice user on a browser without writable file handles, I want saving to produce a normal file download, so that local file saving remains available.
9. As a VibeOffice user, I want to create a blank local document or one from a template, so that I can author new Office files.
10. As a VibeOffice user, I want to open an Office file from an HTTPS URL, so that I can edit a file already served by an accessible web location.
11. As a VibeOffice user, I want to browse a WebDAV location and open a document, so that I can work with files on a WebDAV server.
12. As a VibeOffice user, I want to save or create a document at a WebDAV location, so that the result is stored with my remote files.
13. As a user embedding an editor, I want to provide a storage provider for my server's list, authentication, and POST contract, so that the editor works with my application's existing file service.
14. As a user embedding an editor, I want my host to provide short-lived S3 upload and download authorization, so that documents can be stored in S3 without AWS credentials in the editor.
15. As a host application developer, I want the provider interface to report file revisions when available, so that the editor can warn before overwriting a remotely changed document.
16. As a host application developer, I want storage-provider origins to be limited by deployment configuration, so that browser network requests follow the embedding application's security policy.
17. As a host application developer, I want credentials to be supplied by the host or held only for the active session, so that the editor does not persist secrets.
18. As a host application developer, I want each editor as a complete React component with typed configuration and callbacks, so that I can embed it without adopting VibeOffice's Start Center.
19. As a host application developer, I want the editor packages to declare React as a peer dependency and expose their built styles and type declarations, so that they integrate into an existing React application.
20. As a VibeOffice user, I want each editor to run independently and the suite to run from one Start Center, so that the same packages support standalone and combined use.
21. As a VibeOffice user, I want the suite to remain installable and usable offline for local documents, so that the Vite build retains the current PWA behavior.
22. As a VibeOffice user, I want unsaved changes to be recoverable after a browser crash, so that removing Recent Files does not remove local draft recovery.
23. As a VibeOffice user, I want the Start Center and app menus to omit Recent Files, so that the new product no longer offers that feature.
24. As a VibeOffice user, I want the new version not to silently erase old Recent Files data already stored in my browser, so that migration does not unexpectedly delete my data.
25. As a VibeOffice user, I want saved documents to remain interoperable with Microsoft Office and LibreOffice, so that I can exchange files with those applications.

## Implementation Decisions

- Use a Vite workspace and strict TypeScript for new React components, package APIs, and browser-facing adapters. Existing document engines remain JavaScript and are accessed through typed facades.
- Create editor packages in the requested `wordx`, `excelx`, and `pptx` directories, with a separate shared package for code used by multiple editors.
- Each editor package exports a full React editor, a standalone Vite entry, its stylesheet, and TypeScript declarations. React and React DOM are peer dependencies. The suite entry composes these packages; public registry publication is not part of this work.
- The React editor packages mount the existing JavaScript document engines directly into host-owned DOM containers. Do not use iframes, nested browsing contexts, or cross-window messaging to host an editor. Refactor engine bootstraps and selectors behind per-editor mount/dispose APIs as needed; retain the document models, file-format engines, and editing behavior.
- Preserve each editor's commands, file-format behavior, templates, looks, language selection, keyboard and pointer interactions, and browser launch behavior while adopting the Wordx, Excelx, and PPTX identities.
- Keep the current Start Center's Create File gallery and make it the default. Remove Recent Files cards, links, settings, clear controls, app-menu entries, and recent-file persistence from the new web product. Leave pre-existing browser data untouched. Keep unsaved drafts and Document Recovery in a separate local-only store.
- Define a typed storage-provider interface for listing, reading, creating, saving, reporting revisions, and handling provider-specific authentication. The standalone web product includes local file access, read-only URL loading, and WebDAV browsing/open/save. Embedded hosts can supply custom server POST and S3 providers, including list and signing operations.
- Do not persist remote credentials. A host may inject its own authentication callback; standalone WebDAV credentials are held only in memory for the active session.
- Configure allowed remote origins for each deployment. Remote services must allow the required CORS requests. Do not allow package code or document data to load from third-party script or font servers.
- Use conditional saves when a provider supplies a revision such as an ETag. On a detected conflict, show the user a choice before replacing remote content. If a provider supplies no revision, report that the save cannot detect concurrent changes.
- Use writable browser file handles when available and obtained through a user action. Fall back to file-input open and download-based save on browsers without those APIs.
- Keep the Rust/Tauri implementation and its build files untouched. The new web packages do not implement or depend on native desktop windows, dialogs, or host commands.
- Complete migration by moving all three legacy app JS module sets into their new package source trees and replacing UI modules with React components; preserve algorithmic document/file services as non-UI modules. Migrate shared `public/common/` UI components into the shared package and keep shared services single-sourced.
- Use Wordx, Excelx, and PPTX consistently in all web-owned identifiers and strings. Update routes, storage namespace keys, clipboard MIME types, tests, documentation, and generated web outputs together; do not rename standards-defined OOXML names or desktop/Tauri code.
- Preserve the current Content Security Policy principles: no inline scripts, dynamic code evaluation, CDN code, or runtime build dependency. Configure only document-service connections needed by a deployment.

## Testing Decisions

- Test behavior at the browser-visible production-build seam in Chromium, using real keyboard and pointer input rather than React implementation details.
- Use current-version screenshots as the visual baseline for the Start Center and all three editors. Compare the existing looks and representative narrow and normal window sizes; only the explicitly removed Recent Files UI and added file-provider entry points may differ.
- Exercise the suite Start Center, each independent editor entry, and a host page that mounts each exported React component. Cover create blank/template, edit, open, save, and host callbacks.
- Inject fake storage providers at the package boundary. Verify URL reads, WebDAV listing/read/write requests, custom POST calls, S3 signing callbacks, origin restrictions, authentication lifetime, and conflict handling without depending on production credentials or third-party services.
- Verify local file open and save paths in Chromium, plus the download fallback where writable handles are unavailable.
- Keep existing Node engine tests, OOXML round trips, corpus compatibility checks, and Office/LibreOffice validation as protection for the reused document engines.
- Keep PWA/service-worker checks for offline local editing, install metadata, and production assets. Check browser console output and network requests against CSP and the configured remote-origin policy.
- Treat the current UI smoke suite, Start Center launch checks, app-specific UI suites, and file-format round-trip harnesses as prior art. Assertions target visible behavior and saved file content, not React component internals.

## Out of Scope

- Rewriting Quire, Ledger, or Lectern document models, OOXML readers/writers, spreadsheet formulas, or other document engines in React or TypeScript.
- Developing or hosting a VibeOffice server, authentication service, WebDAV server, custom application API, or S3 signing service.
- Publishing packages to a public or private registry.
- Migrating Rust/Tauri, changing desktop installers, or deleting desktop source, tests, documentation, or release tooling.
- Retaining Recent Files lists, saved-file copies, or recent-document navigation in the new web product.
- Changing Office file formats, adding new document features, translating new UI text, or redesigning existing editor screens.

## Further Notes

- Current browser storage couples saved Recent Files copies, recent lists, autosave drafts, and Document Recovery in one service. The implementation must separate local draft recovery from Recent Files before removing recent persistence.
- Existing recent data already in a browser is left untouched by the new version; it is no longer read by the new product.
- User-directed document requests may connect to configured file-service origins. CORS and Content Security Policy settings are deployment requirements, not functionality that VibeOffice can override in the browser.
- The project remains a static browser application. Vite is a development and packaging tool, not a server runtime dependency.
