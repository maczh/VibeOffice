# 14: Host integration example and package delivery

**What to build:** A React application outside the suite can install and embed each complete editor package using its public types, styles, and storage-provider contract.

**Blocked by:** 00: Direct editor-engine mount foundation; 13: Suite Start Center, Recent Files removal, and PWA behavior; 16: Shared UI module migration to React; 17: Wordx full module migration and naming; 18: Excelx full module migration and naming; 19: PPTX full module migration and naming; 20: Web-wide legacy brand name cleanup.

**Status:** in-progress

The React packages now mount their current editor engines directly in the consumer's document without iframes. Their engines and much of the shared/app UI still come from the legacy public script bundles, so package delivery remains dependent on the full module migrations and final brand audit.

- [x] Each editor package builds independently with its public React component, styles, and TypeScript declarations.
- [x] React and React DOM are declared as peer dependencies.
- [x] A consumer example mounts each editor without rendering the suite Start Center.
- [x] The example demonstrates local file operations and a host-provided remote storage adapter.
- [x] Public API documentation describes configuration, callbacks, styles, provider behavior, and deployment CSP/CORS requirements.
- [x] Package artifacts do not require VibeOffice backend services or runtime development-server dependencies.
