# React editor packages

The browser workspace is built from `commonx`, `wordx`, `excelx`, and `pptx`. The editor packages expose `WordxEditor`, `ExcelxEditor`, and `PptxEditor`; React and React DOM are peer dependencies. `npm run build` emits JavaScript declarations and `style.css` artifacts under each package's `dist/`. `npm run build:apps` emits independently hosted React shells plus the existing static document engine under each package's `standalone/` directory. Serve the contents of that directory at its origin root; no VibeOffice API server is required.

The target architecture mounts each existing editor engine directly into the React package's host-owned DOM; editor packages must not use iframes or another nested browsing context. Each legacy app now exposes `L.app.mount(host)`, which scopes its existing initializer and returns an idempotent disposer; standalone pages use the same API with `document.body`. A host can mark an existing shell with `[data-vibeoffice-engine-root]` so classic-page bootstrap mounts there instead of on `body`. Keyboard handlers (and Lectern clipboard handlers) are registered during mount and removed with its disposer. This is an intermediate lifecycle seam, not yet a direct React integration: `host` must already contain the app's complete legacy shell, app state remains global, and some listeners are still installed while the classic scripts load. The current React adapters still launch same-origin editor pages in iframes, so the embedding contract is not yet met.

Example package use:

```tsx
import { WordxEditor } from '@vibeoffice/wordx';
import '@vibeoffice/wordx/style.css';

export function DocumentEditor() {
  return <WordxEditor />;
}
```

The final host page must also load the shared look variables (`/common/looks.css`) before the editor stylesheet, and serve the selected engine assets plus `/common/`, `/fonts/`, and `/icons/` from the same origin. Standalone Vite entries include these shared look styles and runtime assets.

`ExcelxEditor` and `PptxEditor` have equivalent APIs. The host may provide `onOpenFile`, `onSaveFile`, `onDraftSaved`, `onEngineReady`, and a `DraftStore`. The default local adapter uses the File System Access API when available and otherwise downloads saves. Crash drafts use a dedicated `vibeoffice-react-drafts` IndexedDB database; the legacy Recent Files database is not read or deleted in React-host mode.

The `@vibeoffice/core` exports include `RemoteStorageProvider`, `createWebDavProvider`, and `createUrlReadProvider`. A host can implement `RemoteStorageProvider` for custom POST endpoints or S3; signing, listing, and authentication remain the host's responsibility. Pass host-owned file pickers as `onOpenFile`/`onSaveFile` callbacks. The WebDAV credential callback is invoked for each operation and credentials are not stored by the provider. Revisions are represented as ETags; conditional writes use `If-Match`, and HTTP 409/412 responses report a revision conflict. Filesystems or APIs without revisions cannot protect against concurrent changes.

For an embedded custom provider, keep the user's listing choice in the host and adapt it to the editor callbacks:

```tsx
import { createRemoteFileCallbacks } from '@vibeoffice/core';
import { ExcelxEditor } from '@vibeoffice/excelx';

const remote = createRemoteFileCallbacks(myProvider, chooseRemoteFile, chooseRemoteDestination);
<ExcelxEditor
  engine={{ url: '/ledger/' }}
  onOpenFile={remote.onOpenFile}
  onSaveFile={remote.onSaveFile}
/>
```

`myProvider` implements the typed list/read/create/save contract. A host that uses S3 supplies its own short-lived signed requests and never passes AWS credentials to the editor. The suite's Start Center also offers Open URL and WebDAV; build it with `VITE_REMOTE_ORIGINS=https://dav.example,https://files.example` so both the provider origin allowlist and generated CSP are explicit.

For remote access, the embedding application's Content Security Policy must list only the intended service origins in `connect-src`, and the remote service must allow the app origin in its CORS policy. The URL reader is read-only. WebDAV uses browser `PROPFIND`, `GET`, and `PUT`; the server must allow those methods and the `Authorization`, `If-Match`, and `If-None-Match` request headers, and expose `ETag` to browser JavaScript. Do not place long-lived secrets in package configuration or browser storage.

The shared React editor chrome provides New/Open/Save commands and look/language selectors for every package. New creates a blank document directly; returning from the suite to Create File asks the engine to resolve unsaved changes and keeps the editor mounted if the user cancels. The mature vanilla-JavaScript document engines and OOXML handling will be mounted directly into each package host; the existing app ribbon, menus, and dialogs are not yet React components. The standalone packages include the existing app resources and service worker. The Vite configuration is a build/development tool, not a production runtime dependency.
