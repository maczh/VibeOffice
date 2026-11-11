# React editor packages

The browser workspace is built from `commonx`, `wordx`, `excelx`, and `pptx`. The editor packages expose `WordxEditor`, `ExcelxEditor`, and `PptxEditor`; React and React DOM are peer dependencies. `npm run build` emits JavaScript declarations and `style.css` artifacts under each package's `dist/`. `npm run build:apps` emits independently hosted React shells plus the existing static document engine under each package's `standalone/` directory. Serve the contents of that directory at its origin root; no VibeOffice API server is required.

An embedding host supplies an engine URL served from the same origin as its page, for example:

```tsx
import { WordxEditor } from '@vibeoffice/wordx';
import '@vibeoffice/wordx/style.css';

export function DocumentEditor() {
  return <WordxEditor engine={{ url: '/quire/' }} />;
}
```

`ExcelxEditor` and `PptxEditor` take equivalent `engine` props. The host may instead provide `mountEngine`, `onOpenFile`, `onSaveFile`, `onDraftSaved`, `onEngineReady`, and a `DraftStore`. The default local adapter uses the File System Access API when available and otherwise downloads saves. Crash drafts use a dedicated `vibeoffice-react-drafts` IndexedDB database; the legacy Recent Files database is not read or deleted in React-host mode.

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

The iframe currently preserves the mature vanilla-JavaScript document engines and their OOXML handling. React provides the package and host boundary; it does not replace each app's internal editing UI or engine. The standalone packages include the existing app resources and service worker. The Vite configuration is a build/development tool, not a production runtime dependency.
