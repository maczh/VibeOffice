# React packages

The React migration is split into four workspaces:

- `rshared/` contains the framework-neutral React package boundary: `OfficeApi`, online source/target types, browser file transfer helpers, the error boundary and shared Office-style chrome.
- `rword/` exports `RWord`, the Quire React entry point and its Vite library configuration.
- `rexcel/` exports `RExcel`, the Ledger React entry point and its Vite library configuration.
- `rpptx/` exports `RPptx`, the Lectern React entry point and its Vite library configuration.

Each package also exposes `./engine`, which contains its self-contained OOXML model/read/write/render primitives. The React engines and shared online/core services no longer depend on Vue packages.

Each editor accepts `src?: File | Blob | string | OnlineSource`. `string` is an HTTP GET source. `OnlineSource` supports WebDAV, S3-compatible endpoints, presigned GET/PUT URLs and authenticated HTTP GET. `saveTo()` additionally accepts an HTTP multipart POST target. Local saves are delivered through `onSave` and `saveAs()` downloads a browser file.

The React chrome includes Office-style File/Edit/View/Insert/Format/Tools/Help menus, toolbars, Ribbon tabs, status bars, context menus, online-file dialogs, About dialogs, keyboard shortcuts and bounded undo/redo history. Quire includes Find/Replace; Ledger includes formula recalculation, sheet/row/column insertion; Lectern includes multi-slide text modeling, New Slide and a browser Slide Show preview.

```tsx
import { RWord, type OfficeApi } from '@voffice/rword'
import '@voffice/rword/styles'

const ref = useRef<OfficeApi>(null)
return <RWord ref={ref} fileName="Letter.docx" onSave={(blob, name) => { /* persist */ }} />
```

Run a package with `npm run dev -w rword`, `npm run dev -w rexcel` or `npm run dev -w rpptx`. Run the integrated host with `npm run dev:react`. Build libraries with `npm run build -w rword` (and the matching package). `npm run build` builds all React workspaces, including the integrated host; `npm run build:react` is an alias.

Validation currently includes strict TypeScript checking, all four React workspace builds, package export smoke checks, OOXML core regression tests, and online connector coverage for WebDAV, S3 SigV4, HTTP multipart POST and presigned URLs. Real-browser visual verification requires a browser instance in the execution environment.

The three React packages now own their DOCX/XLSX/PPTX readers and writers for real file open/save. The component API accepts complete WebDAV/S3/presigned/HTTP source objects, including credentials and S3 region; `OnlineSourceDialog` covers interactive WebDAV/S3/presigned/HTTP POST configuration. The original `public/` implementation remains the reference implementation; this branch contains only React component workspaces.
