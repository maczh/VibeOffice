# React packages

The React migration is split into four workspaces:

- `rshared/` contains the framework-neutral React package boundary: `OfficeApi`, online source/target types, browser file transfer helpers, the error boundary and shared Office-style chrome.
- `rword/` exports `RWord`, the Quire React entry point and its Vite library configuration.
- `rexcel/` exports `RExcel`, the Ledger React entry point and its Vite library configuration.
- `rpptx/` exports `RPptx`, the Lectern React entry point and its Vite library configuration.

Each package also exposes `./engine`, which contains the existing framework-neutral OOXML model/read/write/render primitives. The React build aliases their old Vue-facing core import to a Vue-free `rshared/src/office-core.ts` boundary, so importing an engine does not pull Vue UI components into the React bundle.

Each editor accepts `src?: File | Blob | string | OnlineSource`. `string` is an HTTP GET source. `OnlineSource` supports WebDAV, S3-compatible endpoints, presigned GET/PUT URLs and authenticated HTTP GET. `saveTo()` additionally accepts an HTTP multipart POST target. Local saves are delivered through `onSave` and `saveAs()` downloads a browser file.

```tsx
import { RWord, type OfficeApi } from '@voffice/rword'
import '@voffice/rword/styles'

const ref = useRef<OfficeApi>(null)
return <RWord ref={ref} fileName="Letter.docx" onSave={(blob, name) => { /* persist */ }} />
```

Run a package with `npm run dev -w rword`, `npm run dev -w rexcel` or `npm run dev -w rpptx`. Build libraries with `npm run build -w rword` (and the matching package). Build all React workspaces with `npm run build`.

The original `public/` implementation and the existing Vue packages remain untouched. The React packages are intentionally a separate migration track; the document engines and full Office-2003 UI still need to be ported into the three React packages before claiming feature parity with the legacy applications.
