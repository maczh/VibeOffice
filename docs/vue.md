# VibeOffice — Vue 3 component rewrite

This branch (`v1-vue`) carries a second front end for the suite: the three apps rebuilt as
**Vite + Vue 3 + TypeScript** components, in `vword/`, `vexcel/` and `vpptx/`. The original
classic-script apps under `public/` are untouched and keep working as before; nothing here is
loaded by them. The desktop (Tauri) code is not part of this rewrite — these components are
pure browser front ends.

The rewrite reuses the suite's *design* (Office-2003 chrome, the three looks, the OOXML
preservation rules) but the code is new: TypeScript engines, Vue components, Composition API
with `<script setup>`. The shared chrome CSS (`looks.css`, `ui.css`) is carried over unchanged
so the components look like the original apps.

## Directory layout

| Directory | Responsibility |
|---|---|
| `voffice-shared/` | Everything the three components share: the ZIP/XML/OPC engines, the online-file connectors (WebDAV, S3, HTTP), the Office-2003 UI kit (menus, toolbars, dialogs, ribbon, status bar, task pane), the looks, icons, i18n `t()`, and the common component contract (`api.ts`). Published as `@voffice/shared`. |
| `vword/` | **VWord** — the word processor (Quire). Document model, DOCX read/write, paginated layout, editing. Published as `@voffice/vword`. |
| `vexcel/` | **VExcel** — the spreadsheet (Ledger). Workbook model, canvas grid, formula engine, XLSX/CSV read/write. Published as `@voffice/vexcel`. |
| `vpptx/` | **VPptx** — the presentation editor (Lectern). Slide model, PPTX read/write, shape rendering/editing, slideshow. Published as `@voffice/vpptx`. |
| `playground/` | A host app that runs all three components together on one page (tabbed), demonstrating integration, the shared look switch, and host-side open/save handling. Not published. |

Each app package has the same shape:

```
vword/
  src/
    VWord.vue          # the root component (chrome + editing surface)
    index.ts           # public entry: exports the component + a Vue plugin
    engine/            # TypeScript engines (model, read, write, layout, …)
    styles/app.css     # the app's own chrome CSS (from the original index.html)
  dev/                 # standalone dev harness (main.ts + App.vue)
  index.html           # standalone dev entry
  vite.config.ts       # library build + dev server
  tsconfig.json        # typecheck (includes shared sources)
  tsconfig.build.json  # declaration emit for the published types
```

## The component contract

Every office component implements the same interface, defined in
[`voffice-shared/src/api.ts`](../voffice-shared/src/api.ts). A host drives a component through
**props in, events out, methods on a template ref**.

### Props (`OfficeProps`)

| Prop | Type | Meaning |
|---|---|---|
| `src` | `string \| File \| Blob \| OnlineSource` | A file to open on mount: a local `File`/`Blob`, an `http(s)` URL, or a typed online source. Omit for a blank document. |
| `fileName` | `string` | Name to show/save as when `src` carries none. |
| `look` | `'classic' \| 'paper' \| 'paper2016'` | The chrome look. Defaults to the stored suite setting. |
| `readonly` | `boolean` | Open for viewing only. |

### Events (`OfficeEmits`)

| Event | Payload | When |
|---|---|---|
| `ready` | — | The component mounted and its first document is loaded. |
| `change` | `dirty: boolean` | The document's dirty state changed. |
| `save` | `blob: Blob, name: string` | The user saved and the component produced the file; the host keeps it (e.g. downloads it). |
| `error` | `err: Error` | An open/save/edit failure the host should surface. |
| `open-request` | — | The user chose *File ▸ Open*; the host should show its own picker and call `open()`. |
| `save-request` | — | The user chose *Save As* with no writable source; the host should call `saveAs()`/`saveTo()`. |

### Methods (`OfficeApi`, via `defineExpose` / template ref)

| Method | Returns | Meaning |
|---|---|---|
| `open(input, name?)` | `Promise<void>` | Open a `File`/`Blob`, a URL string, or an `OnlineSource`. |
| `newDocument()` | `void` | Start a blank document. |
| `getBlob()` | `Promise<Blob>` | The current document as a file, in the format it was opened in. |
| `save()` | `Promise<void>` | Write back to the source it came from when writable; otherwise emit `save`/`save-request`. |
| `saveAs(name?)` | `Promise<Blob>` | Produce the file (and download it in the standalone harness). |
| `saveTo(target, name?)` | `Promise<void>` | Save to an explicit online `SaveTarget`. |
| `name` / `dirty` | `string` / `boolean` | Read-only state. |

## Online files

Opening and saving remote files is a first-class feature, handled by shared connectors in
[`voffice-shared/src/online/`](../voffice-shared/src/online/). All requests run in the browser;
the remote server must allow the page's origin via CORS. Credentials never leave the page.

| Source / target `kind` | Open | Save | Notes |
|---|---|---|---|
| `webdav` | `GET` | `PUT` | Basic (`username`/`password`) or `token` (Bearer). `listOnline` does `PROPFIND`. |
| `s3` | `GET` | `PUT` | AWS Signature V4 signed in-browser with Web Crypto (`accessKeyId`/`secretAccessKey`, optional `sessionToken`). Works with AWS S3, MinIO, etc.; `forcePathStyle` for path-style addressing. `listOnline` does ListObjectsV2. |
| `presigned` | `GET url` | `PUT putUrl` | No credentials in the front end; the host's backend mints the URLs. Read-only without `putUrl`. |
| `http` | `GET` | — | Plain download; honours `Content-Disposition` filename. |
| `post` | — | `POST` | Multipart upload to an endpoint (`field`, extra `fields`). |

Helpers: `fetchOnline(source)`, `saveOnline(target, blob, name)`, `listOnline(source, pathOrPrefix)`.
The shared `VoOnlineDialog` component collects these fields from the user and emits a typed
source/target; each app's *File ▸ Open from Online… / Save to Online…* uses it.

## Running

Install once at the repo root (npm workspaces):

```bash
npm install
```

### Each component on its own

```bash
npm run dev:word     # vword standalone  (vite dev server)
npm run dev:excel    # vexcel standalone
npm run dev:pptx     # vpptx standalone
```

Each serves its `index.html` → `dev/main.ts`, which mounts the component full-viewport and wires
`save` to a download and `open-request` to a file picker.

### All three together

```bash
npm run dev          # playground: tabbed host running VWord + VExcel + VPptx
```

### As a component package in another app

Build the libraries (each emits an ES module, a single stylesheet, and `.d.ts` types):

```bash
npm run build        # builds shared, vword, vexcel, vpptx, playground
```

Then in a host Vue 3 app:

```ts
import { createApp } from 'vue'
import { VWord } from '@voffice/vword'
import '@voffice/vword/styles'        // the chrome + app CSS (looks included)

createApp(App).component('VWord', VWord).mount('#app')
```

```vue
<template>
  <VWord ref="word" :src="file" look="paper2016"
         @save="onSave" @error="onError"
         @open-request="pickAndOpen" @save-request="saveAs" />
</template>

<script setup lang="ts">
import { ref } from 'vue'
import type { OfficeApi } from '@voffice/shared'
const word = ref<OfficeApi | null>(null)
function onSave(blob: Blob, name: string) { /* keep the file */ }
function onError(err: Error) { /* surface it */ }
async function pickAndOpen() { /* show a picker, then */ await word.value?.open(pickedFile) }
async function saveAs() { await word.value?.saveAs() }
</script>
```

Or install the whole plugin: `import VWordPlugin from '@voffice/vword'; app.use(VWordPlugin)`.
The same pattern applies to `VExcel` (`@voffice/vexcel`) and `VPptx` (`@voffice/vpptx`). Only
`vue` is a peer dependency; the shared core is bundled into each library.

Give the component a sized container (it fills `height: 100%`). Several components can share a
page — each keeps its own state and scopes its keyboard handling to its own root.

## Tests

```bash
node --test voffice-shared/test/      # core engines + online connectors (against the built lib)
npm run typecheck                     # vue-tsc across the workspace
```

The online tests run the connectors against a local mock server; the S3 Signature V4 output is
cross-checked against an independent `node:crypto` implementation.

## Relationship to the original apps

The original `public/quire`, `public/ledger`, `public/lectern` remain the reference for behaviour
and format fidelity, and their compatibility corpora and tests still apply to them. This rewrite
reimplements the engines in TypeScript. It is a **working functional subset**, not a line-for-line
port: the core open/edit/save loop, the Office-2003 chrome and the three looks are reproduced, while
the originals' long tail of advanced features is not (yet) present. Coverage and gaps per app:

### VWord (`vword/`)

Implemented: DOCX read/write (paragraphs, runs with bold/italic/underline/strike/font/size/color/
highlight, hyperlinks, inline pictures, tables with merges/shading, lists, page size/margins, page
breaks, named styles); paginated Print Layout + Web Layout; contenteditable editing with
execCommand formatting; find/replace; insert table/picture/hyperlink/symbol/date; undo/redo; word
count; zoom; print; the full menu/toolbar/ribbon chrome. Untouched source parts (numbering, media,
headers, …) are written back byte-for-byte.

Gaps vs Quire: no fields/mail-merge/RTF/Markdown/HTML I/O, no comments/revisions/track-changes, no
footnotes/endnotes/TOC, no headers/footers editing, no spelling/grammar, no rulers, no SmartArt/
WordArt/charts, no mid-paragraph page splitting (paragraphs/tables stay whole on a page), no
multi-window. Complex unmodelable runs are kept as raw XML and shown as `[object]`.

### VExcel (`vexcel/`)

Implemented: XLSX read/write (cells, formulas, shared strings, styles: font/fill/border/alignment/
number-format, merges, column widths/row heights, freeze panes), CSV/TSV import/export; canvas grid
with headers, selection, styled cells; cell + formula-bar editing; keyboard navigation; copy/cut/
paste (internal with relative-ref translation, and TSV); a formula engine (Pratt parser) with ~90
functions (SUM/AVERAGE/IF/VLOOKUP/INDEX/MATCH/SUMIF/COUNTIF/text/date/math/logic), dependency-ordered
recalc and circular-ref detection; number formats; AutoSum; sort; insert/delete rows/cols/sheets;
undo/redo; zoom; gridlines toggle.

Gaps vs Ledger: no charts/pivot tables/slicers, no conditional formatting, no data validation, no
named ranges/defined names (workbook.xml is regenerated when the sheet set changes), no comments,
no print-area/page-setup, no external links, limited array/formula edge cases, ROW/COLUMN/OFFSET/
INDIRECT are partial. Print uses the browser's default sheet print.

### VPptx (`vpptx/`)

Implemented: PPTX read/write (slide size, shapes with EMU geometry/rotation, preset geometry, fill/
line, text paragraphs/runs, pictures, notes); ~20 preset shapes rendered as SVG; select/move/resize
with handles; double-click text editing; insert text box/shape/picture; slide add/duplicate/delete/
reorder; slide background; slideshow (F5, arrow/space/click, Esc); undo/redo; zoom; thumbnails; the
menu/toolbar/drawing-bar/ribbon chrome. A brand-new deck writes a minimal Office-valid master/
layout/theme/presentation part set.

Gaps vs Lectern: no SmartArt/diagrams, no charts, no tables editing (graphicFrame kept as raw), no
animations/transitions, no master/layout editing, no comments, no designer/ideas, limited theme
inheritance (layouts/masters kept but not applied for placeholder styling), no print handouts.

### Fidelity note

Saved files round-trip through the same reader with content preserved (verified against the
`tests/quire/lines/linecases.docx` corpus file for VWord, and via save→reopen for VExcel/VPptx).
Because the engines are a subset, files that use features outside the subset keep those parts
unchanged where possible, but editing such content is not supported. For full-fidelity work the
original `public/` apps remain the reference implementation.