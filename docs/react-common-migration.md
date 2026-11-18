# React shared migration map

`rshared/src/common.tsx` is the React counterpart of the Vue shared migration layer. It exposes
functional components and hooks without loading the classic `public/common` scripts.

| Vue migration | React counterpart |
|---|---|
| `useLook()` | `useLook()` |
| `useRecentFiles()` / `VoStartCenter.vue` | `useRecentFiles()` / `StartCenter` |
| `useI18n()` / `VoLanguageSelect.vue` | `useI18n()` / `LanguageSelect` |
| `usePwa()` / `VoPwaUpdate.vue` | `usePwa()` / `PwaUpdate` |
| `useAssistant()` / `VoAssistant.vue` | `useAssistant()` / `Assistant` |
| `VoFileDropZone.vue` | `FileDropZone` |
| `useDirtyGuard()` | `useDirtyGuard()` |
| `VoMenu`/`VoContextMenu` and command UI | `ContextMenu`, `MenuBar`, `Toolbar`, `Ribbon` |
| `VoCommandSearch` | `CommandSearch` |

`OnlineSourceDialog` is the typed React entry for WebDAV, S3/MinIO, presigned URL and HTTP POST
sources/targets. The existing `OnlineDialog` remains as the backwards-compatible URL-only API.

`commandsFromMenus()` converts the existing declarative menu definitions into the command list used
by `CommandSearch`, keeping menu execution as the single source of truth.

All hooks clean up browser listeners in `useEffect`. Components use semantic buttons, labels and
ARIA roles, and can be imported from `@voffice/react-shared` individually.

The integrated React host is `rplayground/`:

```bash
npm run dev:react
```

It runs all three React editors in one page with tabs, a Start Center, recent files, look switching
and local drag/drop opening.
