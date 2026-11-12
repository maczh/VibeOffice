import { createIndexedDbDraftStore, createUrlReadProvider, openLocalDocument, saveLocalDocument, type DraftStore, type EditorMount, type LocalDocument, type LocalDraft } from './index.js';
import { mountLegacyEngineDocument } from './direct-engine.js';

export interface DirectHostOptions {
  editor: 'wordx' | 'excelx' | 'pptx';
  engine: { url: string | URL; launch?: 'new' | 'sample' };
  title?: string;
  accept: Record<string, string[]>;
  allowedRemoteOrigins?: readonly string[];
  draftStore?: DraftStore<Blob>;
  onDraftSaved?: (draft: LocalDraft<Blob>) => void;
  onOpenFile?: (signal: AbortSignal) => Promise<LocalDocument | null>;
  onOpenUrl?: (url: string, signal: AbortSignal) => Promise<LocalDocument | null>;
  initialFile?: File;
  initialDocument?: LocalDocument;
  onSaveFile?: (name: string, blob: Blob, current?: LocalDocument, signal?: AbortSignal, operation?: SaveFileOperation) => Promise<LocalDocument | void | null>;
  onEngineReady?: (restore: (draft: LocalDraft<Blob>) => void) => void;
  onDraftsAvailable?: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void;
}

/** Context passed to host save callbacks so Save As never silently becomes an in-place remote overwrite. */
export interface SaveFileOperation { saveAs: boolean }

type EngineHooks = {
  open?: (file: File, options?: Record<string, unknown>) => Promise<void> | void;
  close?: () => Promise<boolean> | boolean;
  current?: () => unknown;
  isDirty?: (document: unknown) => boolean;
  message?: (text: string) => void;
};

/** Mount the existing engine in the host document and route file operations without postMessage. */
export function createDirectEngineBridge(options: DirectHostOptions): EditorMount {
  const store = options.draftStore ?? createIndexedDbDraftStore<Blob>();
  return async ({ host, signal }) => {
    const view = host.ownerDocument.defaultView;
    if (!view) throw new Error('The editor host document has no window');
    const source = new URL(options.engine.url, view.location.href);
    if (source.origin !== view.location.origin) throw new TypeError('The editor runtime must be served from the host origin');
    const remoteOrigins = options.allowedRemoteOrigins ?? [view.location.origin];
    const allowedRemoteOrigins = new Set(remoteOrigins.map((origin) => new URL(origin).origin));
    const urlProvider = createUrlReadProvider(remoteOrigins, view.fetch.bind(view));
    let disposed = false;
    let current: LocalDocument | undefined;
    let hooks: EngineHooks | undefined;
    let stopEngine: void | (() => void);
    const live = () => !disposed && !signal.aborted;
    const restore = (draft: LocalDraft<Blob>) => {
      if (!isDraft(draft, options.editor)) throw new TypeError(`Invalid ${options.editor} draft`);
      if (!hooks?.open) return;
      void hooks.open(new File([draft.snapshot], draft.name), { draft: true, name: draft.name, saved: false, lossState: draft.lossState });
    };
    const template = source.searchParams.get('template');
    const save = async (name: string, blob: Blob, saveAs: boolean) => {
      if (!live()) throw new DOMException('The editor mount was cancelled', 'AbortError');
      const result = options.onSaveFile
        ? await options.onSaveFile(name, blob, current, signal, saveAs ? { saveAs: true } : undefined)
        : await saveLocalDocument(saveAs ? { file: new File([], name) } : current ?? { file: new File([], name) }, blob, name, view, signal);
      if (!live()) throw new DOMException('The editor mount was cancelled', 'AbortError');
      if (result === null) return 'cancelled';
      if (result) current = result;
      return 'saved';
    };
    const runtime = {
      launch: options.engine.launch === 'sample' ? null : template ? { template } : { new: true },
      async open() {
        if (!live()) return;
        const file = options.onOpenFile ? await options.onOpenFile(signal) : await openLocalDocument(options.accept, view, signal);
        if (!file || !live() || (hooks?.close && !(await hooks.close())) || !live()) return;
        current = file;
        await hooks?.open?.(file.file);
      },
      async openUrl(value: string) {
        if (!live()) return;
        const url = new URL(value);
        if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password) throw new TypeError('Only credential-free HTTP and HTTPS document URLs can be opened');
        if (!allowedRemoteOrigins.has(url.origin)) throw new TypeError('URL origin is not allowed');
        const file = options.onOpenUrl
          ? await options.onOpenUrl(url.href, signal)
          : await urlProvider.read(url.href, signal).then(({ file: remoteFile, blob }) => ({
            file: new File([blob], remoteFile.name, { type: remoteFile.mimeType || blob.type }),
            remote: { provider: urlProvider, file: remoteFile },
          }));
        if (!file || !live() || (hooks?.close && !(await hooks.close())) || !live()) return;
        current = file;
        await hooks?.open?.(file.file, { url: url.href });
      },
      save(name: string, blob: Blob) { return save(name, blob, false); },
      saveAs(name: string, blob: Blob) { return save(name, blob, true); },
      async saveDraft(draft: LocalDraft<Blob>) {
        if (!live() || !isDraft(draft, options.editor)) return;
        await store.put(draft, signal);
        if (live()) options.onDraftSaved?.(draft);
      },
      async removeDraft(id: string) {
        if (live()) await store.remove(id, signal);
      },
      ready(_app: string, mountedHooks: EngineHooks) {
        if (!live()) return;
        hooks = mountedHooks;
        const initial = options.initialDocument ?? (options.initialFile ? { file: options.initialFile } : undefined);
        if (initial) {
          current = initial;
          void mountedHooks.open?.(initial.file, { initial: true });
        }
        options.onEngineReady?.(restore);
        void store.list(options.editor, signal).then((drafts) => {
          if (live()) options.onDraftsAvailable?.(drafts, restore);
        }).catch((error: unknown) => { if (live()) console.warn(`${options.editor} draft list:`, error); });
      },
    };
    const cleanupDocument = await mountLegacyEngineDocument(source, { host, signal, directHost: runtime });
    const root = host.querySelector<HTMLElement>('[data-vibeoffice-engine-root]');
    const runtimeWindow = view as Window & { L?: { app?: { mount?: (element: HTMLElement) => void | (() => void) | Promise<void | (() => void)>; newDocument?: () => void; newWorkbook?: () => void; newPresentation?: () => void }; ui?: { exec?(id: string): unknown; searchResults?(query: string, layout: string): unknown[] } } };
    if (!root || !runtimeWindow.L?.app?.mount) { cleanupDocument(); throw new Error(`The ${options.editor} runtime does not expose a direct mount API`); }
    const onCommand = (event: Event) => {
      const command = (event as CustomEvent<{ command?: unknown }>).detail?.command;
      const app = runtimeWindow.L?.app;
      if (command === 'new') {
        const create = options.editor === 'wordx' ? app?.newDocument : options.editor === 'excelx' ? app?.newWorkbook : app?.newPresentation;
        create?.call(app);
      } else if (typeof command === 'string') runtimeWindow.L?.ui?.exec?.(command);
    };
    const onChrome = (event: Event) => {
      const detail = (event as CustomEvent<{ look?: string; language?: string }>).detail;
      const runtimeWindow = view as Window & { VO?: { look?: { set(id: string): void }; lang?: { set(id: string): void } } };
      if (detail.look) runtimeWindow.VO?.look?.set(detail.look);
      if (detail.language) runtimeWindow.VO?.lang?.set(detail.language);
    };
    const onOpenUrl = (event: Event) => {
      const detail = (event as CustomEvent<{ url?: unknown; respond?: (error?: string) => void }>).detail;
      const url = detail?.url;
      if (typeof url !== 'string' || !url.trim()) return;
      void runtime.openUrl(url).then(() => detail?.respond?.()).catch((error: unknown) => {
        const message = error instanceof Error ? error.message : String(error);
        if (live()) {
          if (detail?.respond) detail.respond(message);
          else hooks?.message?.(message);
        }
      });
    };
    const onSearch = (event: Event) => {
      const detail = (event as CustomEvent<{ query?: unknown; layout?: unknown; respond?: (results: unknown[]) => void }>).detail;
      if (typeof detail?.query !== 'string' || typeof detail.respond !== 'function') return;
      detail.respond(runtimeWindow.L?.ui?.searchResults?.(detail.query, typeof detail.layout === 'string' ? detail.layout : 'bars') ?? []);
    };
    host.addEventListener('vibeoffice-react-command', onCommand);
    host.addEventListener('vibeoffice-react-chrome', onChrome);
    host.addEventListener('vibeoffice-react-open-url', onOpenUrl);
    host.addEventListener('vibeoffice-react-search', onSearch);
    try { stopEngine = await runtimeWindow.L.app.mount(root); }
    catch (error) {
      host.removeEventListener('vibeoffice-react-command', onCommand);
      host.removeEventListener('vibeoffice-react-chrome', onChrome);
      host.removeEventListener('vibeoffice-react-open-url', onOpenUrl);
      host.removeEventListener('vibeoffice-react-search', onSearch);
      cleanupDocument();
      throw error;
    }
    const dispose = () => {
      if (disposed) return;
      disposed = true;
      host.removeEventListener('vibeoffice-react-command', onCommand);
      host.removeEventListener('vibeoffice-react-chrome', onChrome);
      host.removeEventListener('vibeoffice-react-open-url', onOpenUrl);
      host.removeEventListener('vibeoffice-react-search', onSearch);
      if (typeof stopEngine === 'function') stopEngine();
      cleanupDocument();
    };
    signal.addEventListener('abort', dispose, { once: true });
    return () => { signal.removeEventListener('abort', dispose); dispose(); };
  };
}

function isDraft(value: unknown, editor: string): value is LocalDraft<Blob> {
  if (!value || typeof value !== 'object') return false;
  const draft = value as Partial<LocalDraft<Blob>>;
  return draft.editor === editor && typeof draft.id === 'string' && typeof draft.name === 'string'
    && typeof draft.updatedAt === 'number' && draft.snapshot instanceof Blob;
}
