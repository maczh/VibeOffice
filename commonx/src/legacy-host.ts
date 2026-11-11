import { createIndexedDbDraftStore, openLocalDocument, saveLocalDocument, type DraftStore, type EditorMount, type LegacyEngineFrameOptions, type LocalDocument, type LocalDraft } from './index.js';

export interface LegacyHostOptions {
  editor: 'wordx' | 'excelx' | 'pptx';
  engine: Omit<LegacyEngineFrameOptions, 'title'>;
  title: string;
  accept: Record<string, string[]>;
  draftStore?: DraftStore<Blob>;
  onDraftSaved?: (draft: LocalDraft<Blob>) => void;
  onOpenFile?: (signal: AbortSignal) => Promise<LocalDocument | null>;
  initialFile?: File;
  initialDocument?: LocalDocument;
  onSaveFile?: (name: string, blob: Blob, current?: LocalDocument, signal?: AbortSignal) => Promise<LocalDocument | void | null>;
  onEngineReady?: (restore: (draft: LocalDraft<Blob>) => void) => void;
  onDraftsAvailable?: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void;
}

/** Adapt the existing document engines' same-origin postMessage protocol to local files and isolated crash drafts. */
export function createLegacyHostBridge(options: LegacyHostOptions): EditorMount {
  const { editor, engine, accept, title } = options;
  const source = new URL(engine.url, window.location.href);
  if (source.origin !== window.location.origin) throw new TypeError('The editor runtime must be served from the host origin');
  const mountFrame = createLegacyEngineFrame({ ...engine, title });
  const store = options.draftStore ?? createIndexedDbDraftStore<Blob>();
  return ({ host, signal }) => {
    if (signal.aborted) return;
    const cleanup = mountFrame({ host, editor, signal });
    const frame = host.querySelector('iframe');
    if (!frame) return cleanup;
    let disposed = false;
    let current: LocalDocument | undefined;
    let engineInitialized = false;
    const pendingMessages: Record<string, unknown>[] = [];
    const postToEngine = (message: Record<string, unknown>) => {
      if (disposed || signal.aborted) return;
      if (!engineInitialized) pendingMessages.push(message);
      else frame.contentWindow?.postMessage(message, source.origin);
    };
    const restoreDraft = (draft: LocalDraft<Blob>) => {
      if (!isDraft(draft, editor)) throw new TypeError(`Invalid ${editor} draft`);
      postToEngine({ type: 'vibeoffice-react-restore', id: draft.id, name: draft.name, blob: draft.snapshot, lossState: draft.lossState });
    };
    let readinessTimer: number | undefined;
    const exposeDrafts = () => {
      if (disposed || signal.aborted) return;
      options.onEngineReady?.(restoreDraft);
      void store.list(editor, signal).then((drafts) => { if (!disposed && !signal.aborted) options.onDraftsAvailable?.(drafts, restoreDraft); }).catch((error: unknown) => { if (!disposed && !signal.aborted) console.warn(`${editor} draft list:`, error); });
    };
    const initializeEngine = () => {
      if (engineInitialized || disposed || signal.aborted) return;
      engineInitialized = true;
      if (readinessTimer !== undefined) window.clearInterval(readinessTimer);
      const initial = options.initialDocument ?? (options.initialFile ? { file: options.initialFile } : undefined);
      if (initial) {
        current = initial;
        pendingMessages.unshift({ type: 'vibeoffice-react-open', file: initial.file, initial: true });
      }
      for (const message of pendingMessages.splice(0)) frame.contentWindow?.postMessage(message, source.origin);
      exposeDrafts();
    };
    const sendCommand = (event: Event) => {
      const command = (event as CustomEvent<{ command?: unknown }>).detail?.command;
      if (typeof command === 'string' && /^[a-z][a-zA-Z0-9]*$/.test(command)) {
        postToEngine({ type: 'vibeoffice-react-command', command });
      }
    };
    const syncChrome = (event: Event) => {
      const detail = (event as CustomEvent<{ look?: unknown; language?: unknown }>).detail;
      postToEngine({ type: 'vibeoffice-react-chrome', look: detail?.look, language: detail?.language });
    };
    const receive = (event: MessageEvent) => {
      const message = event.data;
      if (event.source !== frame.contentWindow || event.origin !== source.origin) return;
      if (message?.type === 'vibeoffice-react-ready') {
        initializeEngine();
      } else if (message?.type === 'vibeoffice-react-request-open') {
        void (async () => {
          try {
            const opened = options.onOpenFile ? await options.onOpenFile(signal) : await openLocalDocument(accept, window, signal);
            if (!opened || disposed || signal.aborted) return;
            current = opened;
            postToEngine({ type: 'vibeoffice-react-open', file: opened.file });
          } catch (error) { if (!disposed && !signal.aborted && !(error instanceof DOMException && error.name === 'AbortError')) console.warn(`${editor} open:`, error); }
        })();
      } else if (message?.type === 'vibeoffice-react-request-save' && typeof message.requestId === 'string' && typeof message.name === 'string' && message.blob instanceof Blob) {
        void (async () => {
          try {
            const saved = options.onSaveFile ? await options.onSaveFile(message.name, message.blob, current, signal)
              : await saveLocalDocument(current ?? { file: new File([], message.name) }, message.blob, message.name, window, signal);
            if (disposed || signal.aborted) return;
            frame.contentWindow?.postMessage({ type: 'vibeoffice-react-save-result', requestId: message.requestId, result: saved === null ? 'cancelled' : 'saved' }, source.origin);
            if (saved) current = saved;
          } catch (error) { if (!disposed && !signal.aborted) frame.contentWindow?.postMessage({ type: 'vibeoffice-react-save-result', requestId: message.requestId, error: error instanceof Error ? error.message : String(error) }, source.origin); }
        })();
      } else if (message?.type === 'vibeoffice-react-draft-remove' && typeof message.id === 'string') {
        void store.remove(message.id, signal).catch((error: unknown) => { if (!disposed && !signal.aborted) console.warn(`${editor} draft removal:`, error); });
      } else if (message?.type === 'vibeoffice-react-draft' && isDraft(message.draft, editor)) {
        const draft = message.draft as LocalDraft<Blob>;
        void store.put(draft, signal).then(() => { if (!disposed && !signal.aborted) options.onDraftSaved?.(draft); }).catch((error: unknown) => { if (!disposed && !signal.aborted) console.warn(`${editor} draft storage:`, error); });
      }
    };
    const onFrameLoad = () => {
      const child = frame.contentWindow as (Window & { VO?: { reactHost?: boolean; reactReady?: boolean } }) | null;
      if (child?.VO?.reactHost && child.VO.reactReady) initializeEngine();
    };
    frame.addEventListener('load', onFrameLoad, { once: true });
    readinessTimer = window.setInterval(onFrameLoad, 50);
    window.addEventListener('message', receive);
    host.addEventListener('vibeoffice-react-command', sendCommand);
    host.addEventListener('vibeoffice-react-chrome', syncChrome);
    const dispose = () => {
      if (disposed) return;
      disposed = true;
      if (readinessTimer !== undefined) window.clearInterval(readinessTimer);
      frame.removeEventListener('load', onFrameLoad);
      window.removeEventListener('message', receive);
      host.removeEventListener('vibeoffice-react-command', sendCommand);
      host.removeEventListener('vibeoffice-react-chrome', syncChrome);
      signal.removeEventListener('abort', dispose);
      if (typeof cleanup === 'function') cleanup();
    };
    signal.addEventListener('abort', dispose, { once: true });
    return dispose;
  };
}

function isDraft(value: unknown, editor: string): value is LocalDraft<Blob> {
  if (!value || typeof value !== 'object') return false;
  const draft = value as Partial<LocalDraft<Blob>>;
  return draft.editor === editor && typeof draft.id === 'string' && typeof draft.name === 'string'
    && typeof draft.updatedAt === 'number' && draft.snapshot instanceof Blob;
}

// Imported from the barrel after type checking to keep this adapter's public imports concise.
import { createLegacyEngineFrame } from './editor-frame.js';
