import { createIndexedDbDraftStore, openLocalDocument, saveLocalDocument, type DraftStore, type EditorMount, type LegacyEngineFrameOptions, type LocalDocument, type LocalDraft } from './index.js';

export interface LegacyHostOptions {
  editor: 'wordx' | 'excelx' | 'pptx';
  engine: Omit<LegacyEngineFrameOptions, 'title'>;
  title: string;
  accept: Record<string, string[]>;
  draftStore?: DraftStore<Blob>;
  onDraftSaved?: (draft: LocalDraft<Blob>) => void;
  onOpenFile?: () => Promise<LocalDocument | null>;
  initialFile?: File;
  initialDocument?: LocalDocument;
  onSaveFile?: (name: string, blob: Blob, current?: LocalDocument) => Promise<LocalDocument | void | null>;
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
  return ({ host }) => {
    const cleanup = mountFrame({ host, editor });
    const frame = host.querySelector('iframe');
    if (!frame) return cleanup;
    let current: LocalDocument | undefined;
    const restoreDraft = (draft: LocalDraft<Blob>) => {
      if (!isDraft(draft, editor)) throw new TypeError(`Invalid ${editor} draft`);
      frame.contentWindow?.postMessage({ type: 'vibeoffice-react-restore', id: draft.id, name: draft.name, blob: draft.snapshot, lossState: draft.lossState }, source.origin);
    };
    const exposeDrafts = () => {
      options.onEngineReady?.(restoreDraft);
      void store.list(editor).then((drafts) => options.onDraftsAvailable?.(drafts, restoreDraft)).catch((error: unknown) => console.warn(`${editor} draft list:`, error));
    };
    const receive = (event: MessageEvent) => {
      const message = event.data;
      if (event.source !== frame.contentWindow || event.origin !== source.origin) return;
      if (message?.type === 'vibeoffice-react-ready') {
        const initial = options.initialDocument ?? (options.initialFile ? { file: options.initialFile } : undefined);
        if (initial) {
          current = initial;
          frame.contentWindow?.postMessage({ type: 'vibeoffice-react-open', file: initial.file, initial: true }, source.origin);
        }
        exposeDrafts();
      } else if (message?.type === 'vibeoffice-react-request-open') {
        void (async () => {
          try {
            const opened = options.onOpenFile ? await options.onOpenFile() : await openLocalDocument(accept);
            if (!opened) return;
            current = opened;
            frame.contentWindow?.postMessage({ type: 'vibeoffice-react-open', file: opened.file }, source.origin);
          } catch (error) { if (!(error instanceof DOMException && error.name === 'AbortError')) console.warn(`${editor} open:`, error); }
        })();
      } else if (message?.type === 'vibeoffice-react-request-save' && typeof message.requestId === 'string' && typeof message.name === 'string' && message.blob instanceof Blob) {
        void (async () => {
          try {
            const saved = options.onSaveFile ? await options.onSaveFile(message.name, message.blob, current)
              : await saveLocalDocument(current ?? { file: new File([], message.name) }, message.blob, message.name);
            frame.contentWindow?.postMessage({ type: 'vibeoffice-react-save-result', requestId: message.requestId, result: saved === null ? 'cancelled' : 'saved' }, source.origin);
            if (saved) current = saved;
          } catch (error) { frame.contentWindow?.postMessage({ type: 'vibeoffice-react-save-result', requestId: message.requestId, error: error instanceof Error ? error.message : String(error) }, source.origin); }
        })();
      } else if (message?.type === 'vibeoffice-react-draft-remove' && typeof message.id === 'string') {
        void store.remove(message.id).catch((error: unknown) => console.warn(`${editor} draft removal:`, error));
      } else if (message?.type === 'vibeoffice-react-draft' && isDraft(message.draft, editor)) {
        const draft = message.draft as LocalDraft<Blob>;
        void store.put(draft).then(() => options.onDraftSaved?.(draft)).catch((error: unknown) => console.warn(`${editor} draft storage:`, error));
      }
    };
    frame.addEventListener('load', exposeDrafts, { once: true });
    window.addEventListener('message', receive);
    return () => { frame.removeEventListener('load', exposeDrafts); window.removeEventListener('message', receive); if (typeof cleanup === 'function') cleanup(); };
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
