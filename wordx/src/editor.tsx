import { createIndexedDbDraftStore, createLegacyEngineFrame, EngineEditorFrame, openLocalDocument, saveLocalDocument } from '@vibeoffice/core';
import { VibeOfficeProvider } from '@vibeoffice/core';
import type { VibeOfficeServiceOverrides } from '@vibeoffice/core';
import type { DraftStore, EditorMount, LegacyEngineFrameOptions, LocalDocument, LocalDraft } from '@vibeoffice/core';
import { useMemo, type ComponentProps } from 'react';

export interface WordxEngineControls {
  restoreDraft(draft: LocalDraft<Blob>): void;
}

export type WordxEditorProps = Omit<ComponentProps<typeof EngineEditorFrame>, 'appName' | 'editor'> & {
  services?: VibeOfficeServiceOverrides;
  /** Host the existing Quire engine here; omitted until the isolated browser runtime is configured. */
  mountEngine?: EditorMount;
  /** URL of the existing Quire browser runtime, served by the host application. */
  engine?: Omit<LegacyEngineFrameOptions, 'title'>;
  /** Optional isolated draft store; defaults to the React workspace draft database. */
  draftStore?: DraftStore<Blob>;
  onDraftSaved?: (draft: LocalDraft<Blob>) => void;
  onDraftsAvailable?: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void;
  onEngineReady?: (controls: WordxEngineControls) => void;
  onOpenFile?: () => Promise<LocalDocument | null>;
  initialFile?: File;
  initialDocument?: LocalDocument;
  onSaveFile?: (name: string, blob: Blob, current?: LocalDocument) => Promise<LocalDocument | void | null>;
};

export function WordxEditor({ services, mountEngine, engine, draftStore, onDraftSaved, onDraftsAvailable, onEngineReady, onOpenFile, onSaveFile, initialFile, initialDocument, ...props }: WordxEditorProps) {
  const mount = useMemo<EditorMount | undefined>(() => {
    if (mountEngine) return mountEngine;
    if (!engine) return undefined;
    const frameMount = createLegacyEngineFrame({ ...engine, title: 'Quire editor' });
    return ({ host, editor }) => {
      const expectedOrigin = new URL(engine.url, window.location.href).origin;
      if (expectedOrigin !== window.location.origin) throw new TypeError('The Quire runtime must be served from the host origin');
      const cleanup = frameMount({ host, editor });
      const frame = host.querySelector('iframe');
      if (!frame) return cleanup;
      let localDocument: LocalDocument | undefined;
      const receive = (event: MessageEvent) => {
        const message = event.data;
        if (event.source !== frame.contentWindow || event.origin !== expectedOrigin) return;
        if (message?.type === 'vibeoffice-react-ready') {
          const restoreDraft = (draft: LocalDraft<Blob>) => {
            if (!isDraft(draft)) throw new TypeError('Invalid Wordx draft');
            frame.contentWindow?.postMessage({ type: 'vibeoffice-react-restore', id: draft.id, name: draft.name, blob: draft.snapshot, lossState: draft.lossState }, expectedOrigin);
          };
          onEngineReady?.({ restoreDraft });
          void (draftStore ?? createIndexedDbDraftStore<Blob>()).list('wordx').then((drafts) => onDraftsAvailable?.(drafts, restoreDraft)).catch((error: unknown) => console.warn('Wordx draft list:', error));
          const initial = initialDocument ?? (initialFile ? { file: initialFile } : undefined);
          if (initial) {
            localDocument = initial;
            frame.contentWindow?.postMessage({ type: 'vibeoffice-react-open', file: initial.file, initial: true }, expectedOrigin);
          }
          return;
        }
        if (message?.type === 'vibeoffice-react-request-open') {
          void (async () => {
            try {
              const opened = onOpenFile ? await onOpenFile() : await openLocalDocument({
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx', '.docm', '.dotx', '.dotm'],
                'text/plain': ['.txt', '.md', '.markdown', '.rtf', '.html', '.htm', '.zip'],
              });
              if (!opened) return;
              localDocument = opened;
              frame.contentWindow?.postMessage({ type: 'vibeoffice-react-open', file: opened.file }, expectedOrigin);
            } catch (error) { if (!(error instanceof DOMException && error.name === 'AbortError')) console.warn('Wordx open:', error); }
          })();
          return;
        }
        if (message?.type === 'vibeoffice-react-request-save' && typeof message.requestId === 'string'
          && typeof message.name === 'string' && message.blob instanceof Blob) {
          void (async () => {
            let result = 'saved';
            try {
              if (onSaveFile) {
                const saved = await onSaveFile(message.name, message.blob, localDocument);
                if (saved === null) result = 'cancelled';
                else if (saved) localDocument = saved;
              } else {
                localDocument = await saveLocalDocument(localDocument ?? { file: new File([], message.name) }, message.blob, message.name);
              }
              frame.contentWindow?.postMessage({ type: 'vibeoffice-react-save-result', requestId: message.requestId, result }, expectedOrigin);
            } catch (error) {
              frame.contentWindow?.postMessage({ type: 'vibeoffice-react-save-result', requestId: message.requestId, error: error instanceof Error ? error.message : String(error) }, expectedOrigin);
            }
          })();
          return;
        }
        if (message?.type === 'vibeoffice-react-draft-remove' && typeof message.id === 'string') {
          void (draftStore ?? createIndexedDbDraftStore<Blob>()).remove(message.id).catch((error: unknown) => console.warn('Wordx draft removal:', error));
          return;
        }
        if (message?.type !== 'vibeoffice-react-draft' || !isDraft(message.draft)) return;
        const draft = message.draft as LocalDraft<Blob>;
        void (draftStore ?? createIndexedDbDraftStore<Blob>()).put(draft)
          .then(() => onDraftSaved?.(draft))
          .catch((error: unknown) => console.warn('Wordx draft storage:', error));
      };
      window.addEventListener('message', receive);
      return () => { window.removeEventListener('message', receive); if (typeof cleanup === 'function') cleanup(); };
    };
  }, [mountEngine, engine, draftStore, onDraftSaved, onDraftsAvailable, onEngineReady, onOpenFile, onSaveFile, initialFile, initialDocument]);
  return <VibeOfficeProvider {...(services ? { services } : {})}><EngineEditorFrame appName="Quire" editor="wordx" {...(mount ? { mount } : {})} {...props} /></VibeOfficeProvider>;
}

function isDraft(value: unknown): value is LocalDraft<Blob> {
  if (!value || typeof value !== 'object') return false;
  const draft = value as Partial<LocalDraft<Blob>>;
  return typeof draft.id === 'string' && draft.editor === 'wordx' && typeof draft.name === 'string'
    && typeof draft.updatedAt === 'number' && draft.snapshot instanceof Blob;
}
