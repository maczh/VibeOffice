import { createLegacyHostBridge, EngineEditorFrame, VibeOfficeProvider } from '@vibeoffice/core';
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
  onOpenFile?: (signal: AbortSignal) => Promise<LocalDocument | null>;
  initialFile?: File;
  initialDocument?: LocalDocument;
  onSaveFile?: (name: string, blob: Blob, current?: LocalDocument, signal?: AbortSignal) => Promise<LocalDocument | void | null>;
};

export function WordxEditor({ services, mountEngine, engine, draftStore, onDraftSaved, onDraftsAvailable, onEngineReady, onOpenFile, onSaveFile, initialFile, initialDocument, ...props }: WordxEditorProps) {
  const mount = useMemo<EditorMount | undefined>(() => {
    if (mountEngine) return mountEngine;
    if (!engine) return undefined;
    return createLegacyHostBridge({
      editor: 'wordx', engine, title: 'Wordx editor',
      accept: {
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx', '.docm', '.dotx', '.dotm'],
        'text/plain': ['.txt', '.md', '.markdown', '.rtf', '.html', '.htm', '.zip'],
      },
      ...(draftStore ? { draftStore } : {}),
      ...(onDraftSaved ? { onDraftSaved } : {}),
      ...(onDraftsAvailable ? { onDraftsAvailable } : {}),
      ...(onEngineReady ? { onEngineReady: (restoreDraft) => onEngineReady({ restoreDraft }) } : {}),
      ...(onOpenFile ? { onOpenFile } : {}), ...(onSaveFile ? { onSaveFile } : {}),
      ...(initialFile ? { initialFile } : {}), ...(initialDocument ? { initialDocument } : {}),
    });
  }, [mountEngine, engine, draftStore, onDraftSaved, onDraftsAvailable, onEngineReady, onOpenFile, onSaveFile, initialFile, initialDocument]);
  return <VibeOfficeProvider {...(services ? { services } : {})}><EngineEditorFrame appName="Wordx" editor="wordx" {...(mount ? { mount } : {})} {...props} /></VibeOfficeProvider>;
}
