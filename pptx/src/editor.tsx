import { createLegacyHostBridge, EngineEditorFrame } from '@vibeoffice/core';
import { VibeOfficeProvider } from '@vibeoffice/core';
import type { VibeOfficeServiceOverrides } from '@vibeoffice/core';
import type { DraftStore, LegacyEngineFrameOptions, LocalDocument, LocalDraft } from '@vibeoffice/core';
import { useMemo, type ComponentProps } from 'react';

export type PptxEditorProps = Omit<ComponentProps<typeof EngineEditorFrame>, 'appName' | 'editor'> & { services?: VibeOfficeServiceOverrides; engine?: Omit<LegacyEngineFrameOptions, 'title'>; draftStore?: DraftStore<Blob>; onDraftSaved?: (draft: LocalDraft<Blob>) => void; onEngineReady?: (restore: (draft: LocalDraft<Blob>) => void) => void; onDraftsAvailable?: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void; onOpenFile?: () => Promise<LocalDocument | null>; onSaveFile?: (name: string, blob: Blob, current?: LocalDocument) => Promise<LocalDocument | void | null>; initialFile?: File; initialDocument?: LocalDocument };

export function PptxEditor({ services, engine, draftStore, onDraftSaved, onEngineReady, onDraftsAvailable, onOpenFile, onSaveFile, initialFile, initialDocument, ...props }: PptxEditorProps) {
  const mount = useMemo(() => engine ? createLegacyHostBridge({ editor: 'pptx', engine, title: 'Lectern editor', accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm'], 'text/plain': ['.txt'] }, ...(draftStore ? { draftStore } : {}), ...(onDraftSaved ? { onDraftSaved } : {}), ...(onEngineReady ? { onEngineReady } : {}), ...(onDraftsAvailable ? { onDraftsAvailable } : {}), ...(onOpenFile ? { onOpenFile } : {}), ...(onSaveFile ? { onSaveFile } : {}), ...(initialFile ? { initialFile } : {}), ...(initialDocument ? { initialDocument } : {}) }) : undefined, [engine, draftStore, onDraftSaved, onEngineReady, onDraftsAvailable, onOpenFile, onSaveFile, initialFile, initialDocument]);
  return <VibeOfficeProvider {...(services ? { services } : {})}><EngineEditorFrame appName="Lectern" editor="pptx" {...(mount ? { mount } : {})} {...props} /></VibeOfficeProvider>;
}
