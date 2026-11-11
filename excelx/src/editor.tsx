import { createLegacyHostBridge, EngineEditorFrame } from '@vibeoffice/core';
import { VibeOfficeProvider } from '@vibeoffice/core';
import type { VibeOfficeServiceOverrides } from '@vibeoffice/core';
import type { DraftStore, LegacyEngineFrameOptions, LocalDocument, LocalDraft } from '@vibeoffice/core';
import { useMemo, type ComponentProps } from 'react';

export type ExcelxEditorProps = Omit<ComponentProps<typeof EngineEditorFrame>, 'appName' | 'editor'> & { services?: VibeOfficeServiceOverrides; engine?: Omit<LegacyEngineFrameOptions, 'title'>; draftStore?: DraftStore<Blob>; onDraftSaved?: (draft: LocalDraft<Blob>) => void; onEngineReady?: (restore: (draft: LocalDraft<Blob>) => void) => void; onDraftsAvailable?: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void; onOpenFile?: (signal: AbortSignal) => Promise<LocalDocument | null>; onSaveFile?: (name: string, blob: Blob, current?: LocalDocument, signal?: AbortSignal) => Promise<LocalDocument | void | null>; initialFile?: File; initialDocument?: LocalDocument };

export function ExcelxEditor({ services, engine, draftStore, onDraftSaved, onEngineReady, onDraftsAvailable, onOpenFile, onSaveFile, initialFile, initialDocument, ...props }: ExcelxEditorProps) {
  const mount = useMemo(() => engine ? createLegacyHostBridge({ editor: 'excelx', engine, title: 'Excelx editor', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm', '.xltx', '.xltm'], 'text/plain': ['.csv', '.txt', '.tsv', '.prn', '.xml'] }, ...(draftStore ? { draftStore } : {}), ...(onDraftSaved ? { onDraftSaved } : {}), ...(onEngineReady ? { onEngineReady } : {}), ...(onDraftsAvailable ? { onDraftsAvailable } : {}), ...(onOpenFile ? { onOpenFile } : {}), ...(onSaveFile ? { onSaveFile } : {}), ...(initialFile ? { initialFile } : {}), ...(initialDocument ? { initialDocument } : {}) }) : undefined, [engine, draftStore, onDraftSaved, onEngineReady, onDraftsAvailable, onOpenFile, onSaveFile, initialFile, initialDocument]);
  return <VibeOfficeProvider {...(services ? { services } : {})}><EngineEditorFrame appName="Excelx" editor="excelx" {...(mount ? { mount } : {})} {...props} /></VibeOfficeProvider>;
}
