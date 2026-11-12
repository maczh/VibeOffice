import { createDirectEngineBridge, EngineEditorFrame } from '@vibeoffice/core';
import { VibeOfficeProvider } from '@vibeoffice/core';
import type { VibeOfficeServiceOverrides } from '@vibeoffice/core';
import type { DraftStore, EngineRuntimeOptions, LocalDocument, LocalDraft, SaveFileOperation } from '@vibeoffice/core';
import { useMemo, type ComponentProps } from 'react';
import { PptxNotesLayer } from './notes.js';
import { PptxNotesPageLayer } from './notes-page.js';
import { PptxPrintPreviewLayer } from './print-preview.js';
import { PptxShowControlsLayer } from './show-controls.js';
import { PptxSlidesLayer } from './slides-pane.js';
import { PptxTaskContentLayer } from './task-content.js';

export type PptxEditorProps = Omit<ComponentProps<typeof EngineEditorFrame>, 'appName' | 'editor'> & { services?: VibeOfficeServiceOverrides; engine?: EngineRuntimeOptions; draftStore?: DraftStore<Blob>; onDraftSaved?: (draft: LocalDraft<Blob>) => void; onEngineReady?: (restore: (draft: LocalDraft<Blob>) => void) => void; onDraftsAvailable?: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void; onOpenFile?: (signal: AbortSignal) => Promise<LocalDocument | null>; allowedRemoteOrigins?: readonly string[]; onOpenUrl?: (url: string, signal: AbortSignal) => Promise<LocalDocument | null>; onSaveFile?: (name: string, blob: Blob, current?: LocalDocument, signal?: AbortSignal, operation?: SaveFileOperation) => Promise<LocalDocument | void | null>; initialFile?: File; initialDocument?: LocalDocument };

export function PptxEditor({ services, engine, draftStore, onDraftSaved, onEngineReady, onDraftsAvailable, onOpenFile, allowedRemoteOrigins, onOpenUrl, onSaveFile, initialFile, initialDocument, ...props }: PptxEditorProps) {
  const mount = useMemo(() => engine ? createDirectEngineBridge({ editor: 'pptx', engine, title: 'PPTX editor', accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm'], 'text/plain': ['.txt'] }, ...(draftStore ? { draftStore } : {}), ...(onDraftSaved ? { onDraftSaved } : {}), ...(onEngineReady ? { onEngineReady } : {}), ...(onDraftsAvailable ? { onDraftsAvailable } : {}), ...(onOpenFile ? { onOpenFile } : {}), ...(allowedRemoteOrigins ? { allowedRemoteOrigins } : {}), ...(onOpenUrl ? { onOpenUrl } : {}), ...(onSaveFile ? { onSaveFile } : {}), ...(initialFile ? { initialFile } : {}), ...(initialDocument ? { initialDocument } : {}) }) : undefined, [engine, draftStore, onDraftSaved, onEngineReady, onDraftsAvailable, onOpenFile, allowedRemoteOrigins, onOpenUrl, onSaveFile, initialFile, initialDocument]);
  return <VibeOfficeProvider {...(services ? { services } : {})}><EngineEditorFrame appName="PPTX" editor="pptx" {...(mount ? { mount } : {})} {...props} appLayers={(host) => <><PptxNotesLayer host={host} /><PptxNotesPageLayer host={host} /><PptxPrintPreviewLayer host={host} /><PptxShowControlsLayer host={host} /><PptxSlidesLayer host={host} /><PptxTaskContentLayer host={host} />{props.appLayers?.(host)}</>} /></VibeOfficeProvider>;
}
