import { createDirectEngineBridge, EngineEditorFrame, VibeOfficeProvider } from '@vibeoffice/core';
import type { VibeOfficeServiceOverrides } from '@vibeoffice/core';
import type { DraftStore, EditorMount, EngineRuntimeOptions, LocalDocument, LocalDraft, SaveFileOperation } from '@vibeoffice/core';
import { installCommandSearchService } from '@vibeoffice/core';
import { useMemo, type ComponentProps } from 'react';
import { WordxListBoxLayer } from './list-box.js';
import { WordxMenuContentLayer } from './menu-content.js';
import { installWordxRibbonSpec } from './ribbon-spec.js';
import { WordxTaskContentLayer } from './task-content.js';
import { installWordxTemplates } from './templates.js';
import { installWordxLineBoxes } from './line-boxes.js';
import { installWordxDrawTool } from './draw-tool.js';
import { installWordxOutline } from './outline.js';
import { installWordxAutoCorrect } from './autocorrect.js';
import { installWordxRulers } from './rulers.js';
import { installWordxOmml } from './omml.js';
import { installWordxMailMerge } from './mail-merge.js';
import { installWordxFind } from './find.js';
import { installWordxProofing } from './proofing.js';
import { installWordxMarkdown } from './markdown.js';
import { installWordxMarkdownIo } from './markdown-io.js';
import { installWordxHtmlIo } from './html-io.js';
import { installWordxRtf } from './rtf-io.js';
import { installWordxDocxRead } from './docx-read.js';
import { installWordxDocxWrite } from './docx-write.js';
import { installWordxFields } from './fields.js';
import { installWordxCommands } from './commands.js';
import { installWordxOperations } from './operations.js';
import { installWordxDocumentModel } from './dmodel.js';
import { installWordxDrawing } from './drawing.js';
import { installWordxDrawingPreservation } from './preserve-drawing.js';
import { installWordxPreservation } from './preservation.js';
import { WordxEngineShell } from './engine-shell.js';

export interface WordxEngineControls {
  restoreDraft(draft: LocalDraft<Blob>): void;
}
export type WordxEditorProps = Omit<ComponentProps<typeof EngineEditorFrame>, 'appName' | 'editor'> & {
  services?: VibeOfficeServiceOverrides;
  /** Host an existing Wordx browser engine here. */
  mountEngine?: EditorMount;
  /** URL of the existing Wordx browser runtime, served by the host application. */
  engine?: EngineRuntimeOptions;
  /** Optional isolated draft store; defaults to the React workspace draft database. */
  draftStore?: DraftStore<Blob>;
  onDraftSaved?: (draft: LocalDraft<Blob>) => void;
  onTitleChange?: (title: string) => void;
  onDraftsAvailable?: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void;
  onEngineReady?: (controls: WordxEngineControls) => void;
  onOpenFile?: (signal: AbortSignal) => Promise<LocalDocument | null>;
  allowedRemoteOrigins?: readonly string[];
  onOpenUrl?: (url: string, signal: AbortSignal) => Promise<LocalDocument | null>;
  initialFile?: File;
  initialDocument?: LocalDocument;
  onSaveFile?: (name: string, blob: Blob, current?: LocalDocument, signal?: AbortSignal, operation?: SaveFileOperation) => Promise<LocalDocument | void | null>;
};

export function WordxEditor({ services, mountEngine, engine, draftStore, onDraftSaved, onTitleChange, onDraftsAvailable, onEngineReady, onOpenFile, allowedRemoteOrigins, onOpenUrl, onSaveFile, initialFile, initialDocument, ...props }: WordxEditorProps) {
  const mount = useMemo<EditorMount | undefined>(() => {
    if (mountEngine) return mountEngine;
    if (!engine) return undefined;
    return createDirectEngineBridge({
      editor: 'wordx', engine, title: 'Wordx editor',
      shellOwnedByReact: true,
      accept: {
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx', '.docm', '.dotx', '.dotm'],
        'text/plain': ['.txt', '.md', '.markdown', '.rtf', '.html', '.htm', '.zip'],
      },
      beforeLegacyScript: (runtime, script) => { if (script.pathname.endsWith('/common/search.js')) installCommandSearchService(runtime); if (script.pathname.endsWith('/wordx/js/dmodel.js')) installWordxDocumentModel(runtime); if (script.pathname.endsWith('/wordx/js/ops.js')) installWordxOperations(runtime); if (script.pathname.endsWith('/wordx/js/docx-read.js')) installWordxDocxRead(runtime); if (script.pathname.endsWith('/wordx/js/docx-write.js')) installWordxDocxWrite(runtime); if (script.pathname.endsWith('/wordx/js/fields.js')) installWordxFields(runtime); if (script.pathname.endsWith('/wordx/js/commands.js')) installWordxCommands(runtime); if (script.pathname.endsWith('/wordx/js/htmlio.js')) installWordxHtmlIo(runtime); if (script.pathname.endsWith('/wordx/js/drawing.js')) installWordxDrawing(runtime); if (script.pathname.endsWith('/wordx/js/preserve.js')) installWordxPreservation(runtime); if (script.pathname.endsWith('/wordx/js/preserve-drawing.js')) installWordxDrawingPreservation(runtime); if (script.pathname.endsWith('/wordx/js/rtf.js')) installWordxRtf(runtime); if (script.pathname.endsWith('/wordx/js/spell.js')) installWordxProofing(runtime); if (script.pathname.endsWith('/wordx/js/markdown.js')) installWordxMarkdown(runtime); if (script.pathname.endsWith('/wordx/js/mdio.js')) installWordxMarkdownIo(runtime); },
      configureRuntime: (runtime) => { installWordxRibbonSpec(runtime); installWordxTemplates(runtime); installWordxLineBoxes(runtime); installWordxDrawTool(runtime); installWordxOutline(runtime); installWordxAutoCorrect(runtime); installWordxRulers(runtime); installWordxOmml(runtime); installWordxMailMerge(runtime); installWordxFind(runtime); },
      skipLegacyScripts: ['/common/search.js', '/wordx/js/ribbon.js', '/wordx/js/templates.js', '/wordx/js/portions.js', '/wordx/js/drawtool.js', '/wordx/js/outline.js', '/wordx/js/autocorrect.js', '/wordx/js/rulers.js', '/wordx/js/omml.js', '/wordx/js/mailmerge.js', '/wordx/js/find.js', '/wordx/js/spell.js', '/wordx/js/htmlio.js', '/wordx/js/markdown.js', '/wordx/js/mdio.js', '/wordx/js/rtf.js', '/wordx/js/docx-read.js', '/wordx/js/docx-write.js', '/wordx/js/fields.js', '/wordx/js/commands.js', '/wordx/js/ops.js', '/wordx/js/dmodel.js', '/wordx/js/drawing.js', '/wordx/js/preserve-drawing.js', '/wordx/js/preserve.js'],
      ...(allowedRemoteOrigins ? { allowedRemoteOrigins } : {}), ...(onOpenUrl ? { onOpenUrl } : {}),
      ...(draftStore ? { draftStore } : {}),
      ...(onDraftSaved ? { onDraftSaved } : {}),
      ...(onTitleChange ? { onTitleChange } : {}),
      ...(onDraftsAvailable ? { onDraftsAvailable } : {}),
      ...(onEngineReady ? { onEngineReady: (restoreDraft) => onEngineReady({ restoreDraft }) } : {}),
      ...(onOpenFile ? { onOpenFile } : {}), ...(onSaveFile ? { onSaveFile } : {}),
      ...(initialFile ? { initialFile } : {}), ...(initialDocument ? { initialDocument } : {}),
    });
  }, [mountEngine, engine, draftStore, onDraftSaved, onTitleChange, onDraftsAvailable, onEngineReady, onOpenFile, allowedRemoteOrigins, onOpenUrl, onSaveFile, initialFile, initialDocument]);
  const reactEngineShell = Boolean(engine && !mountEngine);
  return <VibeOfficeProvider {...(services ? { services } : {})}><EngineEditorFrame {...props} appName="Wordx" editor="wordx" {...(mount ? { mount } : {})} preserveHostChildrenOnUnmount={reactEngineShell} appLayers={(host) => <><WordxListBoxLayer host={host} /><WordxMenuContentLayer host={host} /><WordxTaskContentLayer host={host} />{props.appLayers?.(host)}</>}>{reactEngineShell ? <><WordxEngineShell />{props.children}</> : props.children}</EngineEditorFrame></VibeOfficeProvider>;
}
