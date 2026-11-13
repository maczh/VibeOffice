import { createDirectEngineBridge, EngineEditorFrame } from '@vibeoffice/core';
import { VibeOfficeProvider } from '@vibeoffice/core';
import type { VibeOfficeServiceOverrides } from '@vibeoffice/core';
import type { DraftStore, EngineRuntimeOptions, LocalDocument, LocalDraft, SaveFileOperation } from '@vibeoffice/core';
import { useMemo, type ComponentProps } from 'react';
import { ExcelxFormulaBarLayer } from './formula-bar.js';
import { ExcelxFileDropLayer } from './file-drop.js';
import { ExcelxGridInteractionLayer } from './grid-interactions.js';
import { ExcelxInsertSheetDialog } from './insert-sheet-dialog.js';
import { ExcelxSheetTabsLayer } from './sheet-tabs.js';
import { ExcelxSizeDialog } from './size-dialog.js';
import { ExcelxShiftCellsDialog } from './shift-cells-dialog.js';
import { ExcelxStatusBarInteractions } from './status-bar-interactions.js';
import { ExcelxZoomDialog } from './zoom-dialog.js';
import { ExcelxUnhideSheetDialog } from './unhide-sheet-dialog.js';
import { ExcelxTabColorDialog } from './tab-color-dialog.js';
import { ExcelxMoveCopySheetDialog } from './move-copy-sheet-dialog.js';
import { ExcelxGoToDialog } from './go-to-dialog.js';
import { ExcelxGoToSpecialDialog } from './go-to-special-dialog.js';
import { ExcelxPasteSpecialDialog } from './paste-special-dialog.js';
import { ExcelxFillAcrossDialog } from './fill-across-dialog.js';
import { ExcelxSeriesDialog } from './series-dialog.js';
import { ExcelxSymbolDialog } from './symbol-dialog.js';
import { ExcelxPropertiesDialog } from './properties-dialog.js';
import { ExcelxSortDialog } from './sort-dialog.js';
import { ExcelxCreateListDialog } from './create-list-dialog.js';
import { ExcelxAdvancedFilterDialog } from './advanced-filter-dialog.js';
import { ExcelxSubtotalsDialog } from './subtotals-dialog.js';
import { ExcelxOutlineSettingsDialog } from './outline-settings-dialog.js';
import { ExcelxDataTableDialog } from './data-table-dialog.js';
import { ExcelxGoalSeekDialog } from './goal-seek-dialog.js';
import { ExcelxConsolidateDialog } from './consolidate-dialog.js';
import { ExcelxPivotDialog } from './pivot-dialog.js';
import { ExcelxDataFormDialog } from './data-form-dialog.js';
import { ExcelxValidationDialog } from './validation-dialog.js';
import { ExcelxTextToColumnsDialog } from './text-to-columns-dialog.js';
import { ExcelxAutoFormatDialog } from './auto-format-dialog.js';
import { ExcelxStyleDialog } from './style-dialog.js';
import { ExcelxConditionalFormatStyleDialog } from './conditional-format-style-dialog.js';
import { ExcelxConditionalFormatDialog } from './conditional-format-dialog.js';

export type ExcelxEditorProps = Omit<ComponentProps<typeof EngineEditorFrame>, 'appName' | 'editor'> & { services?: VibeOfficeServiceOverrides; engine?: EngineRuntimeOptions; draftStore?: DraftStore<Blob>; onDraftSaved?: (draft: LocalDraft<Blob>) => void; onEngineReady?: (restore: (draft: LocalDraft<Blob>) => void) => void; onDraftsAvailable?: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void; onOpenFile?: (signal: AbortSignal) => Promise<LocalDocument | null>; allowedRemoteOrigins?: readonly string[]; onOpenUrl?: (url: string, signal: AbortSignal) => Promise<LocalDocument | null>; onSaveFile?: (name: string, blob: Blob, current?: LocalDocument, signal?: AbortSignal, operation?: SaveFileOperation) => Promise<LocalDocument | void | null>; initialFile?: File; initialDocument?: LocalDocument };

export function ExcelxEditor({ services, engine, draftStore, onDraftSaved, onEngineReady, onDraftsAvailable, onOpenFile, allowedRemoteOrigins, onOpenUrl, onSaveFile, initialFile, initialDocument, ...props }: ExcelxEditorProps) {
  const mount = useMemo(() => engine ? createDirectEngineBridge({ editor: 'excelx', engine, title: 'Excelx editor', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm', '.xltx', '.xltm'], 'text/plain': ['.csv', '.txt', '.tsv', '.prn', '.xml'] }, ...(draftStore ? { draftStore } : {}), ...(onDraftSaved ? { onDraftSaved } : {}), ...(onEngineReady ? { onEngineReady } : {}), ...(onDraftsAvailable ? { onDraftsAvailable } : {}), ...(onOpenFile ? { onOpenFile } : {}), ...(allowedRemoteOrigins ? { allowedRemoteOrigins } : {}), ...(onOpenUrl ? { onOpenUrl } : {}), ...(onSaveFile ? { onSaveFile } : {}), ...(initialFile ? { initialFile } : {}), ...(initialDocument ? { initialDocument } : {}) }) : undefined, [engine, draftStore, onDraftSaved, onEngineReady, onDraftsAvailable, onOpenFile, allowedRemoteOrigins, onOpenUrl, onSaveFile, initialFile, initialDocument]);
  return <VibeOfficeProvider {...(services ? { services } : {})}><EngineEditorFrame appName="Excelx" editor="excelx" {...(mount ? { mount } : {})} {...props} appLayers={(host) => <><ExcelxFormulaBarLayer host={host} /><ExcelxFileDropLayer host={host} /><ExcelxGridInteractionLayer host={host} /><ExcelxSheetTabsLayer host={host} /><ExcelxStatusBarInteractions host={host} /><ExcelxZoomDialog /><ExcelxSizeDialog /><ExcelxShiftCellsDialog /><ExcelxInsertSheetDialog /><ExcelxUnhideSheetDialog /><ExcelxTabColorDialog /><ExcelxMoveCopySheetDialog /><ExcelxGoToDialog /><ExcelxGoToSpecialDialog /><ExcelxPasteSpecialDialog /><ExcelxFillAcrossDialog /><ExcelxSeriesDialog /><ExcelxSymbolDialog /><ExcelxPropertiesDialog /><ExcelxSortDialog /><ExcelxCreateListDialog /><ExcelxAdvancedFilterDialog /><ExcelxSubtotalsDialog /><ExcelxOutlineSettingsDialog /><ExcelxDataTableDialog /><ExcelxGoalSeekDialog /><ExcelxConsolidateDialog /><ExcelxPivotDialog /><ExcelxDataFormDialog /><ExcelxValidationDialog /><ExcelxTextToColumnsDialog /><ExcelxAutoFormatDialog /><ExcelxStyleDialog /><ExcelxConditionalFormatStyleDialog /><ExcelxConditionalFormatDialog />{props.appLayers?.(host)}</>} /></VibeOfficeProvider>;
}
