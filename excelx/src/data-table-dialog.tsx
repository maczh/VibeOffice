import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface DataTableState { row: string; column: string }
interface DataTableRuntime {
  dlg?: { reactDataTableCatalog?(): DataTableState | null; reactDataTableApply?(state: DataTableState): boolean; reactDataTableClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): DataTableRuntime | undefined { return (window as Window & { L?: DataTableRuntime }).L; }

export function ExcelxDataTableDialog() {
  const [catalog, setCatalog] = useState<DataTableState | null>(null);
  const [state, setState] = useState<DataTableState>({ row: '', column: '' });
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactDataTableCatalog?.() || null; setCatalog(next); if (next) setState(next); };
    const off = runtime()?.bus?.on('excelx-data-table-dialog', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactDataTableCatalog?.()) runtime()?.dlg?.reactDataTableClose?.(); };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactDataTableClose?.();
  const apply = () => runtime()?.dlg?.reactDataTableApply?.(state);
  return <OfficeDialog title={translate('Table')} className="excelx-data-table-dialog" style={{ width: 'min(92vw, 380px)' }} onRequestClose={close}>
    <div className="excelx-data-table-dialog__body">
      <label>{translate('&Row input cell:')}<input type="text" aria-label={translate('&Row input cell:').replace('&', '')} value={state.row} onChange={(event) => { const row = event.currentTarget.value; setState((current) => ({ ...current, row })); }} /></label>
      <label>{translate('&Column input cell:')}<input type="text" aria-label={translate('&Column input cell:').replace('&', '')} value={state.column} onChange={(event) => { const column = event.currentTarget.value; setState((current) => ({ ...current, column })); }} /></label>
    </div>
    <div className="excelx-data-table-dialog__actions"><button type="button" className="primary" onClick={apply}>{translate('OK')}</button><button type="button" onClick={close}>{translate('Cancel')}</button></div>
  </OfficeDialog>;
}
