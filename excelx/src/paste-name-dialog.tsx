import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Catalog = { names: Array<{ key: string; name: string }>; pick: string };
type Runtime = { dlg?: { reactPasteNameCatalog?(): Catalog | null; reactPasteNameAction?(action: string, value?: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React dialog for inserting a workbook name or pasting the complete name list. */
export function ExcelxPasteNameDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactPasteNameCatalog?.() || null);
    const off = runtime()?.bus?.on('excelx-paste-name', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactPasteNameCatalog?.()) runtime()?.dlg?.reactPasteNameAction?.('close'); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = runtime()?.dlg?.reactPasteNameAction;
  const insert = () => action?.('insert');
  return <OfficeDialog title={t('Paste Name')} className="excelx-paste-name-dialog" style={{ width: 'min(92vw, 360px)' }} onRequestClose={() => action?.('close')}>
    <label className="excelx-paste-name-dialog__label">{t('Paste &name').replace('&', '')}
      <select autoFocus size={Math.min(9, Math.max(3, catalog.names.length))} aria-label={t('Paste &name').replace('&', '')} value={catalog.pick} onChange={(event) => action?.('select', event.currentTarget.value)} onDoubleClick={insert}>
        {catalog.names.map((item) => <option key={item.key} value={item.key} translate="no">{item.name}</option>)}
      </select>
    </label>
    <div className="excelx-paste-name-dialog__actions"><button type="button" onClick={() => action?.('paste-list')}>{t('Paste &List').replace('&', '')}</button><button type="button" className="primary" onClick={insert}>{t('OK')}</button><button type="button" onClick={() => action?.('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
