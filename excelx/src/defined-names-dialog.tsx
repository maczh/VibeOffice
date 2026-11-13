import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type NameItem = { key: string; name: string; ref: string; comment: string };
type Catalog = { names: NameItem[]; name: string; ref: string; comment: string; error: string };
type Runtime = { dlg?: { reactDefinedNamesCatalog?(): Catalog | null; reactDefinedNamesAction?(action: string, value?: unknown): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React owner for workbook defined-name browsing and editing. */
export function ExcelxDefinedNamesDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactDefinedNamesCatalog?.() || null);
    const off = runtime()?.bus?.on('excelx-defined-names', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactDefinedNamesCatalog?.()) runtime()?.dlg?.reactDefinedNamesAction?.('close'); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = runtime()?.dlg?.reactDefinedNamesAction;
  const change = (key: 'name' | 'ref' | 'comment', value: string) => action?.('change', { [key]: value });
  return <OfficeDialog title={t('Define Name')} className="excelx-defined-names-dialog" style={{ width: 'min(94vw, 520px)' }} onRequestClose={() => action?.('close')}>
    <div className="excelx-defined-names-dialog__body">
      <label>{t('Names in &workbook:').replace('&', '')}<input autoFocus value={catalog.name} onChange={(event) => change('name', event.currentTarget.value)} /></label>
      <select aria-label={t('Names in &workbook:').replace('&', '')} size={Math.min(7, Math.max(2, catalog.names.length))} value="" onChange={(event) => action?.('select', event.currentTarget.value)}>
        {catalog.names.map((item) => <option key={item.key} value={item.key} translate="no">{item.name}</option>)}
      </select>
      <label>{t('Co&mment:').replace('&', '')}<input value={catalog.comment} onChange={(event) => change('comment', event.currentTarget.value)} /></label>
      <label>{t('&Refers to:').replace('&', '')}<input value={catalog.ref} onChange={(event) => change('ref', event.currentTarget.value)} /></label>
      {catalog.error && <p role="alert">{t(catalog.error)}</p>}
    </div>
    <div className="excelx-defined-names-dialog__actions"><button type="button" onClick={() => action?.('add')}>{t('&Add').replace('&', '')}</button><button type="button" onClick={() => action?.('delete')}>{t('&Delete').replace('&', '')}</button><button type="button" className="primary" onClick={() => action?.('commit')}>{t('OK')}</button><button type="button" onClick={() => action?.('close')}>{t('Close')}</button></div>
  </OfficeDialog>;
}
