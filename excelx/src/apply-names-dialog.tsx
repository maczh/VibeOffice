import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Catalog = { names: string[]; selected: string[] };
type Runtime = { dlg?: { reactApplyNamesCatalog?(): Catalog | null; reactApplyNamesAction?(action: string, name?: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React checkbox list for replacing matching cell references with defined names. */
export function ExcelxApplyNamesDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactApplyNamesCatalog?.() || null);
    const off = runtime()?.bus?.on('excelx-apply-names', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactApplyNamesCatalog?.()) runtime()?.dlg?.reactApplyNamesAction?.('close'); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = runtime()?.dlg?.reactApplyNamesAction;
  return <OfficeDialog title={t('Apply Names')} className="excelx-apply-names-dialog" style={{ width: 'min(92vw, 380px)' }} onRequestClose={() => action?.('close')}>
    <fieldset className="excelx-apply-names-dialog__list"><legend>{t('Apply &names:').replace('&', '')}</legend>
      {catalog.names.map((name) => <label key={name}><input type="checkbox" checked={catalog.selected.includes(name)} onChange={() => action?.('toggle', name)} /><span translate="no">{name}</span></label>)}
    </fieldset>
    <div className="excelx-apply-names-dialog__actions"><button type="button" className="primary" onClick={() => action?.('commit')}>{t('OK')}</button><button type="button" onClick={() => action?.('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
