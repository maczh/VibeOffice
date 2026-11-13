import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Catalog = { top: boolean; left: boolean; bottom: boolean; right: boolean };
type Runtime = { dlg?: { reactCreateNamesCatalog?(): Catalog | null; reactCreateNamesAction?(action: string, value?: Partial<Catalog>): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React choices for deriving workbook names from the selected range's edge labels. */
export function ExcelxCreateNamesDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactCreateNamesCatalog?.() || null);
    const off = runtime()?.bus?.on('excelx-create-names', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactCreateNamesCatalog?.()) runtime()?.dlg?.reactCreateNamesAction?.('close'); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const change = (key: keyof Catalog, value: boolean) => runtime()?.dlg?.reactCreateNamesAction?.('change', { [key]: value });
  const check = (key: keyof Catalog, label: string) => <label key={key}><input type="checkbox" checked={catalog[key]} onChange={(event) => change(key, event.currentTarget.checked)} />{t(label).replace('&', '')}</label>;
  return <OfficeDialog title={t('Create Names')} className="excelx-create-names-dialog" style={{ width: 'min(92vw, 340px)' }} onRequestClose={() => runtime()?.dlg?.reactCreateNamesAction?.('close')}>
    <fieldset className="excelx-create-names-dialog__choices"><legend>{t('Create names in')}</legend>{check('top', '&Top row')}{check('left', '&Left column')}{check('bottom', '&Bottom row')}{check('right', '&Right column')}</fieldset>
    <div className="excelx-create-names-dialog__actions"><button type="button" className="primary" onClick={() => runtime()?.dlg?.reactCreateNamesAction?.('commit')}>{t('OK')}</button><button type="button" onClick={() => runtime()?.dlg?.reactCreateNamesAction?.('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
