import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Catalog = { reference: string; formula: string; title: string; description: string; action: string; index: number; count: number; canConvert: boolean };
type Runtime = { dlg?: { reactErrorCheckCatalog?(): Catalog | null; reactErrorCheckAction?(action: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string, ...args: unknown[]): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React Error Checking workflow backed by the existing worksheet scan and correction operations. */
export function ExcelxErrorCheckingDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactErrorCheckCatalog?.() || null);
    const off = runtime()?.bus?.on('excelx-error-checking-dialog', refresh); return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string, ...args: unknown[]) => runtime()?.t?.(message, ...args) || message;
  const run = (action: string) => runtime()?.dlg?.reactErrorCheckAction?.(action);
  return <OfficeDialog title={t('Error Checking')} className="excelx-error-checking-dialog" style={{ width: 'min(94vw, 500px)' }} onRequestClose={() => run('close')}>
    <div className="excelx-error-checking-dialog__body">
      <section><strong>{t('Error in cell {0}:', catalog.reference)}</strong><div className="excelx-error-checking-dialog__formula" translate="no">{catalog.formula}</div><strong>{catalog.title}</strong><p>{catalog.description}</p></section>
      <aside><button type="button" onClick={() => run('help')}>{t('Help on this error')}</button><button type="button" onClick={() => run('fix')}>{catalog.action}</button><button type="button" onClick={() => run('ignore')}>{t('Ignore Error')}</button><button type="button" onClick={() => run('edit')}>{t('Edit in Formula Bar')}</button></aside>
    </div>
    <div className="excelx-error-checking-dialog__actions"><span>{catalog.index + 1} / {catalog.count}</span><button type="button" onClick={() => run('previous')}>{t('&Previous').replace('&', '')}</button><button type="button" className="primary" onClick={() => run('next')}>{t('&Next').replace('&', '')}</button><button type="button" onClick={() => run('close')}>{t('Close')}</button></div>
  </OfficeDialog>;
}
