import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Catalog = { reference: string; expression: string; start: number; end: number; complete: boolean };
type Runtime = { dlg?: { reactEvaluateCatalog?(): Catalog | null; reactEvaluateAction?(action: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React formula-evaluation walkthrough; AST reduction and calculation use the spreadsheet engine. */
export function ExcelxEvaluateFormulaDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactEvaluateCatalog?.() || null);
    const off = runtime()?.bus?.on('excelx-evaluate-formula-dialog', refresh); return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = runtime()?.dlg?.reactEvaluateAction;
  const before = catalog.start >= 0 ? catalog.expression.slice(0, catalog.start) : catalog.expression;
  const part = catalog.start >= 0 ? catalog.expression.slice(catalog.start, catalog.end) : '';
  const after = catalog.end >= 0 ? catalog.expression.slice(catalog.end) : '';
  return <OfficeDialog title={t('Evaluate Formula')} className="excelx-evaluate-formula-dialog" style={{ width: 'min(94vw, 560px)' }} onRequestClose={() => action?.('close')}>
    <div className="excelx-evaluate-formula-dialog__body"><div><span>{t('Reference:')}</span><strong translate="no">{catalog.reference}</strong></div><div className="excelx-evaluate-formula-dialog__expression"><span>{t('Evaluation:')}</span><code translate="no">{before}{part && <u>{part}</u>}{after}</code></div><p>{t('To show the result of the underlined expression, click Evaluate. The most recent result appears italicized.')}</p></div>
    <div className="excelx-evaluate-formula-dialog__actions"><button type="button" className="primary" onClick={() => action?.(catalog.complete ? 'restart' : 'evaluate')}>{t(catalog.complete ? 'Restart' : '&Evaluate').replace('&', '')}</button><button type="button" onClick={() => action?.('close')}>{t('Close')}</button></div>
  </OfficeDialog>;
}
