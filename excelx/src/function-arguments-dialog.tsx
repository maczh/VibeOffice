import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import './function-arguments-dialog.css';

type Argument = { name: string; value: string; optional: boolean; help: string; preview: string };
type Catalog = { name: string; description: string; help: string; repeating: boolean; args: Argument[]; call: string; result: string; total: string };
type Runtime = { dlg?: { reactFunctionArgsCatalog?(): Catalog | null; reactFunctionArgsAction?(action: string, index?: number, value?: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React function-argument editor; formula parsing, evaluation and commit remain in the spreadsheet engine. */
export function ExcelxFunctionArgumentsDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactFunctionArgsCatalog?.() || null);
    const off = runtime()?.bus?.on('excelx-function-arguments-dialog', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactFunctionArgsCatalog?.()) runtime()?.dlg?.reactFunctionArgsAction?.('close'); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = runtime()?.dlg?.reactFunctionArgsAction;
  return <OfficeDialog title={t('Function Arguments')} className="excelx-function-arguments-dialog" style={{ width: 'min(94vw, 560px)' }} onRequestClose={() => action?.('close')}>
    <strong translate="no">{catalog.name}</strong>
    <div className="excelx-function-arguments-dialog__rows">{catalog.args.map((argument, index) => <label className={argument.optional ? 'opt' : 'req'} key={`${catalog.name}-${index}`}>
      <span><span translate="no">{argument.name.replace(/[\[\]]/g, '').replace(/^./, (character) => character.toUpperCase())}</span>{argument.optional && <small> ({t('Optional')})</small>}</span>
      <input value={argument.value} aria-label={argument.name} onFocus={() => action?.('focus', index)} onChange={(event) => action?.('set', index, event.currentTarget.value)} />
      <output translate="no">{argument.preview}</output>
    </label>)}</div>
    <div className="excelx-function-arguments-dialog__result" translate="no">{catalog.result}</div>
    <p className="excelx-function-arguments-dialog__description">{t(catalog.description)}</p>
    <p className="excelx-function-arguments-dialog__help" aria-live="polite">{catalog.help}</p>
    <p className="excelx-function-arguments-dialog__total" aria-live="polite">{catalog.total}</p>
    <div className="excelx-function-arguments-dialog__actions"><button type="button" className="primary" onClick={() => action?.('apply')}>{t('OK')}</button><button type="button" onClick={() => action?.('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
