import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface Choice { value: number; label: string }
interface SubtotalState { at: number; fn: number; add: number[]; replace: boolean; pageBreak: boolean; below: boolean }
interface SubtotalsCatalog { columns: Choice[]; functions: Choice[]; initial: SubtotalState }
interface SubtotalsRuntime {
  dlg?: { reactSubtotalsCatalog?(): SubtotalsCatalog | null; reactSubtotalsApply?(state: SubtotalState): boolean; reactSubtotalsRemove?(): boolean; reactSubtotalsClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): SubtotalsRuntime | undefined { return (window as Window & { L?: SubtotalsRuntime }).L; }

export function ExcelxSubtotalsDialog() {
  const [catalog, setCatalog] = useState<SubtotalsCatalog | null>(null);
  const [state, setState] = useState<SubtotalState>({ at: 0, fn: 9, add: [], replace: true, pageBreak: false, below: true });
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactSubtotalsCatalog?.() || null; setCatalog(next); if (next) setState(next.initial); };
    const off = runtime()?.bus?.on('excelx-subtotals-dialog', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactSubtotalsCatalog?.()) runtime()?.dlg?.reactSubtotalsClose?.(); };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactSubtotalsClose?.();
  const apply = () => runtime()?.dlg?.reactSubtotalsApply?.(state);
  const remove = () => runtime()?.dlg?.reactSubtotalsRemove?.();
  const update = (patch: Partial<SubtotalState>) => setState((current) => ({ ...current, ...patch }));
  const toggleColumn = (column: number, checked: boolean) => setState((current) => ({ ...current, add: checked ? [...new Set([...current.add, column])] : current.add.filter((value) => value !== column) }));
  return <OfficeDialog title={translate('Subtotal')} className="excelx-subtotals-dialog" style={{ width: 'min(92vw, 420px)' }} onRequestClose={close}>
    <div className="excelx-subtotals-dialog__body">
      <label>{translate('&At each change in:')}<select value={state.at} onChange={(event) => update({ at: Number(event.currentTarget.value) })}>{catalog.columns.map((column) => <option key={column.value} value={column.value}>{translate(column.label)}</option>)}</select></label>
      <label>{translate('&Use function:')}<select value={state.fn} onChange={(event) => update({ fn: Number(event.currentTarget.value) })}>{catalog.functions.map((fn) => <option key={fn.value} value={fn.value}>{translate(fn.label)}</option>)}</select></label>
      <fieldset><legend>{translate('Add subtotal to:')}</legend>{catalog.columns.map((column) => <label key={column.value}><input type="checkbox" checked={state.add.includes(column.value)} onChange={(event) => toggleColumn(column.value, event.currentTarget.checked)} />{translate(column.label)}</label>)}</fieldset>
      <label><input type="checkbox" checked={state.replace} onChange={(event) => update({ replace: event.currentTarget.checked })} />{translate('Replace &current subtotals').replace('&', '')}</label>
      <label><input type="checkbox" checked={state.pageBreak} onChange={(event) => update({ pageBreak: event.currentTarget.checked })} />{translate('&Page break between groups').replace('&', '')}</label>
      <label><input type="checkbox" checked={state.below} onChange={(event) => update({ below: event.currentTarget.checked })} />{translate('&Summary below data').replace('&', '')}</label>
    </div>
    <div className="excelx-subtotals-dialog__actions"><button type="button" onClick={remove}>{translate('&Remove All').replace('&', '')}</button><span /><button type="button" className="primary" onClick={apply}>{translate('OK')}</button><button type="button" onClick={close}>{translate('Cancel')}</button></div>
  </OfficeDialog>;
}
