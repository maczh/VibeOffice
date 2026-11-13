import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface PivotChoice { value: number; label: string }
interface PivotState { row: number; col: number | ''; data: number; fn: string; page: number | ''; pageVal: string; where: 'new' | 'here' }
interface PivotCatalog { reference: string; columns: PivotChoice[]; functions: { value: string; label: string }[]; initial: PivotState }
interface PivotRuntime {
  dlg?: { reactPivotCatalog?(): PivotCatalog | null; reactPivotApply?(reference: string, state: PivotState): boolean; reactPivotClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): PivotRuntime | undefined { return (window as Window & { L?: PivotRuntime }).L; }

export function ExcelxPivotDialog() {
  const [catalog, setCatalog] = useState<PivotCatalog | null>(null);
  const [reference, setReference] = useState('');
  const [state, setState] = useState<PivotState>({ row: 0, col: '', data: 0, fn: 'sum', page: '', pageVal: '', where: 'new' });
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactPivotCatalog?.() || null; setCatalog(next); if (next) { setReference(next.reference); setState(next.initial); } };
    const off = runtime()?.bus?.on('excelx-pivot-dialog', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactPivotCatalog?.()) runtime()?.dlg?.reactPivotClose?.(); };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactPivotClose?.();
  const apply = () => runtime()?.dlg?.reactPivotApply?.(reference, state);
  const update = (patch: Partial<PivotState>) => setState((current) => ({ ...current, ...patch }));
  const columnSelect = (key: 'row' | 'col' | 'data' | 'page', optional = false) => <label>{translate(key === 'row' ? '&Row field:' : key === 'col' ? '&Column field:' : key === 'data' ? '&Data field:' : '&Page field:')}<select value={state[key]} onChange={(event) => update({ [key]: event.currentTarget.value === '' ? '' : Number(event.currentTarget.value) })}>{optional && <option value="">{translate('(none)')}</option>}{catalog.columns.map((column) => <option key={column.value} value={column.value}>{translate(column.label)}</option>)}</select></label>;
  return <OfficeDialog title={translate('PivotTable and PivotChart Wizard')} className="excelx-pivot-dialog" style={{ width: 'min(94vw, 520px)' }} onRequestClose={close}>
    <div className="excelx-pivot-dialog__body">
      <label>{translate('&Range:')}<input type="text" value={reference} onChange={(event) => { const value = event.currentTarget.value; setReference(value); }} /></label>
      <fieldset><legend>{translate('Layout')}</legend>{columnSelect('row')}{columnSelect('col', true)}<div className="excelx-pivot-dialog__data-row">{columnSelect('data')}<label>{translate('Summarize by:')}<select value={state.fn} onChange={(event) => update({ fn: event.currentTarget.value })}>{catalog.functions.map((fn) => <option key={fn.value} value={fn.value}>{translate(fn.label)}</option>)}</select></label></div>{columnSelect('page', true)}</fieldset>
      <fieldset><legend>{translate('Where do you want to put the PivotTable report?')}</legend>
        <label><input type="radio" name="excelx-pivot-destination" checked={state.where === 'new'} onChange={() => update({ where: 'new' })} />{translate('&New worksheet').replace('&', '')}</label>
        <label><input type="radio" name="excelx-pivot-destination" checked={state.where === 'here'} onChange={() => update({ where: 'here' })} />{translate('&Existing worksheet (at the active cell)').replace('&', '')}</label>
      </fieldset>
      <p>{translate('Excelx builds the report as a formatted summary of values; rebuild it after the source data changes.')}</p>
    </div>
    <div className="excelx-pivot-dialog__actions"><button type="button" className="primary" onClick={apply}>{translate('&Finish').replace('&', '')}</button><button type="button" onClick={close}>{translate('Cancel')}</button></div>
  </OfficeDialog>;
}
