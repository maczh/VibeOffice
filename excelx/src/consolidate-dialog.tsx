import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface ConsolidateChoice { value: string; label: string }
interface ConsolidateState { refs: string[]; fn: string; top: boolean; left: boolean }
interface ConsolidateCatalog extends ConsolidateState { functions: ConsolidateChoice[] }
interface ConsolidateRuntime {
  dlg?: { reactConsolidateCatalog?(): ConsolidateCatalog | null; reactConsolidateAdd?(reference: string): boolean; reactConsolidateDelete?(index: number): boolean; reactConsolidateApply?(state: ConsolidateState): boolean; reactConsolidateClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): ConsolidateRuntime | undefined { return (window as Window & { L?: ConsolidateRuntime }).L; }

export function ExcelxConsolidateDialog() {
  const [catalog, setCatalog] = useState<ConsolidateCatalog | null>(null);
  const [state, setState] = useState<ConsolidateState>({ refs: [], fn: 'SUM', top: false, left: false });
  const [reference, setReference] = useState('');
  const [selected, setSelected] = useState(-1);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactConsolidateCatalog?.() || null; setCatalog(next); if (next) { setState({ refs: next.refs, fn: next.fn, top: next.top, left: next.left }); setSelected(-1); } };
    const off = runtime()?.bus?.on('excelx-consolidate-dialog', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactConsolidateCatalog?.()) runtime()?.dlg?.reactConsolidateClose?.(); };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactConsolidateClose?.();
  const apply = () => runtime()?.dlg?.reactConsolidateApply?.(state);
  const add = () => { if (runtime()?.dlg?.reactConsolidateAdd?.(reference)) setReference(''); };
  const remove = () => { if (runtime()?.dlg?.reactConsolidateDelete?.(selected)) setSelected(-1); };
  const update = (patch: Partial<ConsolidateState>) => setState((current) => ({ ...current, ...patch }));
  return <OfficeDialog title={translate('Consolidate')} className="excelx-consolidate-dialog" style={{ width: 'min(92vw, 480px)' }} onRequestClose={close}>
    <div className="excelx-consolidate-dialog__body">
      <label>{translate('&Function:')}<select value={state.fn} onChange={(event) => update({ fn: event.currentTarget.value })}>{catalog.functions.map((fn) => <option key={fn.value} value={fn.value}>{translate(fn.label)}</option>)}</select></label>
      <label>{translate('&Reference:')}<input type="text" value={reference} onChange={(event) => { const value = event.currentTarget.value; setReference(value); }} /></label>
      <div className="excelx-consolidate-dialog__reference-actions"><button type="button" onClick={add}>{translate('&Add').replace('&', '')}</button><button type="button" onClick={remove} disabled={selected < 0}>{translate('&Delete').replace('&', '')}</button></div>
      <label>{translate('All r&eferences:')}<select size={5} aria-label={translate('All r&eferences:').replace('&', '')} value={selected < 0 ? '' : String(selected)} onChange={(event) => setSelected(Number(event.currentTarget.value))}>{state.refs.map((item, index) => <option key={`${item}-${index}`} value={index}>{item}</option>)}</select></label>
      <fieldset><legend>{translate('Use labels in')}</legend>
        <label><input type="checkbox" checked={state.top} onChange={(event) => update({ top: event.currentTarget.checked })} />{translate('&Top row').replace('&', '')}</label>
        <label><input type="checkbox" checked={state.left} onChange={(event) => update({ left: event.currentTarget.checked })} />{translate('&Left column').replace('&', '')}</label>
      </fieldset>
    </div>
    <div className="excelx-consolidate-dialog__actions"><button type="button" className="primary" onClick={apply}>{translate('OK')}</button><button type="button" onClick={close}>{translate('Close')}</button></div>
  </OfficeDialog>;
}
