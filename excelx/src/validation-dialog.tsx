import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState } from 'react';
import { useExcelxRuntime } from './editor-host';

interface ValidationState { type: string; op: string; allowBlank: boolean; showDrop: boolean; showInput: boolean; showError: boolean; errorStyle: string; value1: string; value2: string; promptTitle: string; prompt: string; errorTitle: string; error: string; ime: string; applyAll: boolean }
type Choice = [string, string];
interface ValidationCatalog { state: ValidationState; types: Choice[]; operators: Choice[]; errorStyles: Choice[]; imeModes: Choice[] }
interface ValidationRuntime {
  dlg?: { reactValidationCatalog?(): ValidationCatalog | null; reactValidationApply?(state: ValidationState): boolean; reactValidationClear?(): void; reactValidationClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
export function ExcelxValidationDialog() {
  const runtime = useExcelxRuntime<ValidationRuntime>();
  const [catalog, setCatalog] = useState<ValidationCatalog | null>(null);
  const id = useId().replaceAll(':', '');
  const [state, setState] = useState<ValidationState>({ type: 'none', op: 'between', allowBlank: true, showDrop: true, showInput: true, showError: true, errorStyle: 'stop', value1: '', value2: '', promptTitle: '', prompt: '', errorTitle: '', error: '', ime: '', applyAll: false });
  const [tab, setTab] = useState(0);
  useEffect(() => {
    const refresh = () => { const next = runtime?.dlg?.reactValidationCatalog?.() || null; setCatalog(next); if (next) { setState(next.state); setTab(0); } };
    const off = runtime?.bus?.on('excelx-validation-dialog', refresh);
    return () => { off?.(); if (runtime?.dlg?.reactValidationCatalog?.()) runtime?.dlg?.reactValidationClose?.(); };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime?.t?.(message) || message;
  const close = () => runtime?.dlg?.reactValidationClose?.();
  const apply = () => runtime?.dlg?.reactValidationApply?.(state);
  const clear = () => runtime?.dlg?.reactValidationClear?.();
  const update = (patch: Partial<ValidationState>) => setState((current) => ({ ...current, ...patch }));
  const select = (label: string, value: string, choices: Choice[], onChange: (value: string) => void) => <label>{translate(label)}<select value={value} onChange={(event) => onChange(event.currentTarget.value)}>{choices.map(([key, name]) => <option key={key} value={key}>{translate(name)}</option>)}</select></label>;
  const ranged = ['whole', 'decimal', 'date', 'time', 'textLength'].includes(state.type);
  const two = ranged && ['between', 'notBetween'].includes(state.op);
  const firstLabel = state.type === 'list' ? '&Source:' : state.type === 'custom' ? 'F&ormula:' : two ? (state.type === 'date' ? '&Start date:' : state.type === 'time' ? '&Start time:' : '&Minimum:') : state.type === 'date' ? '&Date:' : state.type === 'time' ? '&Time:' : state.type === 'textLength' ? '&Length:' : '&Value:';
  const secondLabel = state.type === 'date' ? '&End date:' : state.type === 'time' ? '&End time:' : 'Ma&ximum:';
  const tabs = ['Settings', 'Input Message', 'Error Alert', 'IME Mode'];
  const tabId = (index: number) => `${id}-validation-tab-${index}`;
  const panelId = `${id}-validation-panel`;
  return <OfficeDialog title={translate('Data Validation')} className="excelx-validation-dialog" style={{ width: 'min(94vw, 580px)' }} onRequestClose={close}>
    <div className="excelx-validation-dialog__tabs" role="tablist" aria-label={translate('Data Validation')}>{tabs.map((name, index) => <button type="button" role="tab" id={tabId(index)} aria-controls={panelId} aria-selected={tab === index} tabIndex={tab === index ? 0 : -1} key={name} onClick={() => setTab(index)} onKeyDown={(event) => {
      const next = event.key === 'ArrowRight' ? (index + 1) % tabs.length
        : event.key === 'ArrowLeft' ? (index - 1 + tabs.length) % tabs.length
          : event.key === 'Home' ? 0
            : event.key === 'End' ? tabs.length - 1 : -1;
      if (next < 0) return;
      event.preventDefault();
      setTab(next);
      event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
    }}>{translate(name)}</button>)}</div>
    <div id={panelId} aria-labelledby={tabId(tab)} role="tabpanel" className="excelx-validation-dialog__body">
      {tab === 0 && <div className="excelx-validation-dialog__settings"><fieldset><legend>{translate('Validation criteria')}</legend>
        {select('&Allow:', state.type, catalog.types, (type) => update({ type }))}
        {ranged && select('&Data:', state.op, catalog.operators, (op) => update({ op }))}
        {state.type !== 'none' && <label>{translate(firstLabel)}<input type="text" value={state.value1} onChange={(event) => update({ value1: event.currentTarget.value })} /></label>}
        {two && <label>{translate(secondLabel)}<input type="text" value={state.value2} onChange={(event) => update({ value2: event.currentTarget.value })} /></label>}
        <label><input type="checkbox" checked={state.allowBlank} disabled={state.type === 'none'} onChange={(event) => update({ allowBlank: event.currentTarget.checked })} />{translate('Ignore &blank').replace('&', '')}</label>
        <label><input type="checkbox" checked={state.showDrop} disabled={state.type !== 'list'} onChange={(event) => update({ showDrop: event.currentTarget.checked })} />{translate('&In-cell dropdown').replace('&', '')}</label>
      </fieldset><label><input type="checkbox" checked={state.applyAll} onChange={(event) => update({ applyAll: event.currentTarget.checked })} />{translate('Apply these changes to all other cells with the same settings')}</label></div>}
      {tab === 1 && <div className="excelx-validation-dialog__message"><label><input type="checkbox" checked={state.showInput} onChange={(event) => update({ showInput: event.currentTarget.checked })} />{translate('&Show input message when cell is selected').replace('&', '')}</label><p>{translate('When cell is selected, show this input message:')}</p><label>{translate('&Title:')}<input type="text" value={state.promptTitle} onChange={(event) => update({ promptTitle: event.currentTarget.value })} /></label><label>{translate('&Input message:')}<textarea rows={4} value={state.prompt} onChange={(event) => update({ prompt: event.currentTarget.value })} /></label></div>}
      {tab === 2 && <div className="excelx-validation-dialog__error"><label><input type="checkbox" checked={state.showError} onChange={(event) => update({ showError: event.currentTarget.checked })} />{translate('&Show error alert after invalid data is entered').replace('&', '')}</label><p>{translate('When user enters invalid data, show this error alert:')}</p>{select('St&yle:', state.errorStyle, catalog.errorStyles, (errorStyle) => update({ errorStyle }))}<label>{translate('&Title:')}<input type="text" value={state.errorTitle} onChange={(event) => update({ errorTitle: event.currentTarget.value })} /></label><label>{translate('&Error message:')}<textarea rows={4} value={state.error} onChange={(event) => update({ error: event.currentTarget.value })} /></label></div>}
      {tab === 3 && <div className="excelx-validation-dialog__ime">{select('&Mode:', state.ime, catalog.imeModes, (ime) => update({ ime }))}</div>}
    </div>
    <div className="excelx-validation-dialog__actions"><button type="button" onClick={clear}>{translate('&Clear All').replace('&', '')}</button><span /><button type="button" className="primary" onClick={apply}>{translate('OK')}</button><button type="button" onClick={close}>{translate('Cancel')}</button></div>
  </OfficeDialog>;
}
