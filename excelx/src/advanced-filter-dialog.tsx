import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface AdvancedFilterState { list: string; criteria: string; destination: string; copy: boolean; unique: boolean }
interface AdvancedFilterRuntime {
  dlg?: { reactAdvancedFilterCatalog?(): AdvancedFilterState | null; reactAdvancedFilterApply?(state: AdvancedFilterState): boolean; reactAdvancedFilterClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): AdvancedFilterRuntime | undefined { return (window as Window & { L?: AdvancedFilterRuntime }).L; }

export function ExcelxAdvancedFilterDialog() {
  const [catalog, setCatalog] = useState<AdvancedFilterState | null>(null);
  const [state, setState] = useState<AdvancedFilterState>({ list: '', criteria: '', destination: '', copy: false, unique: false });
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactAdvancedFilterCatalog?.() || null;
      setCatalog(next);
      if (next) setState(next);
    };
    const off = runtime()?.bus?.on('excelx-advanced-filter-dialog', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactAdvancedFilterCatalog?.()) runtime()?.dlg?.reactAdvancedFilterClose?.(); };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactAdvancedFilterClose?.();
  const update = (patch: Partial<AdvancedFilterState>) => setState((current) => ({ ...current, ...patch }));
  const textField = (key: 'list' | 'criteria' | 'destination', label: string) => <label>{translate(label)}<input type="text" aria-label={translate(label).replace('&', '')} value={state[key]} disabled={key === 'destination' && !state.copy} onChange={(event) => update({ [key]: event.currentTarget.value })} /></label>;
  const apply = () => runtime()?.dlg?.reactAdvancedFilterApply?.(state);
  return <OfficeDialog title={translate('Advanced Filter')} className="excelx-advanced-filter-dialog" style={{ width: 'min(92vw, 440px)' }} onRequestClose={close}>
    <div className="excelx-advanced-filter-dialog__body">
      <fieldset><legend>{translate('Action|filter')}</legend>
        <label><input type="radio" name="excelx-advanced-filter-action" checked={!state.copy} onChange={() => update({ copy: false })} />{translate('&Filter the list, in-place').replace('&', '')}</label>
        <label><input type="radio" name="excelx-advanced-filter-action" checked={state.copy} onChange={() => update({ copy: true })} />{translate('C&opy to another location').replace('&', '')}</label>
      </fieldset>
      {textField('list', '&List range:')}
      {textField('criteria', '&Criteria range:')}
      {textField('destination', 'Copy &to:')}
      <label><input type="checkbox" checked={state.unique} onChange={(event) => update({ unique: event.currentTarget.checked })} />{translate('Unique &records only').replace('&', '')}</label>
    </div>
    <div className="excelx-advanced-filter-dialog__actions"><button type="button" className="primary" onClick={apply}>{translate('OK')}</button><button type="button" onClick={close}>{translate('Cancel')}</button></div>
  </OfficeDialog>;
}
