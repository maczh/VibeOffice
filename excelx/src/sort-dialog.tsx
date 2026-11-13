import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface SortKey { c: number | null; desc: boolean }
interface SortState { keys: SortKey[]; caseSensitive: boolean; byCols: boolean; list: number | null; header: boolean }
interface SortChoice { value: number; label: string }
interface SortCatalog { range: { r1: number; c1: number; r2: number; c2: number }; header: boolean; selection: { row: number; column: number }; state: Omit<SortState, 'header'>; columns: SortChoice[]; rows: SortChoice[]; lists: { value: number; label: string }[] }
interface SortRuntime {
  dlg?: { reactSortCatalog?(): SortCatalog | null; reactSortApply?(state: SortState): boolean; reactSortClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}

function runtime(): SortRuntime | undefined {
  return (window as Window & { L?: SortRuntime }).L;
}

export function ExcelxSortDialog() {
  const [catalog, setCatalog] = useState<SortCatalog | null>(null);
  const [state, setState] = useState<SortState>({ keys: [{ c: null, desc: false }, { c: null, desc: false }, { c: null, desc: false }], caseSensitive: false, byCols: false, list: null, header: false });
  const [optionsOpen, setOptionsOpen] = useState(false);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactSortCatalog?.() || null;
      setCatalog(next);
      if (next) { setState({ ...next.state, header: next.header }); setOptionsOpen(false); }
    };
    const off = runtime()?.bus?.on('excelx-sort-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactSortCatalog?.()) runtime()?.dlg?.reactSortClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string, ...args: unknown[]) => runtime()?.t?.(message, ...args) || message;
  const close = () => { setOptionsOpen(false); runtime()?.dlg?.reactSortClose?.(); };
  const apply = () => runtime()?.dlg?.reactSortApply?.(state);
  const updateKey = (index: number, update: Partial<SortKey>) => setState((previous) => ({ ...previous, keys: previous.keys.map((key, i) => i === index ? { ...key, ...update } : key) }));
  const options = state.byCols ? catalog.rows : catalog.columns;
  const orientation = (byCols: boolean) => setState((previous) => ({ ...previous, byCols, keys: previous.keys.map((key) => ({ ...key, c: null })).map((key, index) => index === 0 ? { ...key, c: byCols ? catalog.selection.row : catalog.selection.column } : key) }));
  return <>
    <OfficeDialog title={translate('Sort')} className="excelx-sort-dialog" style={{ width: 'min(92vw, 460px)' }} onRequestClose={close}>
      <div className="excelx-sort-dialog__keys">
        {state.keys.map((key, index) => <fieldset key={index}>
          <legend>{translate(index === 0 ? 'Sort by' : 'Then by')}</legend>
          <select aria-label={translate(index === 0 ? 'Sort by' : 'Then by')} value={key.c ?? ''} onChange={(event) => { const value = event.currentTarget.value; updateKey(index, { c: value === '' ? null : Number(value) }); }}>
            {index > 0 && <option value="">{translate('(none)')}</option>}
            {options.map(({ value, label }) => <option key={value} value={value} translate="no">{label}</option>)}
          </select>
          <label><input type="radio" name={`excelx-sort-order-${index}`} value="asc" checked={!key.desc} onChange={() => updateKey(index, { desc: false })} />{translate('&Ascending').replace('&', '')}</label>
          <label><input type="radio" name={`excelx-sort-order-${index}`} value="desc" checked={key.desc} onChange={() => updateKey(index, { desc: true })} />{translate('&Descending').replace('&', '')}</label>
        </fieldset>)}
      </div>
      <fieldset className="excelx-sort-dialog__header"><legend>{translate('My data range has')}</legend>
        <label><input type="radio" name="excelx-sort-header" checked={state.header} onChange={() => setState((previous) => ({ ...previous, header: true }))} />{translate('Header &row').replace('&', '')}</label>
        <label><input type="radio" name="excelx-sort-header" checked={!state.header} onChange={() => setState((previous) => ({ ...previous, header: false }))} />{translate('No header ro&w').replace('&', '')}</label>
      </fieldset>
      <div className="excelx-sort-dialog__actions">
        <button type="button" onClick={() => setOptionsOpen(true)}>{translate('&Options...').replace('&', '')}</button>
        <button type="button" className="primary" onClick={apply}>{translate('OK')}</button>
        <button type="button" onClick={close}>{translate('Cancel')}</button>
      </div>
    </OfficeDialog>
    {optionsOpen && <OfficeDialog title={translate('Sort Options')} className="excelx-sort-options-dialog" style={{ width: 'min(92vw, 360px)' }} onRequestClose={() => setOptionsOpen(false)}>
      <label>{translate('&First key sort order').replace('&', '')}<select aria-label={translate('&First key sort order').replace('&', '')} value={state.list ?? ''} onChange={(event) => { const value = event.currentTarget.value; setState((previous) => ({ ...previous, list: value === '' ? null : Number(value) })); }}><option value="">{translate('Normal')}</option>{catalog.lists.map(({ value, label }) => <option key={value} value={value} translate="no">{label}</option>)}</select></label>
      <label><input type="checkbox" checked={state.caseSensitive} onChange={(event) => { const checked = event.currentTarget.checked; setState((previous) => ({ ...previous, caseSensitive: checked })); }} />{translate('&Case sensitive').replace('&', '')}</label>
      <fieldset><legend>{translate('Orientation')}</legend>
        <label><input type="radio" name="excelx-sort-orientation" value="rows" checked={!state.byCols} onChange={() => orientation(false)} />{translate('Sort &top to bottom').replace('&', '')}</label>
        <label><input type="radio" name="excelx-sort-orientation" value="columns" checked={state.byCols} onChange={() => orientation(true)} />{translate('Sort &left to right').replace('&', '')}</label>
      </fieldset>
      <div className="excelx-sort-options-dialog__actions"><button type="button" className="primary" onClick={() => setOptionsOpen(false)}>{translate('OK')}</button><button type="button" onClick={() => setOptionsOpen(false)}>{translate('Cancel')}</button></div>
    </OfficeDialog>}
  </>;
}
