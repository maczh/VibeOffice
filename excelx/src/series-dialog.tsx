import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type SeriesType = 'linear' | 'growth' | 'date' | 'auto';
type SeriesUnit = 'day' | 'weekday' | 'month' | 'year';
interface SeriesState { rows: boolean; type: SeriesType; unit: SeriesUnit; trend: boolean; step: string; stop: string }
interface SeriesCatalog { state: SeriesState; range: { rows: number; columns: number } }
interface SeriesRuntime {
  dlg?: { reactSeriesCatalog?(): SeriesCatalog | null; reactSeriesApply?(state: SeriesState): boolean; reactSeriesClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): SeriesRuntime | undefined {
  return (window as Window & { L?: SeriesRuntime }).L;
}

const types: [SeriesType, string][] = [['linear', '&Linear'], ['growth', '&Growth'], ['date', 'Da&te'], ['auto', 'Auto&Fill']];
const units: [SeriesUnit, string][] = [['day', '&Day'], ['weekday', '&Weekday'], ['month', '&Month'], ['year', '&Year']];

export function ExcelxSeriesDialog() {
  const [catalog, setCatalog] = useState<SeriesCatalog | null>(null);
  const [state, setState] = useState<SeriesState>({ rows: true, type: 'linear', unit: 'day', trend: false, step: '1', stop: '' });
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactSeriesCatalog?.() || null;
      setCatalog(next);
      if (next) setState(next.state);
    };
    const off = runtime()?.bus?.on('excelx-series-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactSeriesCatalog?.()) runtime()?.dlg?.reactSeriesClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactSeriesClose?.();
  const apply = () => runtime()?.dlg?.reactSeriesApply?.(state);
  const radioGroup = <T extends string,>(name: string, options: [T, string][], value: T, update: (value: T) => void) => <div role="radiogroup" aria-label={translate(name)}>{options.map(([key, label]) => <label key={key}><input type="radio" name={`excelx-series-${name}`} value={key} checked={value === key} onChange={() => update(key)} />{translate(label).replace('&', '')}</label>)}</div>;
  return <OfficeDialog title={translate('Series')} className="excelx-series-dialog" style={{ width: 'min(94vw, 560px)' }} onRequestClose={close}>
    <div className="excelx-series-dialog__groups">
      <fieldset><legend>{translate('Series in')}</legend>{radioGroup('direction', [['rows', '&Rows'], ['columns', '&Columns']], state.rows ? 'rows' : 'columns', (value) => setState((previous) => ({ ...previous, rows: value === 'rows' })))}</fieldset>
      <fieldset><legend>{translate('Type')}</legend>{radioGroup('type', types, state.type, (type) => setState((previous) => ({ ...previous, type })))}</fieldset>
      <fieldset><legend>{translate('Date unit')}</legend>{radioGroup('unit', units, state.unit, (unit) => setState((previous) => ({ ...previous, unit })))}</fieldset>
    </div>
    <label className="excelx-series-dialog__trend"><input type="checkbox" checked={state.trend} onChange={(event) => { const checked = event.currentTarget.checked; setState((previous) => ({ ...previous, trend: checked })); }} />{translate('T&rend').replace('&', '')}</label>
    <div className="excelx-series-dialog__values">
      <label>{translate('&Step value:').replace('&', '')}<input type="text" value={state.step} onChange={(event) => { const value = event.currentTarget.value; setState((previous) => ({ ...previous, step: value })); }} /></label>
      <label>{translate('St&op value:').replace('&', '')}<input type="text" value={state.stop} onChange={(event) => { const value = event.currentTarget.value; setState((previous) => ({ ...previous, stop: value })); }} /></label>
    </div>
    <div className="excelx-series-dialog__actions">
      <button type="button" className="primary" onClick={apply}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
