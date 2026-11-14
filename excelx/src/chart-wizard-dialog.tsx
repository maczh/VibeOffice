import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import './chart-wizard-dialog.css';

type TypeItem = { kind: string; label: string; icon: string; subtypes?: string[] };
type Catalog = { step: number; steps: number; title: string; types: TypeItem[]; subtypes: TypeItem[]; kind: string; range: string; byRows: boolean | null; titleText: string; axX: string; axY: string; gridY: boolean; legend: string; labels: boolean; where: string; sheetName: string; preview: string; editing: boolean };
type Runtime = { chartWizard?: { reactCatalog?(): Catalog | null; reactAction?(action: string, value?: unknown): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React owns the multi-step chart wizard; chart parsing, preview generation, and workbook writes stay in xchart.js. */
export function ExcelxChartWizardDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => { const refresh = () => setCatalog(runtime()?.chartWizard?.reactCatalog?.() || null); const off = runtime()?.bus?.on('excelx-chart-wizard', refresh); return () => { off?.(); runtime()?.chartWizard?.reactAction?.('unmount'); }; }, []);
  if (!catalog) return null;
  const t = (value: string) => runtime()?.t?.(value) || value;
  const act = (action: string, value?: unknown) => runtime()?.chartWizard?.reactAction?.(action, value);
  const cancel = () => { act('cancel'); setCatalog(null); };
  const finish = () => { act('finish'); setCatalog(null); };
  const update = (patch: Record<string, unknown>) => act('update', patch);
  const radio = (name: string, value: string, current: string, label: string, onChange: () => void) => <label><input type="radio" name={name} value={value} checked={current === value} onChange={onChange} />{t(label)}</label>;
  return <OfficeDialog title={t(catalog.title)} className="excelx-chart-wizard" style={{ width: 'min(94vw, 680px)' }} onRequestClose={cancel}>
    <div className="excelx-chart-wizard__content"><div className="excelx-chart-wizard__form">
      {catalog.step === 1 && <><div className="excelx-chart-wizard__types">{catalog.types.map((item) => <button type="button" key={item.kind} className={catalog.subtypes.some((sub) => sub.kind === catalog.kind) && item.kind === catalog.types.find((type) => type.subtypes?.includes(catalog.kind))?.kind ? 'selected' : ''} onClick={() => act('type', item.kind)}><span dangerouslySetInnerHTML={{ __html: item.icon }} />{t(item.label)}</button>)}</div><div className="excelx-chart-wizard__subtypes">{catalog.subtypes.map((item) => <button type="button" aria-label={t(item.label)} key={item.kind} className={catalog.kind === item.kind ? 'selected' : ''} onClick={() => act('kind', item.kind)}><span dangerouslySetInnerHTML={{ __html: item.icon }} /></button>)}</div></>}
      {catalog.step === 2 && <><label>{t('Data range:')}<input value={catalog.range} onChange={(event) => update({ range: event.currentTarget.value })} /></label>{radio('excelx-series-direction', 'rows', catalog.byRows === true ? 'rows' : 'columns', 'Series in rows', () => update({ byRows: true }))}{radio('excelx-series-direction', 'columns', catalog.byRows === true ? 'rows' : 'columns', 'Series in columns', () => update({ byRows: false }))}{catalog.editing && <p className="hint">{t("Leave the range empty to keep the chart's current series.")}</p>}</>}
      {catalog.step === 3 && <><label>{t('Chart title:')}<input value={catalog.titleText} onChange={(event) => update({ title: event.currentTarget.value })} /></label><label>{t('Category (X) axis:')}<input value={catalog.axX} onChange={(event) => update({ axX: event.currentTarget.value })} /></label><label>{t('Value (Y) axis:')}<input value={catalog.axY} onChange={(event) => update({ axY: event.currentTarget.value })} /></label><label><input type="checkbox" checked={catalog.gridY} onChange={(event) => update({ gridY: event.currentTarget.checked })} />{t('Major gridlines (Value axis)')}</label><label><input type="checkbox" checked={catalog.legend !== 'none'} onChange={(event) => update({ legend: event.currentTarget.checked ? 'r' : 'none' })} />{t('Show legend')}</label>{catalog.legend !== 'none' && <label>{t('Legend position:')}<select value={catalog.legend} onChange={(event) => update({ legend: event.currentTarget.value })}>{[['b','Bottom'],['tr','Corner'],['t','Top'],['r','Right'],['l','Left']].map(([value,label])=><option key={value} value={value}>{t(label || '')}</option>)}</select></label>}<label><input type="checkbox" checked={catalog.labels} onChange={(event) => update({ labels: event.currentTarget.checked })} />{t('Value')}</label></>}
      {catalog.step === 4 && <><div>{t('Place chart:')}</div>{radio('excelx-chart-location', 'sheet', catalog.where, 'As new sheet', () => update({ where: 'sheet' }))}<label>{t('Chart sheet name:')}<input value={catalog.sheetName} onChange={(event) => update({ sheetName: event.currentTarget.value })} /></label>{radio('excelx-chart-location', 'object', catalog.where, 'As object in current sheet', () => update({ where: 'object' }))}</>}
    </div><div className="excelx-chart-wizard__preview" dangerouslySetInnerHTML={{ __html: catalog.preview }} /></div>
    <div className="excelx-chart-wizard__actions"><button type="button" disabled={catalog.step === 1} onClick={() => act('back')}>{t('< &Back').replace('&','')}</button><button type="button" disabled={catalog.step === catalog.steps} onClick={() => act('next')}>{t('&Next >').replace('&','')}</button><button type="button" className="primary" onClick={finish}>{t('&Finish').replace('&','')}</button><button type="button" onClick={cancel}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
