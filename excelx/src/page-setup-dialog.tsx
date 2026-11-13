import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type PrintSettings = { orientation?: string; fit?: boolean; scale?: number; fitW?: number | null; fitH?: number | null; paper?: number; firstPage?: number | null; margins: Record<string, number>; hCenter?: boolean; vCenter?: boolean; header?: string; footer?: string; diffOddEven?: boolean; diffFirst?: boolean; hfScale?: boolean; hfAlign?: boolean; gridLines?: boolean; bw?: boolean; draft?: boolean; headings?: boolean; comments?: string; errors?: string; pageOrder?: string };
type Catalog = { p: PrintSettings; area: string; rows: string; cols: string; first: string; tab: number; error: string; headers: { header: string; footer: string }; headerOptions: Array<{ code: string; label: string }>; papers: Array<{ key: number; label: string }> };
type Runtime = { dlg?: { reactPageSetupCatalog?(): Catalog | null; reactPageSetupAction?(action: string, key?: string, value?: unknown): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string, ...args: unknown[]): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }
const tabs = ['Page', 'Margins', 'Header/Footer', 'Sheet'];

/** React-owned print configuration form backed by Excelx print/workbook services. */
export function ExcelxPageSetupDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactPageSetupCatalog?.() || null);
    const off = runtime()?.bus?.on('excelx-page-setup', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactPageSetupCatalog?.()) runtime()?.dlg?.reactPageSetupAction?.('close'); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = runtime()?.dlg?.reactPageSetupAction;
  const change = (key: string, value: unknown) => action?.('change', key, value);
  const check = (key: keyof PrintSettings, label: string, checked: boolean) => <label key={key}><input type="checkbox" checked={checked} onChange={(event) => change(`print.${key}`, event.currentTarget.checked)} />{t(label).replace('&', '')}</label>;
  const field = (label: string, key: string, value: string | number, type: 'text' | 'number' = 'text', extra?: { min?: number; max?: number; step?: number }) => <label className="excelx-page-setup-dialog__field" key={key}>{t(label).replace('&', '')}<input type={type} value={value} min={extra?.min} max={extra?.max} step={extra?.step} onChange={(event) => change(key, type === 'number' ? Number(event.currentTarget.value) : event.currentTarget.value)} /></label>;
  const radio = (name: string, key: string, value: string, current: string, label: string) => <label key={value}><input type="radio" name={name} value={value} checked={current === value} onChange={() => change(key, value)} />{t(label).replace('&', '')}</label>;
  const p = catalog.p;
  let body;
  if (catalog.tab === 0) body = <div className="excelx-page-setup-dialog__panel">
    <fieldset><legend>{t('Orientation')}</legend>{radio('orientation', 'print.orientation', 'portrait', p.orientation === 'landscape' ? 'landscape' : 'portrait', 'Por&trait')}{radio('orientation', 'print.orientation', 'landscape', p.orientation === 'landscape' ? 'landscape' : 'portrait', '&Landscape')}</fieldset>
    <fieldset><legend>{t('Scaling')}</legend>{radio('scaling', 'print.fit', 'false', p.fit ? 'true' : 'false', '&Adjust to:')}<label><input type="number" min={10} max={400} step={5} value={p.scale || 100} onChange={(event) => change('print.scale', Number(event.currentTarget.value))} />%</label>{radio('scaling', 'print.fit', 'true', p.fit ? 'true' : 'false', '&Fit to:')}<label><input aria-label={t('Pages wide')} type="number" min={0} value={p.fitW ?? 1} onChange={(event) => change('print.fitW', Number(event.currentTarget.value) || null)} />{t('page(s) wide by')}</label><label><input aria-label={t('Pages tall')} type="number" min={0} value={p.fitH ?? 1} onChange={(event) => change('print.fitH', Number(event.currentTarget.value) || null)} />{t('tall')}</label></fieldset>
    <label>{t('Paper si&ze:').replace('&', '')}<select value={p.paper || 1} onChange={(event) => change('print.paper', Number(event.currentTarget.value))}>{catalog.papers.map((paper) => <option key={paper.key} value={paper.key}>{paper.label}</option>)}</select></label>
    {field('Fi&rst page number:', 'first', catalog.first)}
  </div>;
  else if (catalog.tab === 1) body = <div className="excelx-page-setup-dialog__panel">
    <div className="excelx-page-setup-dialog__margins">{field('&Top:', 'margin.t', p.margins.t ?? 0, 'number', { min: 0, max: 20, step: 0.25 })}{field('H&eader:', 'margin.header', p.margins.header ?? 0, 'number', { min: 0, max: 20, step: 0.25 })}{field('&Left:', 'margin.l', p.margins.l ?? 0, 'number', { min: 0, max: 20, step: 0.25 })}{field('&Right:', 'margin.r', p.margins.r ?? 0, 'number', { min: 0, max: 20, step: 0.25 })}{field('&Bottom:', 'margin.b', p.margins.b ?? 0, 'number', { min: 0, max: 20, step: 0.25 })}{field('&Footer:', 'margin.footer', p.margins.footer ?? 0, 'number', { min: 0, max: 20, step: 0.25 })}</div>
    <fieldset><legend>{t('Center on page')}</legend>{check('hCenter', 'Hori&zontally', !!p.hCenter)}{check('vCenter', '&Vertically', !!p.vCenter)}</fieldset><p>{t('Margins are in inches.')}</p>
  </div>;
  else if (catalog.tab === 2) body = <div className="excelx-page-setup-dialog__panel">
    <p>{t('Header:')} {catalog.headers.header}</p><label>{t('He&ader:').replace('&', '')}<select value={p.header || ''} onChange={(event) => change('print.header', event.currentTarget.value)}>{catalog.headerOptions.map((item) => <option key={item.code} value={item.code} translate="no">{item.label}</option>)}</select></label><button type="button" onClick={() => action?.('custom', 'header')}>{t('&Custom Header...').replace('&', '')}</button>
    <p>{t('Footer:')} {catalog.headers.footer}</p><label>{t('&Footer:').replace('&', '')}<select value={p.footer || ''} onChange={(event) => change('print.footer', event.currentTarget.value)}>{catalog.headerOptions.map((item) => <option key={item.code} value={item.code} translate="no">{item.label}</option>)}</select></label><button type="button" onClick={() => action?.('custom', 'footer')}>{t('C&ustom Footer...').replace('&', '')}</button>
    {check('diffOddEven', 'Different &odd and even pages', !!p.diffOddEven)}{check('diffFirst', 'Different first &page', !!p.diffFirst)}{check('hfScale', 'Scale with document', p.hfScale !== false)}{check('hfAlign', 'Align with page margins', p.hfAlign !== false)}
  </div>;
  else body = <div className="excelx-page-setup-dialog__panel">
    {field('Print &area:', 'area', catalog.area)}<fieldset><legend>{t('Print titles')}</legend>{field('&Rows to repeat at top:', 'rows', catalog.rows)}{field('&Columns to repeat at left:', 'cols', catalog.cols)}</fieldset>
    <fieldset><legend>{t('Print')}</legend>{check('gridLines', '&Gridlines', !!p.gridLines)}{check('bw', '&Black and white', !!p.bw)}{check('draft', 'Draft &quality', !!p.draft)}{check('headings', 'Row and colu&mn headings', !!p.headings)}<label>{t('Co&mments:').replace('&', '')}<select value={p.comments || 'none'} onChange={(event) => change('print.comments', event.currentTarget.value)}>{[['none','(None)'],['atEnd','At end of sheet'],['asDisplayed','As displayed on sheet']].map(([value,label]) => <option key={value || ''} value={value || ''}>{t(label || '')}</option>)}</select></label><label>{t('Cell &errors as:').replace('&', '')}<select value={p.errors || 'displayed'} onChange={(event) => change('print.errors', event.currentTarget.value)}>{[['displayed','displayed'],['blank','<blank>'],['dash','--'],['NA','#N/A']].map(([value,label]) => <option key={value || ''} value={value || ''}>{t(label || '')}</option>)}</select></label></fieldset>
    <fieldset><legend>{t('Page order')}</legend>{radio('page-order', 'print.pageOrder', 'downThenOver', p.pageOrder === 'overThenDown' ? 'overThenDown' : 'downThenOver', '&Down, then over')}{radio('page-order', 'print.pageOrder', 'overThenDown', p.pageOrder === 'overThenDown' ? 'overThenDown' : 'downThenOver', 'O&ver, then down')}</fieldset>
  </div>;
  return <OfficeDialog title={t('Page Setup')} className="excelx-page-setup-dialog" style={{ width: 'min(96vw, 620px)' }} onRequestClose={() => action?.('close')}>
    <div className="excelx-page-setup-dialog__tabs" role="tablist">{tabs.map((tab, index) => <button type="button" role="tab" aria-selected={catalog.tab === index} key={tab} onClick={() => change('tab', index)}>{t(tab)}</button>)}</div>
    {body}{catalog.error && <p role="alert">{t(catalog.error)}</p>}
    <div className="excelx-page-setup-dialog__actions"><button type="button" onClick={() => action?.('print')}>{t('&Print...').replace('&', '')}</button><button type="button" onClick={() => action?.('preview')}>{t('Print Previe&w').replace('&', '')}</button><button type="button" className="primary" onClick={() => action?.('apply')}>{t('OK')}</button><button type="button" onClick={() => action?.('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
