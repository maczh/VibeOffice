import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { useExcelxRuntime } from './editor-host.js';
import './print-dialog.css';

type PrintOptions = { what: 'selection' | 'sheets' | 'workbook'; range: 'all' | 'pages'; from: number; to: number; copies: number };
type Catalog = { pageCount: number; options: PrintOptions };
type Runtime = { dlg?: { reactPrintCatalog?(): Catalog | null; reactPrintAction?(action: string, options?: PrintOptions): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };

/** Print options are React-owned; PDF generation and preview remain in the spreadsheet print service. */
export function ExcelxPrintDialog() {
  const runtime = useExcelxRuntime<Runtime>();
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [options, setOptions] = useState<PrintOptions>({ what: 'sheets', range: 'all', from: 1, to: 1, copies: 1 });
  useEffect(() => {
    const refresh = () => { const next = runtime?.dlg?.reactPrintCatalog?.() || null; setCatalog(next); if (next) setOptions(next.options); };
    const off = runtime?.bus?.on('excelx-print-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const action = (kind: string) => runtime?.dlg?.reactPrintAction?.(kind, options);
  const set = (patch: Partial<PrintOptions>) => setOptions((current) => ({ ...current, ...patch }));
  const radio = (group: string, value: string, label: string, selected: string, onSelect: () => void) => <label><input type="radio" name={group} value={value} checked={selected === value} onChange={onSelect} />{t(label).replace('&', '')}</label>;
  return <OfficeDialog title={t('Print')} className="excelx-print-dialog" style={{ width: 'min(94vw, 440px)' }} onRequestClose={() => action('cancel')}>
    <div className="excelx-print-dialog__body">
      <fieldset><legend>{t('Printer')}</legend><label>{t('&Name:').replace('&', '')}<select value="pdf" disabled><option value="pdf">Excelx PDF Writer</option></select></label><p className="hint">{t('Excelx prints by creating a PDF that you can open and print on any printer.')}</p></fieldset>
      <div className="excelx-print-dialog__choices">
        <fieldset><legend>{t('Print range')}</legend>
          {radio('excelx-print-range', 'all', '&All', options.range, () => set({ range: 'all' }))}
          <div className="excelx-print-dialog__pages">{radio('excelx-print-range', 'pages', 'Pa&ge(s)', options.range, () => set({ range: 'pages' }))}<label>{t('From:')}<input aria-label={t('From:')} type="number" min="1" max={catalog.pageCount} value={options.from} disabled={options.range !== 'pages'} onChange={(event) => set({ from: Number(event.currentTarget.value) })} /></label><label>{t('To:')}<input aria-label={t('To:')} type="number" min="1" max={catalog.pageCount} value={options.to} disabled={options.range !== 'pages'} onChange={(event) => set({ to: Number(event.currentTarget.value) })} /></label></div>
        </fieldset>
        <fieldset><legend>{t('Print what')}</legend>
          {radio('excelx-print-what', 'selection', 'Selectio&n', options.what, () => set({ what: 'selection' }))}
          {radio('excelx-print-what', 'sheets', 'Active s&heet(s)', options.what, () => set({ what: 'sheets' }))}
          {radio('excelx-print-what', 'workbook', '&Entire workbook', options.what, () => set({ what: 'workbook' }))}
        </fieldset>
      </div>
      <label className="excelx-print-dialog__copies">{t('Number of &copies:').replace('&', '')}<input aria-label={t('Number of &copies:').replace('&', '')} type="number" min="1" max="99" value={options.copies} onChange={(event) => set({ copies: Math.max(1, Math.min(99, Number(event.currentTarget.value) || 1)) })} /></label>
    </div>
    <div className="excelx-print-dialog__actions"><button type="button" onClick={() => action('preview')}>{t('Pre&view').replace('&', '')}</button><span /><button type="button" className="primary" onClick={() => action('print')}>{t('OK')}</button><button type="button" onClick={() => action('cancel')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
