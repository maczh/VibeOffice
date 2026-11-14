import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type PrintRangeMode = 'all' | 'current' | 'range';
type PrintWhat = 'slides' | 'handouts' | 'notes' | 'outline';
type PrintColor = 'color' | 'gray' | 'bw';
interface PrintCatalog { slides: number; what: PrintWhat; per: number; color: PrintColor; frame: boolean; hidden: boolean; range: string; rangeMode: PrintRangeMode }
interface PrintRuntime {
  dlg?: { reactPrintCatalog?(): Omit<PrintCatalog, 'rangeMode'> & Partial<Pick<PrintCatalog, 'rangeMode'>> | null; reactPrintAction?(action: string, values?: Omit<PrintCatalog, 'slides'>): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): PrintRuntime | undefined { return (window as Window & { L?: PrintRuntime }).L; }

/** React-owned Print options; pagination and the actual print preview remain in the existing print service. */
export function PptxPrintDialog() {
  const [catalog, setCatalog] = useState<PrintCatalog | null>(null);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactPrintCatalog?.() || null; setCatalog(next ? { ...next, rangeMode: next.rangeMode || 'all' } : null); };
    const off = runtime()?.bus?.on('pptx-print-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const update = (changes: Partial<PrintCatalog>) => setCatalog({ ...catalog, ...changes });
  const close = () => runtime()?.dlg?.reactPrintAction?.('close');
  const preview = () => runtime()?.dlg?.reactPrintAction?.('preview', { what: catalog.what, per: catalog.per, color: catalog.color, frame: catalog.frame, hidden: catalog.hidden, range: catalog.range, rangeMode: catalog.rangeMode });
  return <OfficeDialog title={t('Print')} className="pptx-print-dialog" style={{ width: 'min(94vw, 520px)' }} onRequestClose={close}>
    <fieldset><legend>{t('Print range')}</legend>
      <label><input type="radio" name="pptx-print-range" checked={catalog.rangeMode === 'all'} onChange={() => update({ rangeMode: 'all' })} />{t('&All').replace('&', '')}</label>
      <label><input type="radio" name="pptx-print-range" checked={catalog.rangeMode === 'current'} onChange={() => update({ rangeMode: 'current' })} />{t('Curr&ent slide').replace('&', '')}</label>
      <label><input type="radio" name="pptx-print-range" checked={catalog.rangeMode === 'range'} onChange={() => update({ rangeMode: 'range' })} />{t('&Slides:').replace('&', '')}<input aria-label={t('Slide range')} type="text" value={catalog.range} disabled={catalog.rangeMode !== 'range'} onChange={(event) => update({ range: event.currentTarget.value })} /></label>
    </fieldset>
    <div className="row" style={{ gap: 12 }}>
      <label>{t('Print &what:').replace('&', '')}<select aria-label={t('Print what')} value={catalog.what} onChange={(event) => update({ what: event.currentTarget.value as PrintWhat })}><option value="slides">{t('Slides')}</option><option value="handouts">{t('Handouts')}</option><option value="notes">{t('Notes Pages')}</option><option value="outline">{t('Outline View')}</option></select></label>
      <label>{t('Slides per page:')}<select aria-label={t('Slides per page')} value={catalog.per} onChange={(event) => update({ per: Number(event.currentTarget.value) })}>{[1, 2, 3, 4, 6, 9].map((count) => <option key={count} value={count}>{count}</option>)}</select></label>
    </div>
    <label>{t('Color/gra&yscale:').replace('&', '')}<select aria-label={t('Color/grayscale')} value={catalog.color} onChange={(event) => update({ color: event.currentTarget.value as PrintColor })}><option value="color">{t('Color')}</option><option value="gray">{t('Grayscale')}</option><option value="bw">{t('Pure Black and White')}</option></select></label>
    <label><input type="checkbox" checked={catalog.frame} onChange={(event) => update({ frame: event.currentTarget.checked })} />{t('Fra&me slides').replace('&', '')}</label>
    <label><input type="checkbox" checked={catalog.hidden} onChange={(event) => update({ hidden: event.currentTarget.checked })} />{t('Print &hidden slides').replace('&', '')}</label>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={preview}>{t('Preview')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
