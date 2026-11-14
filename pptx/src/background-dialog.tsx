import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface BackgroundCatalog { fill: unknown; omit: boolean; preview: string; swatch: string; colors: Array<{ value: string; label: string; color: string }> }
interface BackgroundRuntime {
  dlg?: { reactBackgroundCatalog?(): BackgroundCatalog | null; reactBackgroundAction?(action: string, value?: boolean): boolean; reactBackgroundColor?(value: string): boolean; reactBackgroundEffects?(): boolean | Promise<boolean> };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): BackgroundRuntime | undefined { return (window as Window & { L?: BackgroundRuntime }).L; }

/** React owns the slide-background dialog; fill colors and model mutations remain engine-backed. */
export function PptxBackgroundDialog() {
  const [catalog, setCatalog] = useState<BackgroundCatalog | null>(null);
  const [paletteOpen, setPaletteOpen] = useState(false);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactBackgroundCatalog?.() || null;
      setCatalog(next ? { ...next } : null);
    };
    const off = runtime()?.bus?.on('pptx-background-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = (name: string, value?: boolean) => runtime()?.dlg?.reactBackgroundAction?.(name, value);
  const chooseColor = (value: string) => { runtime()?.dlg?.reactBackgroundColor?.(value); setPaletteOpen(false); };
  return <OfficeDialog title={t('Background')} className="pptx-background-dialog" style={{ width: 'min(96vw, 520px)' }} onRequestClose={() => action('close')}>
    <div className="row" style={{ alignItems: 'stretch', gap: 14 }}>
      <section className="col" style={{ flex: '1 1 180px' }}><label>{t('Background fill')}</label><div className="background-preview" style={{ width: '100%', minHeight: 150, display: 'grid', placeItems: 'center', background: '#fff' }} dangerouslySetInnerHTML={{ __html: catalog.preview }} /><button type="button" className="btn small" aria-label={t('Background fill color')} aria-expanded={paletteOpen} style={{ minWidth: 110, display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => setPaletteOpen((current) => !current)}><span aria-hidden="true" style={{ display: 'inline-block', width: 70, height: 14, border: '1px solid #555', background: catalog.swatch }} /><span className="dd-arrow" /></button>{paletteOpen && <div className="background-color-palette" role="menu" aria-label={t('Background colors')} style={{ display: 'grid', gridTemplateColumns: 'repeat(8, 24px)', gap: 3, padding: 6, border: '1px solid #888', background: '#f4f4f4', width: 'max-content' }}><button className="cp-wide" type="button" role="menuitem" style={{ gridColumn: '1 / -1' }} onClick={() => chooseColor('bg1')}>{t('No Fill')}</button>{catalog.colors.map((color, index) => <button key={`${color.value}-${index}`} className="cp-sw" type="button" role="menuitem" aria-label={color.label} title={color.label} style={{ width: 24, height: 22, padding: 0, background: color.color }} onClick={() => chooseColor(color.value)} />)}<label className="cp-wide" style={{ gridColumn: '1 / -1' }}>{t('More Fill Colors...')}<input aria-label={t('Choose custom fill color')} type="color" defaultValue="#3366FF" onChange={(event) => chooseColor(event.currentTarget.value.toUpperCase())} /></label><button className="cp-wide" type="button" role="menuitem" style={{ gridColumn: '1 / -1' }} onClick={() => { setPaletteOpen(false); void runtime()?.dlg?.reactBackgroundEffects?.(); }}>{t('Fill Effects...')}</button></div>}</section>
      <section className="col" style={{ flex: '1 1 180px' }}><label><input type="checkbox" checked={catalog.omit} onChange={(event) => action('omit', event.currentTarget.checked)} />{t('Omit background &graphics from master').replace('&', '')}</label></section>
    </div>
    <div className="vibeoffice-dialog__actions"><button type="button" onClick={() => action('apply-all')}>{t('Apply to &All').replace('&', '')}</button><button type="button" className="primary" onClick={() => action('apply')}>{t('&Apply').replace('&', '')}</button><button type="button" onClick={() => action('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
