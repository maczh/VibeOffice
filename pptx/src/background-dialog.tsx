import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useLayoutEffect, useRef, useState, type KeyboardEvent } from 'react';

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
  const [activePaletteIndex, setActivePaletteIndex] = useState(0);
  const paletteId = `${useId().replaceAll(':', '')}-background-colors`;
  const pickerRef = useRef<HTMLButtonElement>(null);
  const paletteItems = useRef<Array<HTMLButtonElement | null>>([]);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactBackgroundCatalog?.() || null;
      setCatalog(next ? { ...next } : null);
    };
    const off = runtime()?.bus?.on('pptx-background-dialog', refresh);
    return () => off?.();
  }, []);
  useLayoutEffect(() => {
    if (!paletteOpen) return;
    setActivePaletteIndex(0);
    paletteItems.current[0]?.focus();
  }, [paletteOpen]);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = (name: string, value?: boolean) => runtime()?.dlg?.reactBackgroundAction?.(name, value);
  const closePalette = () => { setPaletteOpen(false); window.requestAnimationFrame(() => pickerRef.current?.focus()); };
  const chooseColor = (value: string) => { runtime()?.dlg?.reactBackgroundColor?.(value); closePalette(); };
  const paletteItemCount = catalog.colors.length + 2;
  const onPaletteKeyDown = (event: KeyboardEvent<HTMLButtonElement>) => {
    const index = paletteItems.current.indexOf(event.currentTarget);
    let next = index;
    if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); closePalette(); return; }
    if (event.key === 'ArrowLeft') next = index === 0 || index === paletteItemCount - 1 || (index > 1 && (index - 1) % 8 === 0) ? index : index - 1;
    else if (event.key === 'ArrowRight') next = index === 0 || index === paletteItemCount - 1 || index === paletteItemCount - 2 || index % 8 === 0 ? index : index + 1;
    else if (event.key === 'ArrowUp') next = index === 0 ? 0 : index === 1 ? 0 : index === paletteItemCount - 1 ? paletteItemCount - 2 : Math.max(1, index - 8);
    else if (event.key === 'ArrowDown') next = index === 0 ? 1 : index >= paletteItemCount - 2 ? paletteItemCount - 1 : Math.min(paletteItemCount - 2, index + 8);
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = paletteItemCount - 1;
    else if (event.key === 'Enter' || event.key === ' ') { event.stopPropagation(); return; }
    else return;
    event.preventDefault();
    event.stopPropagation();
    setActivePaletteIndex(next);
    paletteItems.current[next]?.focus();
  };
  return <OfficeDialog title={t('Background')} className="pptx-background-dialog" style={{ width: 'min(96vw, 520px)' }} onRequestClose={() => action('close')}>
    <div className="row" style={{ alignItems: 'stretch', gap: 14 }}>
      <section className="col" style={{ flex: '1 1 180px' }}><label>{t('Background fill')}</label><div className="background-preview" style={{ width: '100%', minHeight: 150, display: 'grid', placeItems: 'center', background: '#fff' }} dangerouslySetInnerHTML={{ __html: catalog.preview }} /><button ref={pickerRef} type="button" className="btn small" aria-label={t('Background fill color')} aria-haspopup="menu" aria-expanded={paletteOpen} aria-controls={paletteOpen ? paletteId : undefined} style={{ minWidth: 110, display: 'inline-flex', alignItems: 'center', gap: 6 }} onClick={() => setPaletteOpen((current) => !current)}><span aria-hidden="true" style={{ display: 'inline-block', width: 70, height: 14, border: '1px solid var(--field-border)', background: catalog.swatch }} /><span className="dd-arrow" /></button>{paletteOpen && <div id={paletteId} className="background-color-palette" role="menu" aria-label={t('Background colors')} style={{ display: 'grid', gridTemplateColumns: 'repeat(8, 24px)', gap: 3, padding: 6, border: '1px solid var(--menu-border)', background: 'var(--surface)', width: 'max-content' }}><button ref={(element) => { paletteItems.current[0] = element; }} className="cp-wide" type="button" role="menuitem" tabIndex={activePaletteIndex === 0 ? 0 : -1} style={{ gridColumn: '1 / -1' }} onFocus={() => setActivePaletteIndex(0)} onKeyDown={onPaletteKeyDown} onClick={() => chooseColor('bg1')}>{t('No Fill')}</button>{catalog.colors.map((color, index) => { const itemIndex = index + 1; return <button ref={(element) => { paletteItems.current[itemIndex] = element; }} key={`${color.value}-${index}`} className="cp-sw" type="button" role="menuitem" aria-label={color.label} title={color.label} tabIndex={activePaletteIndex === itemIndex ? 0 : -1} style={{ width: 24, height: 22, padding: 0, background: color.color }} onFocus={() => setActivePaletteIndex(itemIndex)} onKeyDown={onPaletteKeyDown} onClick={() => chooseColor(color.value)} />; })}<label className="cp-wide" style={{ gridColumn: '1 / -1' }}>{t('More Fill Colors...')}<input aria-label={t('Choose custom fill color')} type="color" defaultValue="#3366FF" onChange={(event) => chooseColor(event.currentTarget.value.toUpperCase())} /></label><button ref={(element) => { paletteItems.current[paletteItemCount - 1] = element; }} className="cp-wide" type="button" role="menuitem" tabIndex={activePaletteIndex === paletteItemCount - 1 ? 0 : -1} style={{ gridColumn: '1 / -1' }} onFocus={() => setActivePaletteIndex(paletteItemCount - 1)} onKeyDown={onPaletteKeyDown} onClick={() => { setPaletteOpen(false); void runtime()?.dlg?.reactBackgroundEffects?.(); }}>{t('Fill Effects...')}</button></div>}</section>
      <section className="col" style={{ flex: '1 1 180px' }}><label><input type="checkbox" checked={catalog.omit} onChange={(event) => action('omit', event.currentTarget.checked)} />{t('Omit background &graphics from master').replace('&', '')}</label></section>
    </div>
    <div className="vibeoffice-dialog__actions"><button type="button" onClick={() => action('apply-all')}>{t('Apply to &All').replace('&', '')}</button><button type="button" className="primary" onClick={() => action('apply')}>{t('&Apply').replace('&', '')}</button><button type="button" onClick={() => action('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
