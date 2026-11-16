import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface FinderSlide { index: number; title: string; thumb: string }
interface SlideFinderCatalog { name: string; keep: boolean; slides: FinderSlide[]; selected: number[] }
interface SlideFinderRuntime {
  dlg?: { reactSlideFinderCatalog?(): SlideFinderCatalog | null; reactSlideFinderAction?(action: string, index?: number, value?: boolean): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}
function runtime(): SlideFinderRuntime | undefined { return (window as Window & { L?: SlideFinderRuntime }).L; }

/** React owns slide selection and import controls; presentation parsing and insertion stay engine-backed. */
export function PptxSlideFinderDialog() {
  const [catalog, setCatalog] = useState<SlideFinderCatalog | null>(null);
  const [selected, setSelected] = useState<number[]>([]);
  const [keep, setKeep] = useState(false);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactSlideFinderCatalog?.() || null;
      setCatalog(next ? { ...next, slides: [...next.slides], selected: [...next.selected] } : null);
      if (next) { setSelected([...next.selected]); setKeep(next.keep); }
    };
    const off = runtime()?.bus?.on('pptx-slide-finder-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string, ...args: unknown[]) => runtime()?.t?.(message, ...args) || message;
  const close = () => runtime()?.dlg?.reactSlideFinderAction?.('close');
  const toggle = (index: number, checked: boolean) => {
    runtime()?.dlg?.reactSlideFinderAction?.('select', index, checked);
    setSelected((current) => checked ? [...new Set([...current, index])].sort((a, b) => a - b) : current.filter((item) => item !== index));
  };
  const setKeepFormatting = (checked: boolean) => { setKeep(checked); runtime()?.dlg?.reactSlideFinderAction?.('keep', undefined, checked); };
  const insert = (all: boolean) => runtime()?.dlg?.reactSlideFinderAction?.(all ? 'insert-all' : 'insert');
  return <OfficeDialog title={t('Slide Finder')} className="pptx-slide-finder-dialog" style={{ width: 'min(96vw, 700px)' }} onRequestClose={close}>
    <div className="col">
      <div>{t('Presentation: {0}', catalog.name)}</div>
      <div className="slide-finder-grid" role="group" aria-label={t('Slides in presentation')} style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(130px, 1fr))', gap: 10, maxHeight: 320, overflow: 'auto', padding: 4, background: '#fff', border: '1px solid var(--field-border)' }}>
        {catalog.slides.map((slide) => <label key={slide.index} className="slide-finder-item" style={{ display: 'flex', flexDirection: 'column', gap: 3, alignItems: 'flex-start', cursor: 'default' }}>
          <span className="slide-finder-thumb" style={{ border: '1px solid var(--field-border)' }} dangerouslySetInnerHTML={{ __html: slide.thumb }} />
          <span className="row" style={{ gap: 4, alignItems: 'center' }}><input type="checkbox" aria-label={t('Slide {0}: {1}', slide.index + 1, slide.title)} checked={selected.includes(slide.index)} onChange={(event) => toggle(slide.index, event.currentTarget.checked)} /><span>{slide.index + 1}. {slide.title.slice(0, 24)}</span></span>
        </label>)}
      </div>
      <label><input type="checkbox" checked={keep} onChange={(event) => setKeepFormatting(event.currentTarget.checked)} />{t('&Keep source formatting').replace('&', '')}</label>
    </div>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" disabled={!selected.length} onClick={() => insert(false)}>{t('&Insert').replace('&', '')}</button><button type="button" disabled={!catalog.slides.length} onClick={() => insert(true)}>{t('Insert &All').replace('&', '')}</button><button type="button" onClick={close}>{t('Close')}</button></div>
  </OfficeDialog>;
}
