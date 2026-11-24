import { OfficeDialog } from '@vibeoffice/core';
import type { KeyboardEvent } from 'react';
import { useEffect, useMemo, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';
import { PptxPreviewArtwork } from './slides-pane.js';

interface DiagramLayout { id: string; category: string; name: string; description: string; preview: string }
interface DiagramGalleryCatalog { current: string; categories: Array<[string, string]>; layouts: DiagramLayout[] }
interface DiagramGalleryPreview { name: string; description: string; preview: string }
interface DiagramGalleryRuntime {
  diagram?: { reactGalleryCatalog?(): DiagramGalleryCatalog | null; reactGalleryPreview?(id: string): DiagramGalleryPreview | null; reactGalleryAction?(action: string, id?: string): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function handleOptionKeys(event: KeyboardEvent<HTMLDivElement>) {
  const options = [...event.currentTarget.querySelectorAll<HTMLElement>('[role="option"]')];
  if (!options.length) return;
  const activeElement = event.currentTarget.ownerDocument.activeElement as HTMLElement;
  const current = options.indexOf(activeElement);
  let next = current;
  if (event.key === 'ArrowDown' || event.key === 'ArrowRight') next = (current + 1 + options.length) % options.length;
  else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') next = (current - 1 + options.length) % options.length;
  else if (event.key === 'Home') next = 0;
  else if (event.key === 'End') next = options.length - 1;
  else if (event.key === 'Enter' || event.key === ' ') {
    event.preventDefault();
    event.stopPropagation();
    activeElement?.click();
    return;
  } else return;
  const target = options[next];
  if (!target) return;
  event.preventDefault();
  event.stopPropagation();
  target.focus();
  target.click();
}

/** Full IntelliArt layout chooser; catalog, preview rendering, and commit are supplied by the presentation engine. */
export function PptxDiagramGalleryDialog() {
  const runtime = usePptxRuntime<DiagramGalleryRuntime>();
  const [catalog, setCatalog] = useState<DiagramGalleryCatalog | null>(null);
  const [category, setCategory] = useState('all');
  const [selected, setSelected] = useState('');
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.diagram?.reactGalleryCatalog?.() || null;
      setCatalog(next);
      if (next) { setCategory('all'); setSelected(next.current); }
    };
    const off = runtime?.bus?.on('pptx-diagram-gallery', refresh);
    return () => off?.();
  }, []);
  const visible = useMemo(() => catalog?.layouts.filter((layout) => category === 'all' || layout.category === category) || [], [catalog, category]);
  useEffect(() => {
    const first = visible[0];
    if (first && !visible.some((layout) => layout.id === selected)) setSelected(first.id);
  }, [visible, selected]);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const detail = runtime?.diagram?.reactGalleryPreview?.(selected);
  const close = () => runtime?.diagram?.reactGalleryAction?.('close');
  const apply = () => runtime?.diagram?.reactGalleryAction?.('apply', selected);
  return <OfficeDialog title={t('Change IntelliArt Layout')} className="pptx-diagram-gallery-dialog" style={{ width: 'min(96vw, 720px)' }} onRequestClose={close}>
    <div className="sa-gallery">
      <div className="sa-cats list-box" role="listbox" aria-label={t('Categories')} aria-orientation="vertical" onKeyDownCapture={handleOptionKeys}>
        {catalog.categories.map(([id, label]) => <button key={id} type="button" className={`sa-cat${category === id ? ' on' : ''}`} role="option" aria-selected={category === id} tabIndex={category === id ? 0 : -1} onClick={() => setCategory(id)}>{label}</button>)}
      </div>
      <div className="sa-grid" role="listbox" aria-label={t('Layouts')} aria-orientation="vertical" onKeyDownCapture={handleOptionKeys}>
        {visible.map((layout) => <button key={layout.id} type="button" className={`sa-cell${selected === layout.id ? ' on' : ''}`} data-layout-id={layout.id} role="option" aria-label={layout.name} aria-selected={selected === layout.id} tabIndex={selected === layout.id ? 0 : -1} title={layout.name} onClick={() => setSelected(layout.id)} onDoubleClick={apply}><PptxPreviewArtwork markup={layout.preview} /></button>)}
      </div>
      <div className="sa-side">
        {detail && <>
          <div className="sa-big"><PptxPreviewArtwork markup={detail.preview} /></div>
          <b>{detail.name}</b>
          <div className="sa-desc">{detail.description}</div>
        </>}
      </div>
    </div>
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={apply}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
