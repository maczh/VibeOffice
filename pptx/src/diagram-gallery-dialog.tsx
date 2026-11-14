import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useMemo, useState } from 'react';

interface DiagramLayout { id: string; category: string; name: string; description: string; preview: string }
interface DiagramGalleryCatalog { current: string; categories: Array<[string, string]>; layouts: DiagramLayout[] }
interface DiagramGalleryPreview { name: string; description: string; preview: string }
interface DiagramGalleryRuntime {
  diagram?: { reactGalleryCatalog?(): DiagramGalleryCatalog | null; reactGalleryPreview?(id: string): DiagramGalleryPreview | null; reactGalleryAction?(action: string, id?: string): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): DiagramGalleryRuntime | undefined { return (window as Window & { L?: DiagramGalleryRuntime }).L; }

/** Full IntelliArt layout chooser; catalog, preview rendering, and commit are supplied by the presentation engine. */
export function PptxDiagramGalleryDialog() {
  const [catalog, setCatalog] = useState<DiagramGalleryCatalog | null>(null);
  const [category, setCategory] = useState('all');
  const [selected, setSelected] = useState('');
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.diagram?.reactGalleryCatalog?.() || null;
      setCatalog(next);
      if (next) { setCategory('all'); setSelected(next.current); }
    };
    const off = runtime()?.bus?.on('pptx-diagram-gallery', refresh);
    return () => off?.();
  }, []);
  const visible = useMemo(() => catalog?.layouts.filter((layout) => category === 'all' || layout.category === category) || [], [catalog, category]);
  useEffect(() => {
    const first = visible[0];
    if (first && !visible.some((layout) => layout.id === selected)) setSelected(first.id);
  }, [visible, selected]);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const detail = runtime()?.diagram?.reactGalleryPreview?.(selected);
  const close = () => runtime()?.diagram?.reactGalleryAction?.('close');
  const apply = () => runtime()?.diagram?.reactGalleryAction?.('apply', selected);
  return <OfficeDialog title={t('Change IntelliArt Layout')} className="pptx-diagram-gallery-dialog" style={{ width: 'min(96vw, 720px)' }} onRequestClose={close}>
    <div className="sa-gallery">
      <div className="sa-cats list-box" role="listbox" aria-label={t('Categories')}>
        {catalog.categories.map(([id, label]) => <button key={id} type="button" className={`sa-cat${category === id ? ' on' : ''}`} role="option" aria-selected={category === id} onClick={() => setCategory(id)}>{label}</button>)}
      </div>
      <div className="sa-grid" role="listbox" aria-label={t('Layouts')}>
        {visible.map((layout) => <button key={layout.id} type="button" className={`sa-cell${selected === layout.id ? ' on' : ''}`} data-layout-id={layout.id} aria-label={layout.name} aria-selected={selected === layout.id} title={layout.name} onClick={() => setSelected(layout.id)} onDoubleClick={apply} dangerouslySetInnerHTML={{ __html: layout.preview }} />)}
      </div>
      <div className="sa-side">
        {detail && <>
          <div className="sa-big" dangerouslySetInnerHTML={{ __html: detail.preview }} />
          <b>{detail.name}</b>
          <div className="sa-desc">{detail.description}</div>
        </>}
      </div>
    </div>
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={apply}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
