import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useRef, useState } from 'react';

interface PhotoAlbumCatalog { layouts: Array<[string, string]>; frames: Array<[string, string]> }
interface AlbumPhoto { file: File; name: string; url: string }
interface PhotoAlbumRuntime {
  dlg?: { reactPhotoAlbumCatalog?(): PhotoAlbumCatalog | null; reactPhotoAlbumAction?(action: string, photos?: AlbumPhoto[], layout?: string, frame?: string, captions?: boolean): boolean | Promise<boolean> };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): PhotoAlbumRuntime | undefined { return (window as Window & { L?: PhotoAlbumRuntime }).L; }

/** React-owned photo selection/reordering and preview; image import and slide generation stay engine-backed. */
export function PptxPhotoAlbumDialog() {
  const [catalog, setCatalog] = useState<PhotoAlbumCatalog | null>(null);
  const [photos, setPhotos] = useState<AlbumPhoto[]>([]);
  const [selected, setSelected] = useState(-1);
  const [layout, setLayout] = useState('1');
  const [frame, setFrame] = useState('rect');
  const [captions, setCaptions] = useState(false);
  const [busy, setBusy] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);
  const urls = useRef(new Set<string>());
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactPhotoAlbumCatalog?.() || null;
      setCatalog(next ? { layouts: [...next.layouts], frames: [...next.frames] } : null);
      if (next) { setPhotos([]); setSelected(-1); setLayout('1'); setFrame('rect'); setCaptions(false); setBusy(false); }
    };
    const off = runtime()?.bus?.on('pptx-photo-album-dialog', refresh);
    return () => {
      off?.();
      for (const url of urls.current) URL.revokeObjectURL(url);
      urls.current.clear();
    };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactPhotoAlbumAction?.('close');
  const addFiles = (files: FileList | null) => {
    if (!files) return;
    const added = [...files].filter((file) => file.type.startsWith('image/')).map((file) => {
      const url = URL.createObjectURL(file);
      urls.current.add(url);
      return { file, name: file.name, url };
    });
    if (added.length) {
      const firstAdded = photos.length;
      setSelected(firstAdded);
      setPhotos((current) => [...current, ...added]);
    }
    if (fileInput.current) fileInput.current.value = '';
  };
  const remove = () => {
    if (selected < 0) return;
    const [removed] = photos.slice(selected, selected + 1);
    if (removed) { URL.revokeObjectURL(removed.url); urls.current.delete(removed.url); }
    setPhotos(photos.filter((_, index) => index !== selected));
    setSelected(Math.min(selected, photos.length - 2));
  };
  const move = (direction: -1 | 1) => {
    const next = selected + direction;
    if (selected < 0 || next < 0 || next >= photos.length) return;
    const reordered = [...photos];
    const current = reordered[selected];
    const target = reordered[next];
    if (!current || !target) return;
    reordered[selected] = target;
    reordered[next] = current;
    setPhotos(reordered);
    setSelected(next);
  };
  const create = async () => {
    if (!photos.length || busy) return;
    setBusy(true);
    const result = await runtime()?.dlg?.reactPhotoAlbumAction?.('create', photos, layout, frame, captions);
    if (!result) setBusy(false);
  };
  return <OfficeDialog title={t('Photo Album')} className="pptx-photo-album-dialog" style={{ width: 'min(96vw, 760px)' }} onRequestClose={close}>
    <div className="col">
      <fieldset><legend>{t('Album Content')}</legend><div className="row" style={{ alignItems: 'stretch', gap: 12 }}>
        <section className="col" style={{ flex: '1 1 230px' }}><label>{t('Insert picture from:')}<button type="button" onClick={() => fileInput.current?.click()}>{t('&File/Disk...').replace('&', '')}</button></label><input ref={fileInput} type="file" accept="image/*" multiple hidden onChange={(event) => addFiles(event.currentTarget.files)} />
          <label>{t('Pictures in album:')}<select aria-label={t('Pictures in album')} size={8} value={selected < 0 ? '' : photos[selected]?.url || ''} onChange={(event) => setSelected(photos.findIndex((photo) => photo.url === event.currentTarget.value))}>{photos.map((photo, index) => <option key={photo.url} value={photo.url}>{index + 1}  {photo.name}</option>)}</select></label>
          <div className="row"><button type="button" aria-label={t('Move up')} disabled={selected <= 0} onClick={() => move(-1)}>↑</button><button type="button" aria-label={t('Move down')} disabled={selected < 0 || selected >= photos.length - 1} onClick={() => move(1)}>↓</button><button type="button" disabled={selected < 0} onClick={remove}>{t('Re&move').replace('&', '')}</button></div>
        </section>
        <section className="col" style={{ flex: '1 1 220px' }}><label>{t('Preview:')}<div className="preview-box" style={{ width: '100%', minHeight: 180, display: 'grid', placeItems: 'center' }}>{photos[selected] && <img src={photos[selected].url} alt={photos[selected].name} style={{ maxWidth: '100%', maxHeight: 260, objectFit: 'contain' }} />}</div></label></section>
      </div></fieldset>
      <fieldset><legend>{t('Album Layout')}</legend><div className="col"><label>{t('&Picture layout:').replace('&', '')}<select aria-label={t('Picture layout')} value={layout} onChange={(event) => setLayout(event.currentTarget.value)}>{catalog.layouts.map(([value, label]) => <option key={value} value={value}>{t(label)}</option>)}</select></label><label>{t('Frame s&hape:').replace('&', '')}<select aria-label={t('Frame shape')} value={frame} onChange={(event) => setFrame(event.currentTarget.value)}>{catalog.frames.map(([value, label]) => <option key={value} value={value}>{t(label)}</option>)}</select></label><label><input type="checkbox" checked={captions} onChange={(event) => setCaptions(event.currentTarget.checked)} />{t('Captions &below ALL pictures').replace('&', '')}</label></div></fieldset>
    </div>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" disabled={!photos.length || busy} onClick={create}>{busy ? t('Creating...') : t('&Create').replace('&', '')}</button><button type="button" disabled={busy} onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
