import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface FindCatalog { what: string; with: string; mc: boolean; ww: boolean; replace: boolean }
interface FindRuntime {
  dlg?: { reactFindCatalog?(): FindCatalog | null; reactFindAction?(action: string, values?: FindCatalog): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): FindRuntime | undefined { return (window as Window & { L?: FindRuntime }).L; }

/** React-owned Find/Replace controls; matching and presentation mutations remain in the presentation engine. */
export function PptxFindDialog() {
  const [catalog, setCatalog] = useState<FindCatalog | null>(null);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactFindCatalog?.() || null; setCatalog(next ? { ...next } : null); };
    const off = runtime()?.bus?.on('pptx-find-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const dispatch = (action: string, values?: FindCatalog) => runtime()?.dlg?.reactFindAction?.(action, values);
  const update = (changes: Partial<FindCatalog>) => {
    const next = { ...catalog, ...changes };
    setCatalog(next);
    dispatch('update', next);
  };
  const close = () => dispatch('close');
  return <OfficeDialog title={t(catalog.replace ? 'Replace' : 'Find')} className="pptx-find-dialog" style={{ width: 'min(94vw, 480px)' }} modal={false} onRequestClose={close}>
    <section className="col">
      <label>{t('Fi&nd what:').replace('&', '')}<input autoFocus aria-label={t('Find what:')} type="text" value={catalog.what} onChange={(event) => update({ what: event.currentTarget.value })} /></label>
      {catalog.replace && <label>{t('Re&place with:').replace('&', '')}<input aria-label={t('Replace with:')} type="text" value={catalog.with} onChange={(event) => update({ with: event.currentTarget.value })} /></label>}
      <label><input type="checkbox" checked={catalog.mc} onChange={(event) => update({ mc: event.currentTarget.checked })} />{t('Matc&h case').replace('&', '')}</label>
      <label><input type="checkbox" checked={catalog.ww} onChange={(event) => update({ ww: event.currentTarget.checked })} />{t('Find &whole words only').replace('&', '')}</label>
    </section>
    <div className="vibeoffice-dialog__actions">
      <button className="primary" type="button" onClick={() => dispatch('find-next')}>{t('&Find Next').replace('&', '')}</button>
      {catalog.replace ? <><button type="button" onClick={() => dispatch('replace')}>{t('&Replace').replace('&', '')}</button><button type="button" onClick={() => dispatch('replace-all')}>{t('Replace &All').replace('&', '')}</button></> : <button type="button" onClick={() => dispatch('switch')}>{t('Re&place...').replace('&', '')}</button>}
      <button type="button" onClick={close}>{t('Close')}</button>
    </div>
  </OfficeDialog>;
}
