import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

interface SaveAsCatalog { name: string; type: string; types: Array<[string, string]>; note: string }
interface SaveAsRuntime {
  dlg?: { reactSaveAsCatalog?(): SaveAsCatalog | null; reactSaveAsAction?(action: 'save' | 'close', values?: { name: string; type: string }): Promise<unknown> | unknown };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

/** React-owned format/name selection; file generation and local/remote persistence stay in the presentation services. */
export function PptxSaveAsDialog() {
  const runtime = usePptxRuntime<SaveAsRuntime>();
  const [catalog, setCatalog] = useState<SaveAsCatalog | null>(null);
  const [name, setName] = useState('');
  const [type, setType] = useState('pptx');
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.dlg?.reactSaveAsCatalog?.() || null;
      setCatalog(next ? { ...next, types: [...next.types] } : null);
      if (next) { setName(next.name); setType(next.type); }
    };
    const off = runtime?.bus?.on('pptx-save-as-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const close = () => { void runtime?.dlg?.reactSaveAsAction?.('close'); };
  const save = () => { void runtime?.dlg?.reactSaveAsAction?.('save', { name, type }); };
  return <OfficeDialog title={t('Save As')} className="pptx-save-as-dialog" style={{ width: 'min(94vw, 480px)' }} onRequestClose={close}>
    <form className="col" onSubmit={(event) => { event.preventDefault(); save(); }}>
      <label>{t('File &name:').replace('&', '')}<input autoFocus required type="text" aria-label={t('File name')} value={name} onChange={(event) => setName(event.currentTarget.value)} /></label>
      <label>{t('Save as &type:').replace('&', '')}<select aria-label={t('Save as type')} value={type} onChange={(event) => setType(event.currentTarget.value)}>{catalog.types.map(([id, label]) => <option key={id} value={id}>{t(label)}</option>)}</select></label>
      <p className="tp-note">{t(catalog.note)}</p>
      <div className="vibeoffice-dialog__actions"><button type="submit" className="primary" disabled={!name.trim()}>{t('&Save').replace('&', '')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
    </form>
  </OfficeDialog>;
}
