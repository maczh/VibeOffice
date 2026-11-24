import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

interface PasteSpecialCatalog { label: string; kind: string }
interface PasteSpecialRuntime {
  app?: { reactPasteSpecialCatalog?(): PasteSpecialCatalog | null; reactPasteSpecialAction?(action: string): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}

/** React-owned Paste Special choices; clipboard formats and insertion remain presentation-engine services. */
export function PptxPasteSpecialDialog() {
  const runtime = usePptxRuntime<PasteSpecialRuntime>();
  const [catalog, setCatalog] = useState<PasteSpecialCatalog | null>(null);
  const [mode, setMode] = useState<'native' | 'text'>('native');
  useEffect(() => {
    const refresh = () => { setCatalog(runtime?.app?.reactPasteSpecialCatalog?.() || null); setMode('native'); };
    const off = runtime?.bus?.on('pptx-paste-special-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string, ...args: unknown[]) => runtime?.t?.(message, ...args) || message;
  const close = () => runtime?.app?.reactPasteSpecialAction?.('close');
  return <OfficeDialog title={t('Paste Special')} className="pptx-paste-special-dialog" style={{ width: 'min(92vw, 360px)' }} onRequestClose={close}>
    <div className="vibeoffice-dialog__body">
      <div>{t('Source: {0}', catalog.label)}</div>
      <label><input type="radio" name="pptx-paste-special" value="native" checked={mode === 'native'} onChange={() => setMode('native')} />{t(catalog.kind === 'slides' ? '&Paste as slides' : '&Paste as PPTX objects').replace('&', '')}</label>
      <label><input type="radio" name="pptx-paste-special" value="text" checked={mode === 'text'} onChange={() => setMode('text')} />{t('Paste as &Unformatted Text').replace('&', '')}</label>
    </div>
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={() => runtime?.app?.reactPasteSpecialAction?.(mode)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
