import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

interface WordArtTextCatalog { fonts: string[]; sizes: number[]; font: string; size: number; b: boolean; i: boolean; text: string }
interface WordArtTextRuntime {
  dlg?: { reactWordartTextCatalog?(): WordArtTextCatalog | null; reactWordartTextAction?(action: string, values?: Omit<WordArtTextCatalog, 'fonts' | 'sizes'>): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
  fontStack?(font: string): string;
}

/** React-owned LettersArt text editor; applying style/text continues through the existing presentation callback. */
export function PptxWordArtTextDialog() {
  const runtime = usePptxRuntime<WordArtTextRuntime>();
  const [catalog, setCatalog] = useState<WordArtTextCatalog | null>(null);
  useEffect(() => {
    const refresh = () => { const next = runtime?.dlg?.reactWordartTextCatalog?.() || null; setCatalog(next ? { ...next, fonts: [...next.fonts], sizes: [...next.sizes] } : null); };
    const off = runtime?.bus?.on('pptx-wordart-text-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const update = (changes: Partial<WordArtTextCatalog>) => setCatalog({ ...catalog, ...changes });
  const close = () => runtime?.dlg?.reactWordartTextAction?.('close');
  const apply = () => runtime?.dlg?.reactWordartTextAction?.('apply', { text: catalog.text, font: catalog.font, size: catalog.size, b: catalog.b, i: catalog.i });
  return <OfficeDialog title={t('Edit LettersArt Text')} className="pptx-wordart-text-dialog" style={{ width: 'min(94vw, 560px)' }} onRequestClose={close}>
    <section className="col">
      <div className="row" style={{ gap: 8, alignItems: 'center' }}>
        <label>{t('&Font:').replace('&', '')}<select aria-label={t('Font:')} value={catalog.font} onChange={(event) => update({ font: event.currentTarget.value })}>{catalog.fonts.map((font) => <option key={font} value={font}>{font}</option>)}</select></label>
        <label>{t('&Size:').replace('&', '')}<select aria-label={t('Size:')} value={catalog.size} onChange={(event) => update({ size: Number(event.currentTarget.value) })}>{catalog.sizes.map((size) => <option key={size} value={size}>{size}</option>)}</select></label>
        <label><input type="checkbox" aria-label={t('Bold')} checked={catalog.b} onChange={(event) => update({ b: event.currentTarget.checked })} />{t('&B').replace('&', '')}</label>
        <label><input type="checkbox" aria-label={t('Italic')} checked={catalog.i} onChange={(event) => update({ i: event.currentTarget.checked })} />{t('&I').replace('&', '')}</label>
      </div>
      <label>{t('&Text:').replace('&', '')}<textarea aria-label={t('Text:')} rows={5} value={catalog.text} onChange={(event) => update({ text: event.currentTarget.value })} /></label>
      <div className="font-preview" aria-label={t('Preview')} style={{ minHeight: 38, fontFamily: runtime?.fontStack?.(catalog.font) || catalog.font, fontSize: 18, fontWeight: catalog.b ? 700 : 400, fontStyle: catalog.i ? 'italic' : 'normal', overflowWrap: 'anywhere' }}>{catalog.text}</div>
    </section>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={apply}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
