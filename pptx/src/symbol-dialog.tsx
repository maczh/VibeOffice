import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface SymbolCatalog {
  font: string;
  fonts: string[];
  subset: number;
  subsets: string[];
  selected: string | null;
  code: string;
  recent: string[];
  glyphs: string[];
}
interface SymbolRuntime {
  dlg?: { reactSymbolCatalog?(): SymbolCatalog | null; reactSymbolAction?(action: string, value?: unknown): SymbolCatalog | null };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
  fontStack?(font: string): string;
}
function runtime(): SymbolRuntime | undefined { return (window as Window & { L?: SymbolRuntime }).L; }

/** React-owned Unicode symbol picker; character insertion and custom-bullet callbacks remain engine-backed. */
export function PptxSymbolDialog() {
  const [catalog, setCatalog] = useState<SymbolCatalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactSymbolCatalog?.() || null);
    const off = runtime()?.bus?.on('pptx-symbol-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const act = (action: string, value?: unknown) => {
    const next = runtime()?.dlg?.reactSymbolAction?.(action, value) || null;
    setCatalog(next ? { ...next, glyphs: [...next.glyphs], recent: [...next.recent] } : null);
  };
  const close = () => act('close');
  return <OfficeDialog title={t('Symbol')} className="pptx-symbol-dialog" style={{ width: 'min(94vw, 510px)' }} onRequestClose={close}>
    <div className="row" style={{ gap: 8 }}>
      <label>{t('&Font:').replace('&', '')}<select aria-label={t('&Font:').replace('&', '')} value={catalog.font} onChange={(event) => act('font', event.currentTarget.value)}>
        {catalog.fonts.map((font) => <option key={font} value={font}>{font === 'normal' ? t('(normal text)') : font}</option>)}
      </select></label>
      <label>{t('Su&bset:').replace('&', '')}<select aria-label={t('Su&bset:').replace('&', '')} value={catalog.subset} onChange={(event) => act('subset', Number(event.currentTarget.value))}>
        {catalog.subsets.map((subset, index) => <option key={`${index}:${subset}`} value={index}>{subset}</option>)}
      </select></label>
    </div>
    <div className="sym-grid" translate="no" aria-label={catalog.subsets[catalog.subset]} style={{ fontFamily: catalog.font === 'normal' ? undefined : runtime()?.fontStack?.(catalog.font) }}>
      {catalog.glyphs.map((glyph) => <button key={glyph} type="button" className={catalog.selected === glyph ? 'on' : ''} aria-label={`U+${glyph.codePointAt(0)?.toString(16).toUpperCase()}`} onClick={() => act('select', glyph)} onDoubleClick={() => { act('select', glyph); act('insert'); }}>{glyph}</button>)}
    </div>
    <label>{t('Recently used symbols:')}</label>
    <div className="row" style={{ gap: 2, flexWrap: 'wrap' }}>
      {catalog.recent.map((glyph, index) => <button key={`${glyph}-${index}`} type="button" className="btn small" style={{ minWidth: 24, padding: 0 }} aria-label={glyph} onClick={() => act('select', glyph)}>{glyph}</button>)}
    </div>
    <label>{t('&Character code:').replace('&', '')}<input type="text" aria-label={t('&Character code:').replace('&', '')} value={catalog.code} readOnly style={{ width: 80 }} /></label>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" disabled={!catalog.selected} onClick={() => act('insert')}>{t('&Insert').replace('&', '')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
