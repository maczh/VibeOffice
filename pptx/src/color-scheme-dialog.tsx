import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

interface SchemeColor { key: string; label: string; color: string; hex: string }
interface PaletteColor { value: string; label: string; hex: string }
interface SchemeCatalog { slots: SchemeColor[]; theme: PaletteColor[]; standard: PaletteColor[] }
interface SchemeRuntime {
  app?: { reactColorSchemeCatalog?(): SchemeCatalog | null; reactColorSchemeAction?(action: string, key?: string, value?: string): void };
  ui?: { moreColors(initial: string | null, choose: (value: string | null) => void): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

/** React-owned theme-color editor; palette selection and scheme mutation use the existing shared color and slide services. */
export function PptxColorSchemeDialog() {
  const runtime = usePptxRuntime<SchemeRuntime>();
  const [catalog, setCatalog] = useState<SchemeCatalog | null>(null);
  const [openKey, setOpenKey] = useState<string | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime?.app?.reactColorSchemeCatalog?.() || null);
    const off = runtime?.bus?.on('pptx-edit-color-scheme', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const close = () => runtime?.app?.reactColorSchemeAction?.('close');
  return <OfficeDialog title={t('Edit Color Scheme')} className="pptx-color-scheme-dialog" style={{ width: 'min(92vw, 380px)' }} onRequestClose={close}>
    <div className="vibeoffice-dialog__body">
      {catalog.slots.map((color) => <div className="field wide" key={color.key}>
        <label>{color.label}</label>
        <div className="row" style={{ gap: 6 }}>
          <span aria-hidden="true" style={{ background: color.hex, display: 'inline-block', width: 42, height: 18, border: '1px solid #888' }} />
          <span translate="no">{color.color}</span>
          <button type="button" aria-expanded={openKey === color.key} onClick={() => setOpenKey(openKey === color.key ? null : color.key)}>{t('Change Colors')}</button>
        </div>
        {openKey === color.key && <div className="col" aria-label={t('Colors')}>
          <div className="cp-row">{catalog.theme.map((choice) => <button key={choice.value} type="button" className="cp-sw" data-color-value={choice.value} style={{ background: choice.hex }} aria-label={choice.label} title={choice.label} onClick={() => runtime?.app?.reactColorSchemeAction?.('color', color.key, choice.value)} />)}</div>
          <div className="cp-grid">{catalog.standard.map((choice) => <button key={choice.value} type="button" className="cp-sw" data-color-value={choice.value} style={{ background: choice.hex }} aria-label={choice.label} title={choice.label} onClick={() => runtime?.app?.reactColorSchemeAction?.('color', color.key, choice.value)} />)}</div>
          <button type="button" onClick={() => runtime?.ui?.moreColors(color.hex, (value) => { if (value) runtime?.app?.reactColorSchemeAction?.('color', color.key, value); })}>{t('More Colors...')}</button>
        </div>}
      </div>)}
    </div>
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={() => runtime?.app?.reactColorSchemeAction?.('apply')}>{t('&Apply').replace('&', '')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
