import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useRef, useState } from 'react';

type FontStyle = 'regular' | 'italic' | 'bold' | 'bolditalic';
type ScriptPosition = 'super' | 'sub' | null;
interface FontCatalog { font: string; fonts: string[]; styles: Array<[FontStyle, string]>; style: FontStyle; size: number; sizes: number[]; underline: boolean; shadow: boolean; emboss: boolean; script: ScriptPosition; offset: number; color: string | null }
interface FontApply { font: string; size: number; bold: boolean; italic: boolean; underline: boolean; shadow: boolean; emboss: boolean; script: ScriptPosition; offset: number; color: string | null }
interface FontRuntime {
  dlg?: { reactFontCatalog?(): FontCatalog | null; reactFontOpenColor?(anchor: HTMLElement): boolean; reactFontResolveColor?(color: string): string; reactFontAction?(action: string, values?: FontApply): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
  fontStack?(font: string): string;
}
function runtime(): FontRuntime | undefined { return (window as Window & { L?: FontRuntime }).L; }

/** React-owned Font form and preview; run formatting and WordArt updates stay in the presentation formatter. */
export function PptxFontDialog() {
  const [catalog, setCatalog] = useState<FontCatalog | null>(null);
  const colorButton = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactFontCatalog?.() || null; setCatalog(next ? { ...next, fonts: [...next.fonts], sizes: [...next.sizes], styles: [...next.styles] } : null); };
    const off = runtime()?.bus?.on('pptx-font-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const update = (changes: Partial<FontCatalog>) => setCatalog({ ...catalog, ...changes });
  const close = () => runtime()?.dlg?.reactFontAction?.('close');
  const color = catalog.color ? (catalog.color.startsWith('#') ? catalog.color : runtime()?.dlg?.reactFontResolveColor?.(catalog.color) || catalog.color) : '#000000';
  const bold = catalog.style === 'bold' || catalog.style === 'bolditalic';
  const italic = catalog.style === 'italic' || catalog.style === 'bolditalic';
  const previewStyle = { fontFamily: runtime()?.fontStack?.(catalog.font), fontSize: `${Math.min(40, Math.max(10, Number(catalog.size) || 18))}px`, fontWeight: bold ? 700 : 400, fontStyle: italic ? 'italic' as const : 'normal' as const, textDecoration: catalog.underline ? 'underline' : 'none', textShadow: catalog.shadow ? '2px 2px 0 rgba(0,0,0,.35)' : catalog.emboss ? '-1px -1px 0 #fff, 1px 1px 0 #666' : 'none', color };
  const apply = () => runtime()?.dlg?.reactFontAction?.('apply', { font: catalog.font, size: Number(catalog.size), bold, italic, underline: catalog.underline, shadow: catalog.shadow, emboss: catalog.emboss, script: catalog.script, offset: Number(catalog.offset), color: catalog.color });
  return <OfficeDialog title={t('Font')} className="pptx-font-dialog" style={{ width: 'min(96vw, 620px)' }} onRequestClose={close}>
    <div className="row" style={{ alignItems: 'start', gap: 12 }}>
      <label>{t('&Font:').replace('&', '')}<input aria-label={t('&Font:').replace('&', '')} value={catalog.font} onChange={(event) => update({ font: event.currentTarget.value })} /><select aria-label={t('Font list')} size={7} style={{ width: 190 }} value={catalog.font} onChange={(event) => update({ font: event.currentTarget.value })}>{catalog.fonts.map((font) => <option key={font} value={font}>{font}</option>)}</select></label>
      <label>{t('F&ont style:').replace('&', '')}<select aria-label={t('Font style')} size={4} value={catalog.style} onChange={(event) => update({ style: event.currentTarget.value as FontStyle })}>{catalog.styles.map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></label>
      <label>{t('&Size:').replace('&', '')}<input aria-label={t('&Size:').replace('&', '')} type="number" min="1" max="4000" step="0.5" value={catalog.size} onChange={(event) => update({ size: Number(event.currentTarget.value) })} /><select aria-label={t('Font size list')} size={4} value={String(catalog.size)} onChange={(event) => update({ size: Number(event.currentTarget.value) })}>{catalog.sizes.map((size) => <option key={size} value={size}>{size}</option>)}</select></label>
    </div>
    <div className="row" style={{ alignItems: 'start', gap: 20 }}>
      <fieldset><legend>{t('Effects')}</legend>
        <label><input type="checkbox" checked={catalog.underline} onChange={(event) => update({ underline: event.currentTarget.checked })} />{t('&Underline').replace('&', '')}</label>
        <label><input type="checkbox" checked={catalog.shadow} onChange={(event) => update({ shadow: event.currentTarget.checked })} />{t('&Shadow').replace('&', '')}</label>
        <label><input type="checkbox" checked={catalog.emboss} onChange={(event) => update({ emboss: event.currentTarget.checked })} />{t('&Emboss').replace('&', '')}</label>
      </fieldset>
      <fieldset><legend>{t('Script')}</legend>
        <label><input type="radio" name="pptx-font-script" checked={catalog.script === null} onChange={() => update({ script: null })} />{t('Normal')}</label>
        <label><input type="radio" name="pptx-font-script" checked={catalog.script === 'super'} onChange={() => update({ script: 'super' })} />{t('Su&perscript').replace('&', '')}</label>
        <label><input type="radio" name="pptx-font-script" checked={catalog.script === 'sub'} onChange={() => update({ script: 'sub' })} />{t('Su&bscript').replace('&', '')}</label>
        <label>{t('Offset:')} <input aria-label={t('Offset')} type="number" min="1" max="100" step="1" value={catalog.offset} onChange={(event) => update({ offset: Number(event.currentTarget.value) })} />%</label>
      </fieldset>
      <label>{t('&Color:').replace('&', '')}<button ref={colorButton} type="button" className="btn small" aria-label={t('Font color')} onClick={() => { if (colorButton.current) runtime()?.dlg?.reactFontOpenColor?.(colorButton.current); }}><span style={{ display: 'inline-block', width: 60, height: 14, background: color, border: '1px solid #555', verticalAlign: 'middle' }} /></button></label>
    </div>
    <div className="preview-box font-preview" style={previewStyle} translate="no">AaBbYyZz</div>
    <p className="tp-note">{t('This is a TrueType font. The same font will be used on both your printer and your screen.')}</p>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={apply}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
