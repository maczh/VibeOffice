import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface ZoomCatalog { value: number | 'fit' }
interface ZoomRuntime {
  dlg?: { reactZoomCatalog?(): ZoomCatalog | null; reactZoomApply?(value: number | 'fit'): void; reactZoomClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): ZoomRuntime | undefined { return (window as Window & { L?: ZoomRuntime }).L; }
const PRESETS = [400, 200, 100, 66, 50, 33];

/** React-owned zoom choices; slide scaling and fit calculations remain in the presentation engine. */
export function PptxZoomDialog() {
  const [catalog, setCatalog] = useState<ZoomCatalog | null>(null);
  const [choice, setChoice] = useState('100');
  const [custom, setCustom] = useState('100');
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactZoomCatalog?.() || null;
      setCatalog(next);
      if (!next) return;
      setChoice(next.value === 'fit' ? 'fit' : PRESETS.includes(next.value) ? String(next.value) : 'custom');
      setCustom(next.value === 'fit' ? '100' : String(next.value));
    };
    const off = runtime()?.bus?.on('pptx-zoom-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactZoomClose?.();
  const apply = () => {
    if (choice === 'fit') { runtime()?.dlg?.reactZoomApply?.('fit'); return; }
    const value = Number.parseInt(choice === 'custom' ? custom : choice, 10);
    if (value >= 10 && value <= 400) runtime()?.dlg?.reactZoomApply?.(value);
  };
  return <OfficeDialog title={t('Zoom')} className="pptx-zoom-dialog" style={{ width: 'min(92vw, 320px)' }} onRequestClose={close}>
    <fieldset>
      <legend>{t('Zoom to')}</legend>
      {(['fit', ...PRESETS.map(String)]).map((value) => <label key={value} style={{ display: 'block' }}><input type="radio" name="pptx-zoom" value={value} checked={choice === value} onChange={() => setChoice(value)} />{value === 'fit' ? t('Fit') : `${value}%`}</label>)}
      <label style={{ display: 'block' }}><input type="radio" name="pptx-zoom" value="custom" checked={choice === 'custom'} onChange={() => setChoice('custom')} />{t('&Percent:').replace('&', '')}<input aria-label={t('Percent')} type="number" min={10} max={400} value={custom} onFocus={() => setChoice('custom')} onChange={(event) => { setChoice('custom'); setCustom(event.currentTarget.value); }} />%</label>
    </fieldset>
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={apply}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
