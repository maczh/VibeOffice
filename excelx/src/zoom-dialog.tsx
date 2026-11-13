import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface ZoomCatalog { zoom: number }
interface ZoomRuntime {
  dlg?: { reactZoomCatalog?(): ZoomCatalog | null; reactZoomClose?(): void };
  app?: { zoomToSelection?(): void };
  grid?: { setZoom?(value: number): void };
  ui?: { msg?(text: string, options?: { icon?: string }): Promise<unknown> };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): ZoomRuntime | undefined {
  return (window as Window & { L?: ZoomRuntime }).L;
}

const PRESETS = [200, 100, 75, 50, 25];

/** React-owned Excelx zoom dialog; the grid retains zoom and fit calculations. */
export function ExcelxZoomDialog() {
  const [catalog, setCatalog] = useState<ZoomCatalog | null>(null);
  const [choice, setChoice] = useState('100');
  const [custom, setCustom] = useState('100');
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactZoomCatalog?.() || null;
      setCatalog(next);
      if (next) {
        const zoom = next.zoom || 100;
        setChoice(PRESETS.includes(zoom) ? String(zoom) : 'custom');
        setCustom(`${zoom}%`);
      }
    };
    const off = runtime()?.bus?.on('excelx-zoom', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactZoomClose?.();
  const apply = () => {
    if (choice === 'fit') {
      runtime()?.app?.zoomToSelection?.();
      close();
      return;
    }
    const value = choice === 'custom' ? Number.parseInt(custom, 10) : Number(choice);
    if (!(value >= 10 && value <= 400)) {
      void runtime()?.ui?.msg?.('Enter a number between 10 and 400.', { icon: 'warn' });
      return;
    }
    runtime()?.grid?.setZoom?.(value);
    close();
  };
  return <OfficeDialog title={t('Zoom')} className="excelx-zoom-dialog" style={{ width: 'min(92vw, 320px)' }} onRequestClose={close}>
    <fieldset className="excelx-zoom-dialog__group">
      <legend>{t('Magnification')}</legend>
      {PRESETS.map((value) => <label key={value}><input type="radio" name="excelx-zoom" value={value} checked={choice === String(value)} onChange={() => setChoice(String(value))} />{value}%</label>)}
      <label><input type="radio" name="excelx-zoom" value="fit" checked={choice === 'fit'} onChange={() => setChoice('fit')} />{t('&Fit selection').replace('&', '')}</label>
      <label><input type="radio" name="excelx-zoom" value="custom" checked={choice === 'custom'} onChange={() => setChoice('custom')} />{t('&Custom:').replace('&', '')}<input aria-label={t('Zoom')} type="text" value={custom} onFocus={() => setChoice('custom')} onChange={(event) => { setChoice('custom'); setCustom(event.currentTarget.value); }} /></label>
    </fieldset>
    <div className="excelx-zoom-dialog__actions">
      <button type="button" className="primary" onClick={apply}>{t('OK')}</button>
      <button type="button" onClick={close}>{t('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
