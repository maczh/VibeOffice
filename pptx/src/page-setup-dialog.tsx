import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface SlideSize { key: string; name: string; width: number; height: number }
interface PageSetupCatalog { sizes: SlideSize[]; preset: string; width: number; height: number; firstNum: number; portrait: boolean; asDefault: boolean }
interface PageSetupRuntime {
  dlg?: { reactPageSetupCatalog?(): PageSetupCatalog | null; reactPageSetupAction?(action: string, values?: Omit<PageSetupCatalog, 'sizes'>): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): PageSetupRuntime | undefined { return (window as Window & { L?: PageSetupRuntime }).L; }

/** React-owned slide size/orientation controls; resizing, defaults and slide-number metadata remain engine-owned. */
export function PptxPageSetupDialog() {
  const [catalog, setCatalog] = useState<PageSetupCatalog | null>(null);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactPageSetupCatalog?.() || null; setCatalog(next ? { ...next, sizes: [...next.sizes] } : null); };
    const off = runtime()?.bus?.on('pptx-page-setup-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const update = (changes: Partial<PageSetupCatalog>) => setCatalog({ ...catalog, ...changes });
  const orientation = (portrait: boolean) => {
    if (catalog.portrait === portrait) return;
    update({ portrait, width: catalog.height, height: catalog.width });
  };
  const close = () => runtime()?.dlg?.reactPageSetupAction?.('close');
  const apply = () => runtime()?.dlg?.reactPageSetupAction?.('apply', { preset: catalog.preset, width: catalog.width, height: catalog.height, firstNum: catalog.firstNum, portrait: catalog.portrait, asDefault: catalog.asDefault });
  return <OfficeDialog title={t('Page Setup')} className="pptx-page-setup-dialog" style={{ width: 'min(94vw, 560px)' }} onRequestClose={close}>
    <div className="row" style={{ gap: 16 }}>
      <section className="col" style={{ flex: '1 1 300px' }}>
        <label>{t('&Slides sized for:').replace('&', '')}<select aria-label={t('Slides sized for')} value={catalog.preset} onChange={(event) => {
          const preset = event.currentTarget.value;
          const size = catalog.sizes.find((item) => item.key === preset);
          update({ preset, ...(size && preset !== 'custom' ? { width: size.width, height: size.height } : {}) });
        }}>{catalog.sizes.map((size) => <option key={size.key} value={size.key}>{t(size.name)}</option>)}</select></label>
        <label>{t('&Width:').replace('&', '')}<input aria-label={t('Width:')} type="number" min={1} max={56} step={0.1} value={catalog.width} onChange={(event) => update({ width: Number(event.currentTarget.value), preset: 'custom' })} /></label>
        <label>{t('H&eight:').replace('&', '')}<input aria-label={t('Height:')} type="number" min={1} max={56} step={0.1} value={catalog.height} onChange={(event) => update({ height: Number(event.currentTarget.value), preset: 'custom' })} /></label>
        <label>{t('&Number slides from:').replace('&', '')}<input aria-label={t('Number slides from:')} type="number" min={0} max={9999} step={1} value={catalog.firstNum} onChange={(event) => update({ firstNum: Number(event.currentTarget.value) })} /></label>
        <label><input type="checkbox" checked={catalog.asDefault} onChange={(event) => update({ asDefault: event.currentTarget.checked })} />{t('&Use for new presentations').replace('&', '')}</label>
      </section>
      <fieldset className="col"><legend>{t('Orientation')}</legend><strong>{t('Slides')}</strong>
        <label><input type="radio" name="pptx-page-orientation" checked={catalog.portrait} onChange={() => orientation(true)} />{t('&Portrait').replace('&', '')}</label>
        <label><input type="radio" name="pptx-page-orientation" checked={!catalog.portrait} onChange={() => orientation(false)} />{t('&Landscape').replace('&', '')}</label>
      </fieldset>
    </div>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={apply}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
