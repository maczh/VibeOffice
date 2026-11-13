import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface TabColorOption { color: string; label: string }
interface TabColorCatalog { options: TabColorOption[]; selected: string | null }
interface TabColorRuntime {
  dlg?: { reactTabColorCatalog?(): TabColorCatalog | null; reactTabColorApply?(color: string | null): boolean; reactTabColorClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): TabColorRuntime | undefined {
  return (window as Window & { L?: TabColorRuntime }).L;
}

/** React UI for setting the selected worksheet tabs' standard colors. */
export function ExcelxTabColorDialog() {
  const [catalog, setCatalog] = useState<TabColorCatalog | null>(null);
  const [selected, setSelected] = useState<string | null>(null);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactTabColorCatalog?.() || null;
      setCatalog(next);
      setSelected(next?.selected ?? null);
    };
    const off = runtime()?.bus?.on('excelx-tab-color-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactTabColorCatalog?.()) runtime()?.dlg?.reactTabColorClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactTabColorClose?.();
  const apply = () => runtime()?.dlg?.reactTabColorApply?.(selected);
  return <OfficeDialog title={translate('Format Tab Color')} className="excelx-tab-color-dialog" style={{ width: 'min(92vw, 300px)' }} onRequestClose={close}>
    <fieldset className="excelx-tab-color-dialog__palette">
      <legend>{translate('Tab Color')}</legend>
      <button type="button" className="excelx-tab-color-dialog__none" aria-pressed={selected === null} onClick={() => setSelected(null)}>{translate('No Color')}</button>
      <div className="excelx-tab-color-dialog__swatches" role="radiogroup" aria-label={translate('Tab Color')}>
        {catalog.options.map(({ color, label }) => <button key={color} type="button" role="radio" aria-checked={selected === color} aria-label={label} title={label} style={{ backgroundColor: color }} onClick={() => setSelected(color)} />)}
      </div>
    </fieldset>
    <div className="excelx-tab-color-dialog__actions">
      <button type="button" className="primary" onClick={apply}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
