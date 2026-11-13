import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface SizeCatalog { title: string; label: string; value: string; min: number; max: number; kind: string }
interface SizeRuntime {
  dlg?: { reactSizeCatalog?(): SizeCatalog | null; reactSizeApply?(value: number): boolean; reactSizeClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}

function runtime(): SizeRuntime | undefined {
  return (window as Window & { L?: SizeRuntime }).L;
}

/** Shared React form for Excelx row-height, column-width, and standard-width dialogs. */
export function ExcelxSizeDialog() {
  const [catalog, setCatalog] = useState<SizeCatalog | null>(null);
  const [value, setValue] = useState('');
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactSizeCatalog?.() || null;
      setCatalog(next);
      if (next) setValue(next.value);
    };
    return runtime()?.bus?.on('excelx-size-dialog', refresh);
  }, []);
  if (!catalog) return null;
  const translate = (message: string, ...args: unknown[]) => runtime()?.t?.(message, ...args) || message;
  const label = translate(catalog.label).replace(/&(?=\S)/, '');
  const apply = () => {
    const parsed = Number.parseFloat(value.replace(/,/g, ''));
    runtime()?.dlg?.reactSizeApply?.(parsed);
  };
  const close = () => runtime()?.dlg?.reactSizeClose?.();
  const range = translate('Enter a number between {0} and {1}.', catalog.min, catalog.max);
  return <OfficeDialog title={translate(catalog.title)} className={`excelx-size-dialog excelx-size-dialog--${catalog.kind}`} style={{ width: 'min(92vw, 340px)' }} onRequestClose={close}>
    <form className="excelx-size-dialog__form" onSubmit={(event) => { event.preventDefault(); apply(); }}>
      <label>{label}<input autoFocus aria-label={label} type="text" value={value} onChange={(event) => setValue(event.currentTarget.value)} /></label>
      <p className="excelx-size-dialog__range">{range}</p>
      <div className="excelx-size-dialog__actions">
        <button type="submit" className="primary">{translate('OK')}</button>
        <button type="button" onClick={close}>{translate('Cancel')}</button>
      </div>
    </form>
  </OfficeDialog>;
}
