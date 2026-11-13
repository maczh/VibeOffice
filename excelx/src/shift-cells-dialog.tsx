import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface ShiftOption { value: string; label: string }
interface ShiftCatalog { title: string; selected: string; options: ShiftOption[] }
interface ShiftRuntime {
  dlg?: { reactShiftCatalog?(): ShiftCatalog | null; reactShiftApply?(key: string): boolean; reactShiftClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): ShiftRuntime | undefined {
  return (window as Window & { L?: ShiftRuntime }).L;
}

/** React-owned Insert/Delete Cells choice dialog; cell movement remains engine-backed. */
export function ExcelxShiftCellsDialog() {
  const [catalog, setCatalog] = useState<ShiftCatalog | null>(null);
  const [selected, setSelected] = useState('');
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactShiftCatalog?.() || null;
      setCatalog(next);
      if (next) setSelected(next.selected);
    };
    const off = runtime()?.bus?.on('excelx-shift-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactShiftCatalog?.()) runtime()?.dlg?.reactShiftClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactShiftClose?.();
  const commit = () => runtime()?.dlg?.reactShiftApply?.(selected);
  return <OfficeDialog title={translate(catalog.title)} className="excelx-shift-cells-dialog" style={{ width: 'min(92vw, 340px)' }} onRequestClose={close}>
    <fieldset className="excelx-shift-cells-dialog__options">
      <legend>{translate(catalog.title.replace(/\.\.\.$/, ''))}</legend>
      {catalog.options.map((option) => <label key={option.value}><input type="radio" name="excelx-shift-cells" value={option.value} checked={selected === option.value} onChange={() => setSelected(option.value)} />{translate(option.label).replace(/&(?=\S)/g, '')}</label>)}
    </fieldset>
    <div className="excelx-shift-cells-dialog__actions">
      <button type="button" className="primary" onClick={commit}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
