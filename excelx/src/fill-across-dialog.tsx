import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { useExcelxRuntime } from './editor-host.js';

interface FillAcrossOption { value: string; label: string }
interface FillAcrossCatalog { options: FillAcrossOption[]; selected: string }
interface FillAcrossRuntime {
  dlg?: { reactFillAcrossCatalog?(): FillAcrossCatalog | null; reactFillAcrossApply?(value: string): boolean; reactFillAcrossClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

export function ExcelxFillAcrossDialog() {
  const runtime = useExcelxRuntime<FillAcrossRuntime>();
  const [catalog, setCatalog] = useState<FillAcrossCatalog | null>(null);
  const [selected, setSelected] = useState('all');
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.dlg?.reactFillAcrossCatalog?.() || null;
      setCatalog(next);
      if (next) setSelected(next.selected);
    };
    const off = runtime?.bus?.on('excelx-fill-across-dialog', refresh);
    return () => {
      off?.();
      if (runtime?.dlg?.reactFillAcrossCatalog?.()) runtime?.dlg?.reactFillAcrossClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime?.t?.(message) || message;
  const close = () => runtime?.dlg?.reactFillAcrossClose?.();
  const apply = () => runtime?.dlg?.reactFillAcrossApply?.(selected);
  return <OfficeDialog title={translate('Fill Across Worksheets')} className="excelx-fill-across-dialog" style={{ width: 'min(92vw, 300px)' }} onRequestClose={close}>
    <fieldset className="excelx-fill-across-dialog__choices">
      <legend>{translate('Fill')}</legend>
      {catalog.options.map(({ value, label }) => <label key={value}><input type="radio" name="excelx-fill-across" value={value} checked={selected === value} onChange={() => setSelected(value)} />{translate(label).replace('&', '')}</label>)}
    </fieldset>
    <div className="excelx-fill-across-dialog__actions">
      <button type="button" className="primary" onClick={apply}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
