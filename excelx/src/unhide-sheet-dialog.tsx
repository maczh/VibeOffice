import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface UnhideSheetOption { index: number; label: string }
interface UnhideSheetCatalog { options: UnhideSheetOption[]; selected: number | null }
interface UnhideSheetRuntime {
  dlg?: { reactUnhideCatalog?(): UnhideSheetCatalog | null; reactUnhideApply?(index: number): boolean; reactUnhideClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): UnhideSheetRuntime | undefined {
  return (window as Window & { L?: UnhideSheetRuntime }).L;
}

/** React UI for restoring a hidden worksheet. Sheet names remain document data, not translated UI. */
export function ExcelxUnhideSheetDialog() {
  const [catalog, setCatalog] = useState<UnhideSheetCatalog | null>(null);
  const [selected, setSelected] = useState<number | null>(null);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactUnhideCatalog?.() || null;
      setCatalog(next);
      setSelected(next?.selected ?? null);
    };
    const off = runtime()?.bus?.on('excelx-unhide-sheet-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactUnhideCatalog?.()) runtime()?.dlg?.reactUnhideClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactUnhideClose?.();
  const apply = () => { if (selected != null) runtime()?.dlg?.reactUnhideApply?.(selected); };
  return <OfficeDialog title={translate('Unhide')} className="excelx-unhide-sheet-dialog" style={{ width: 'min(92vw, 360px)' }} onRequestClose={close}>
    <div className="excelx-unhide-sheet-dialog__list">
      <div>{translate('&Unhide sheet:').replace('&', '')}</div>
      <div role="radiogroup" aria-label={translate('&Unhide sheet:').replace('&', '')}>
      {catalog.options.map((option) => <label key={option.index}><input type="radio" name="excelx-unhide-sheet" value={option.index} checked={selected === option.index} onChange={() => setSelected(option.index)} /><span translate="no">{option.label}</span></label>)}
      </div>
    </div>
    <div className="excelx-unhide-sheet-dialog__actions">
      <button type="button" className="primary" disabled={catalog.options.length === 0 || selected == null} onClick={apply}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
