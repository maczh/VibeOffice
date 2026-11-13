import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface PasteChoice { value: string; label: string }
interface PasteSpecialState { what: string; op: string; skipBlanks: boolean; transpose: boolean }
interface PasteSpecialCatalog { what: PasteChoice[]; ops: PasteChoice[]; state: PasteSpecialState }
interface PasteSpecialRuntime {
  dlg?: { reactPasteSpecialCatalog?(): PasteSpecialCatalog | null; reactPasteSpecialApply?(state: PasteSpecialState, link?: boolean): boolean; reactPasteSpecialClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): PasteSpecialRuntime | undefined {
  return (window as Window & { L?: PasteSpecialRuntime }).L;
}

export function ExcelxPasteSpecialDialog() {
  const [catalog, setCatalog] = useState<PasteSpecialCatalog | null>(null);
  const [state, setState] = useState<PasteSpecialState>({ what: 'all', op: 'none', skipBlanks: false, transpose: false });
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactPasteSpecialCatalog?.() || null;
      setCatalog(next);
      if (next) setState(next.state);
    };
    const off = runtime()?.bus?.on('excelx-paste-special-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactPasteSpecialCatalog?.()) runtime()?.dlg?.reactPasteSpecialClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactPasteSpecialClose?.();
  const apply = (link = false) => runtime()?.dlg?.reactPasteSpecialApply?.(state, link);
  const choices = (items: PasteChoice[], group: 'what' | 'op') => <div className="excelx-paste-special-dialog__choices" role="radiogroup" aria-label={translate(group === 'what' ? 'Paste' : 'Operation')}>
    {items.map(({ value, label }) => <label key={value}><input type="radio" name={`excelx-paste-special-${group}`} value={value} checked={state[group] === value} onChange={() => setState((previous) => ({ ...previous, [group]: value }))} />{translate(label).replace('&', '')}</label>)}
  </div>;
  return <OfficeDialog title={translate('Paste Special')} className="excelx-paste-special-dialog" style={{ width: 'min(92vw, 460px)' }} onRequestClose={close}>
    <div className="excelx-paste-special-dialog__groups">
      <fieldset><legend>{translate('Paste')}</legend>{choices(catalog.what, 'what')}</fieldset>
      <fieldset><legend>{translate('Operation')}</legend>{choices(catalog.ops, 'op')}</fieldset>
    </div>
    <div className="excelx-paste-special-dialog__toggles">
      <label><input type="checkbox" checked={state.skipBlanks} onChange={(event) => { const checked = event.currentTarget.checked; setState((previous) => ({ ...previous, skipBlanks: checked })); }} />{translate('Skip &blanks').replace('&', '')}</label>
      <label><input type="checkbox" checked={state.transpose} onChange={(event) => { const checked = event.currentTarget.checked; setState((previous) => ({ ...previous, transpose: checked })); }} />{translate('Transpos&e').replace('&', '')}</label>
    </div>
    <div className="excelx-paste-special-dialog__actions">
      <button type="button" onClick={() => apply(true)}>{translate('Paste &Link').replace('&', '')}</button>
      <button type="button" className="primary" onClick={() => apply()}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
