import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface CreateListCatalog { reference: string; header: boolean }
interface CreateListRuntime {
  dlg?: { reactCreateListCatalog?(): CreateListCatalog | null; reactCreateListSelection?(): string; reactCreateListApply?(reference: string, header: boolean): boolean; reactCreateListClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): CreateListRuntime | undefined {
  return (window as Window & { L?: CreateListRuntime }).L;
}

export function ExcelxCreateListDialog() {
  const [catalog, setCatalog] = useState<CreateListCatalog | null>(null);
  const [reference, setReference] = useState('');
  const [header, setHeader] = useState(false);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactCreateListCatalog?.() || null;
      setCatalog(next);
      if (next) { setReference(next.reference); setHeader(next.header); }
    };
    const off = runtime()?.bus?.on('excelx-create-list-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactCreateListCatalog?.()) runtime()?.dlg?.reactCreateListClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactCreateListClose?.();
  const apply = () => runtime()?.dlg?.reactCreateListApply?.(reference, header);
  const pickSelection = () => setReference(runtime()?.dlg?.reactCreateListSelection?.() || reference);
  return <OfficeDialog title={translate('Create List')} className="excelx-create-list-dialog" style={{ width: 'min(92vw, 380px)' }} onRequestClose={close}>
    <div className="excelx-create-list-dialog__body">
      <div>{translate('Where is the data for your list?')}</div>
      <div className="excelx-create-list-dialog__reference">
        <input aria-label={translate('&Range:').replace('&', '')} value={reference} onChange={(event) => { const value = event.currentTarget.value; setReference(value); }} />
        <button type="button" aria-label={translate('Select range')} title={translate('Use the current selection')} onClick={pickSelection}>…</button>
      </div>
      <label><input type="checkbox" checked={header} onChange={(event) => { const checked = event.currentTarget.checked; setHeader(checked); }} />{translate('&My list has headers').replace('&', '')}</label>
    </div>
    <div className="excelx-create-list-dialog__actions">
      <button type="button" className="primary" onClick={apply}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
