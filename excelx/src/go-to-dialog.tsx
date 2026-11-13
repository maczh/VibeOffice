import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface GoToCatalog { options: string[] }
interface GoToRuntime {
  dlg?: { reactGoToCatalog?(): GoToCatalog | null; reactGoToApply?(reference: string): boolean; reactGoToSpecial?(): void; reactGoToClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): GoToRuntime | undefined {
  return (window as Window & { L?: GoToRuntime }).L;
}

/** React UI for navigating to a cell reference, defined name, or table. */
export function ExcelxGoToDialog() {
  const [catalog, setCatalog] = useState<GoToCatalog | null>(null);
  const [reference, setReference] = useState('');
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactGoToCatalog?.() || null;
      setCatalog(next);
      if (next) setReference('');
    };
    const off = runtime()?.bus?.on('excelx-go-to-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactGoToCatalog?.()) runtime()?.dlg?.reactGoToClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactGoToClose?.();
  const apply = () => runtime()?.dlg?.reactGoToApply?.(reference);
  const select = (value: string) => { setReference(value); runtime()?.dlg?.reactGoToApply?.(value); };
  return <OfficeDialog title={translate('Go To')} className="excelx-go-to-dialog" style={{ width: 'min(92vw, 360px)' }} onRequestClose={close}>
    <div className="excelx-go-to-dialog__form">
      <label>{translate('Go to:')}<select size={Math.min(7, Math.max(2, catalog.options.length))} aria-label={translate('Go to:')} value="" onChange={(event) => setReference(event.currentTarget.value)} onDoubleClick={(event) => select(event.currentTarget.value)}>
        {catalog.options.map((option) => <option key={option} value={option} translate="no">{option}</option>)}
      </select></label>
      <label>{translate('&Reference:').replace('&', '')}<input autoFocus aria-label={translate('&Reference:').replace('&', '')} value={reference} onChange={(event) => setReference(event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); apply(); } }} /></label>
    </div>
    <div className="excelx-go-to-dialog__actions">
      <button type="button" onClick={() => runtime()?.dlg?.reactGoToSpecial?.()}>{translate('&Special...').replace('&', '')}</button>
      <button type="button" className="primary" onClick={apply}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
