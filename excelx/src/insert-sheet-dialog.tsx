import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState } from 'react';

interface InsertSheetOption { value: string; label: string }
interface InsertSheetCatalog { options: InsertSheetOption[] }
interface InsertSheetRuntime {
  dlg?: { reactInsertSheetCatalog?(): InsertSheetCatalog | null; reactInsertSheetApply?(value: string): boolean; reactInsertSheetClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): InsertSheetRuntime | undefined {
  return (window as Window & { L?: InsertSheetRuntime }).L;
}

/** React UI for Worksheet, Chart, and built-in template sheet insertion. */
export function ExcelxInsertSheetDialog() {
  const [catalog, setCatalog] = useState<InsertSheetCatalog | null>(null);
  const [selected, setSelected] = useState('ws');
  const [tab, setTab] = useState('General');
  const id = useId().replaceAll(':', '');
  const tabId = (key: 'general' | 'solutions') => `${id}-insert-tab-${key}`;
  const panelId = (key: 'general' | 'solutions') => `${id}-insert-panel-${key}`;
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactInsertSheetCatalog?.() || null;
      setCatalog(next);
      if (next) {
        setSelected('ws');
        setTab('General');
      }
    };
    const off = runtime()?.bus?.on('excelx-insert-sheet-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactInsertSheetCatalog?.()) runtime()?.dlg?.reactInsertSheetClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactInsertSheetClose?.();
  const apply = () => runtime()?.dlg?.reactInsertSheetApply?.(selected);
  return <OfficeDialog title={translate('Insert')} className="excelx-insert-sheet-dialog" style={{ width: 'min(92vw, 440px)' }} onRequestClose={close}>
    <div className="excelx-insert-sheet-dialog__tabs" role="tablist" aria-label={translate('Insert')}>
      {(['General', 'Spreadsheet Solutions'] as const).map((name, index) => {
        const key = index === 0 ? 'general' : 'solutions';
        return <button key={name} id={tabId(key)} type="button" role="tab" aria-controls={panelId(key)} aria-selected={tab === name} tabIndex={tab === name ? 0 : -1} onClick={() => setTab(name)} onKeyDown={(event) => {
          const nextIndex = event.key === 'ArrowRight' ? (index + 1) % 2
            : event.key === 'ArrowLeft' ? (index - 1 + 2) % 2
              : event.key === 'Home' ? 0
                : event.key === 'End' ? 1 : -1;
          if (nextIndex < 0) return;
          event.preventDefault();
          const nextName = nextIndex === 0 ? 'General' : 'Spreadsheet Solutions';
          setTab(nextName);
          event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[nextIndex]?.focus();
        }}>{translate(name)}</button>;
      })}
    </div>
    <div id={panelId('general')} aria-labelledby={tabId('general')} role="tabpanel" hidden={tab !== 'General'}>
      <div className="excelx-insert-sheet-dialog__options" role="radiogroup" aria-label={translate('General')}>
        {catalog.options.map((option) => <label key={option.value}><input type="radio" name="excelx-insert-sheet" value={option.value} checked={selected === option.value} onChange={() => setSelected(option.value)} />{option.label}</label>)}
      </div>
    </div>
    <div id={panelId('solutions')} aria-labelledby={tabId('solutions')} role="tabpanel" hidden={tab !== 'Spreadsheet Solutions'}>
      <p className="tp-note">{translate('Templates are listed on the General tab.')}</p>
    </div>
    <div className="excelx-insert-sheet-dialog__actions">
      <button type="button" className="primary" disabled={tab !== 'General'} onClick={apply}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
