import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Catalog = { mode: 'protect' | 'confirm' | 'unprotect'; passwordRequired?: boolean; allow?: Record<string, boolean> };
type Runtime = { dlg?: { reactProtectionCatalog?(): Catalog | null; reactProtectionAction?(action: string, value?: unknown): Promise<boolean> | boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }
const PERMISSIONS: [string, string][] = [
  ['selectLockedCells', 'Select locked cells'], ['selectUnlockedCells', 'Select unlocked cells'], ['formatCells', 'Format cells'], ['formatColumns', 'Format columns'], ['formatRows', 'Format rows'], ['insertColumns', 'Insert columns'], ['insertRows', 'Insert rows'], ['insertHyperlinks', 'Insert hyperlinks'], ['deleteColumns', 'Delete columns'], ['deleteRows', 'Delete rows'], ['sort', 'Sort'], ['autoFilter', 'Use AutoFilter'], ['pivotTables', 'Use PivotTable reports'], ['objects', 'Edit objects'], ['scenarios', 'Edit scenarios'],
];

/** React-owned worksheet-protection workflow; password hashes and protection edits remain in the engine adapter. */
export function ExcelxProtectionDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [allow, setAllow] = useState<Record<string, boolean>>({});
  const [password, setPassword] = useState('');
  const [confirmation, setConfirmation] = useState('');
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactProtectionCatalog?.() || null;
      setCatalog(next);
      if (next?.mode === 'protect') setAllow(next.allow || {});
      setPassword(''); setConfirmation('');
    };
    const off = runtime()?.bus?.on('excelx-protection-dialog', refresh);
    refresh();
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => { void runtime()?.dlg?.reactProtectionAction?.('close'); };
  const title = catalog.mode === 'protect' ? 'Protect Sheet' : catalog.mode === 'confirm' ? 'Confirm Password' : 'Unprotect Sheet';
  const submit = async () => {
    if (catalog.mode === 'unprotect') await runtime()?.dlg?.reactProtectionAction?.('submit', password);
    else if (catalog.mode === 'protect') await runtime()?.dlg?.reactProtectionAction?.('submit', { allow, password });
    else await runtime()?.dlg?.reactProtectionAction?.('confirm', confirmation);
  };
  return <OfficeDialog title={t(title)} className="excelx-protection-dialog" style={{ width: 'min(94vw, 420px)' }} onRequestClose={close}>
    <div className="excelx-protection-dialog__body">
      {catalog.mode === 'protect' && <>
        <p>{t('Protect worksheet and contents of locked cells')}</p>
        <label>{t('Password to unprotect sheet:')}<input type="password" autoComplete="new-password" value={password} onChange={(event) => { const value = event.currentTarget.value; setPassword(value); }} /></label>
        <fieldset><legend>{t('Allow all users of this worksheet to:')}</legend>{PERMISSIONS.map(([key, label]) => <label key={key}><input type="checkbox" checked={!!allow[key]} onChange={(event) => { const checked = event.currentTarget.checked; setAllow((old) => ({ ...old, [key]: checked })); }} />{t(label)}</label>)}</fieldset>
      </>}
      {catalog.mode === 'confirm' && <><label>{t('&Reenter password to proceed.')}<input type="password" autoComplete="new-password" value={confirmation} onChange={(event) => { const value = event.currentTarget.value; setConfirmation(value); }} /></label><p>{t('Caution: If you lose or forget the password, it cannot be recovered.')}</p></>}
      {catalog.mode === 'unprotect' && <label>{t('&Password:')}<input type="password" autoComplete="off" value={password} onChange={(event) => { const value = event.currentTarget.value; setPassword(value); }} /></label>}
    </div>
    <div className="excelx-protection-dialog__actions"><button type="button" className="primary" onClick={submit}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}

type WorkbookCatalog = { mode: 'protect' | 'unprotect'; structure?: boolean; windows?: boolean };
type WorkbookRuntime = { dlg?: { reactWorkbookProtectionCatalog?(): WorkbookCatalog | null; reactWorkbookProtectionAction?(action: string, value?: unknown): Promise<boolean> | boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function workbookRuntime(): WorkbookRuntime | undefined { return (window as Window & { L?: WorkbookRuntime }).L; }

/** React-owned workbook structure/window protection dialog. */
export function ExcelxWorkbookProtectionDialog() {
  const [catalog, setCatalog] = useState<WorkbookCatalog | null>(null);
  const [structure, setStructure] = useState(true);
  const [windows, setWindows] = useState(false);
  const [password, setPassword] = useState('');
  useEffect(() => {
    const refresh = () => { const next = workbookRuntime()?.dlg?.reactWorkbookProtectionCatalog?.() || null; setCatalog(next); if (next) { setStructure(next.structure !== false); setWindows(!!next.windows); } setPassword(''); };
    const off = workbookRuntime()?.bus?.on('excelx-workbook-protection-dialog', refresh); refresh(); return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => workbookRuntime()?.t?.(message) || message;
  const action = workbookRuntime()?.dlg?.reactWorkbookProtectionAction;
  const close = () => { void action?.('close'); };
  const submit = () => catalog.mode === 'protect' ? action?.('protect', { structure, windows, password }) : action?.('unprotect', password);
  return <OfficeDialog title={t(catalog.mode === 'protect' ? 'Protect Workbook' : 'Unprotect Workbook')} className="excelx-workbook-protection-dialog" style={{ width: 'min(94vw, 360px)' }} onRequestClose={close}>
    <div className="excelx-workbook-protection-dialog__body">
      {catalog.mode === 'protect' ? <><fieldset><legend>{t('Protect workbook for')}</legend><label><input type="checkbox" checked={structure} onChange={(event) => { const checked = event.currentTarget.checked; setStructure(checked); }} />{t('&Structure').replace('&', '')}</label><label><input type="checkbox" checked={windows} onChange={(event) => { const checked = event.currentTarget.checked; setWindows(checked); }} />{t('&Windows').replace('&', '')}</label></fieldset><label>{t('&Password (optional):').replace('&', '')}<input type="password" autoComplete="new-password" value={password} onChange={(event) => { const value = event.currentTarget.value; setPassword(value); }} /></label></> : <label>{t('&Password:')}<input type="password" autoComplete="off" value={password} onChange={(event) => { const value = event.currentTarget.value; setPassword(value); }} /></label>}
    </div>
    <div className="excelx-protection-dialog__actions"><button type="button" className="primary" onClick={() => { void submit(); }}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
