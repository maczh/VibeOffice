import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import './save-as-dialog.css';
import { useExcelxRuntime } from './editor-host';

type Catalog = { name: string; type: string; types: [string, string][]; password: string };
type Runtime = { dlg?: { reactSaveAsCatalog?(): Catalog | null; reactSaveAsAction?(action: string, name?: string, type?: string, password?: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
/** Save As and its encryption-password confirmation are React-owned; workbook writing stays in the engine. */
export function ExcelxSaveAsDialog() {
  const runtime = useExcelxRuntime<Runtime>();
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [view, setView] = useState<'save' | 'options' | 'confirm'>('save');
  const [name, setName] = useState('');
  const [type, setType] = useState('xlsx');
  const [password, setPassword] = useState('');
  const [confirmation, setConfirmation] = useState('');
  const [error, setError] = useState('');
  useEffect(() => {
    const refresh = () => { const next = runtime?.dlg?.reactSaveAsCatalog?.() || null; setCatalog(next); if (next) { setName(next.name); setType(next.type); setPassword(next.password); setConfirmation(''); setError(''); setView('save'); } };
    const off = runtime?.bus?.on('excelx-save-as-dialog', refresh); return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (text: string) => runtime?.t?.(text) || text;
  const action = runtime?.dlg?.reactSaveAsAction;
  const cancel = () => action?.('cancel');
  const save = () => {
    if (password) { setConfirmation(''); setError(''); setView('confirm'); return; }
    action?.('save', name, type, '');
  };
  const confirm = () => {
    if (confirmation !== password) { setError(t('Confirmation password is not identical.')); return; }
    action?.('save', name, type, password);
  };
  return <OfficeDialog title={t('Save As')} className="excelx-save-as-dialog" style={{ width: 'min(94vw, 460px)' }} onRequestClose={cancel}>
    {view === 'save' && <div className="excelx-save-as-dialog__body"><label>{t('File &name:').replace('&', '')}<input autoFocus value={name} onChange={(event) => setName(event.currentTarget.value)} /></label><label>{t('Save as &type:').replace('&', '')}<select value={type} onChange={(event) => setType(event.currentTarget.value)}>{catalog.types.map(([value, label]) => <option key={value} value={value}>{t(label)}</option>)}</select></label><p className="hint">{t("The file is saved to your browser's downloads.")}</p><button type="button" onClick={() => setView('options')}>{t('Tools')}</button><div className="excelx-save-as-dialog__actions"><button type="button" className="primary" onClick={save}>{t('&Save').replace('&', '')}</button><button type="button" onClick={cancel}>{t('Cancel')}</button></div></div>}
    {view === 'options' && <div className="excelx-save-as-dialog__body"><h3>{t('General Options...')}</h3><label>{t('Password to &open:').replace('&', '')}<input type="password" autoComplete="new-password" value={password} onChange={(event) => setPassword(event.currentTarget.value)} /></label><p className="hint">{t('Workbooks saved with a password are encrypted (AES-256) and open in Excel 2010 and later.')}</p><div className="excelx-save-as-dialog__actions"><button type="button" className="primary" onClick={() => setView('save')}>{t('OK')}</button><button type="button" onClick={() => { setPassword(catalog.password); setView('save'); }}>{t('Cancel')}</button></div></div>}
    {view === 'confirm' && <div className="excelx-save-as-dialog__body"><h3>{t('Confirm Password')}</h3><label>{t('Reenter password to proceed.')}<input type="password" autoComplete="new-password" value={confirmation} onChange={(event) => { setConfirmation(event.currentTarget.value); setError(''); }} /></label><p className="hint">{t('Caution: If you lose or forget the password, it cannot be recovered.')}</p>{error && <p className="excelx-save-as-dialog__error" role="alert">{error}</p>}<div className="excelx-save-as-dialog__actions"><button type="button" className="primary" onClick={confirm}>{t('OK')}</button><button type="button" onClick={() => setView('options')}>{t('Cancel')}</button></div></div>}
  </OfficeDialog>;
}
