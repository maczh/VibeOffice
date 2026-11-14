import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import './customize-dialog.css';

type Catalog = { toolbars: { id: string; label: string; visible: boolean }[] };
type Runtime = { dlg?: { reactCustomizeCatalog?(): Catalog | null; reactCustomizeToggle?(id: string, visible: boolean): boolean; reactCustomizeClose?(): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }
const OPTIONS = ['Show Standard and Formatting toolbars on &two rows', 'Show Screen&Tips on toolbars', 'Show shortcut &keys in ScreenTips'];

/** Toolbar visibility controls are React-owned; the Excelx app retains its toolbar operations. */
export function ExcelxCustomizeDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [tab, setTab] = useState('Toolbars');
  const [options, setOptions] = useState([true, true, true]);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactCustomizeCatalog?.() || null; setCatalog(next); if (next) setTab('Toolbars'); };
    const off = runtime()?.bus?.on('excelx-customize-dialog', refresh); return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactCustomizeClose?.();
  return <OfficeDialog title={t('Customize')} className="excelx-customize-dialog" style={{ width: 'min(94vw, 400px)' }} onRequestClose={close}>
    <div className="excelx-customize-dialog__tabs" role="tablist">{['Toolbars','Options'].map((name) => <button key={name} type="button" role="tab" aria-selected={tab === name} onClick={() => setTab(name)}>{t(name)}</button>)}</div>
    {tab === 'Toolbars' ? <div className="excelx-customize-dialog__toolbars" role="tabpanel">{catalog.toolbars.map((toolbar) => <label key={toolbar.id}><input type="checkbox" checked={toolbar.visible} onChange={(event) => { const visible = event.currentTarget.checked; if (runtime()?.dlg?.reactCustomizeToggle?.(toolbar.id, visible)) setCatalog((current) => current ? { ...current, toolbars: current.toolbars.map((item) => item.id === toolbar.id ? { ...item, visible } : item) } : current); }} />{t(toolbar.label)}</label>)}</div>
      : <div className="excelx-customize-dialog__options" role="tabpanel">{OPTIONS.map((label, index) => <label key={label}><input type="checkbox" checked={options[index]} onChange={(event) => setOptions((current) => current.map((value, item) => item === index ? event.currentTarget.checked : value))} />{t(label).replace('&', '')}</label>)}</div>}
    <div className="excelx-customize-dialog__actions"><button type="button" className="primary" onClick={close}>{t('Close')}</button></div>
  </OfficeDialog>;
}
