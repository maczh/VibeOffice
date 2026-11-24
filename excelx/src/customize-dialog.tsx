import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState } from 'react';
import './customize-dialog.css';
import { useExcelxRuntime } from './editor-host';

type Catalog = { toolbars: { id: string; label: string; visible: boolean }[] };
type Runtime = { dlg?: { reactCustomizeCatalog?(): Catalog | null; reactCustomizeToggle?(id: string, visible: boolean): boolean; reactCustomizeClose?(): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
const OPTIONS = ['Show Standard and Formatting toolbars on &two rows', 'Show Screen&Tips on toolbars', 'Show shortcut &keys in ScreenTips'];

/** Toolbar visibility controls are React-owned; the Excelx app retains its toolbar operations. */
export function ExcelxCustomizeDialog() {
  const runtime = useExcelxRuntime<Runtime>();
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [tab, setTab] = useState('Toolbars');
  const id = useId().replaceAll(':', '');
  const [options, setOptions] = useState([true, true, true]);
  useEffect(() => {
    const refresh = () => { const next = runtime?.dlg?.reactCustomizeCatalog?.() || null; setCatalog(next); if (next) setTab('Toolbars'); };
    const off = runtime?.bus?.on('excelx-customize-dialog', refresh); return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const close = () => runtime?.dlg?.reactCustomizeClose?.();
  const tabs = ['Toolbars', 'Options'];
  const tabId = (index: number) => `${id}-customize-tab-${index}`;
  const panelId = (index: number) => `${id}-customize-panel-${index}`;
  return <OfficeDialog title={t('Customize')} className="excelx-customize-dialog" style={{ width: 'min(94vw, 400px)' }} onRequestClose={close}>
    <div className="excelx-customize-dialog__tabs" role="tablist" aria-label={t('Customize')}>{tabs.map((name, index) => <button key={name} id={tabId(index)} type="button" role="tab" aria-controls={panelId(index)} aria-selected={tab === name} tabIndex={tab === name ? 0 : -1} onClick={() => setTab(name)} onKeyDown={(event) => {
      const next = event.key === 'ArrowRight' ? (index + 1) % tabs.length
        : event.key === 'ArrowLeft' ? (index - 1 + tabs.length) % tabs.length
          : event.key === 'Home' ? 0
            : event.key === 'End' ? tabs.length - 1 : -1;
      if (next < 0) return;
      event.preventDefault();
      const nextTab = tabs[next];
      if (nextTab) setTab(nextTab);
      event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
    }}>{t(name)}</button>)}</div>
    <div id={panelId(0)} aria-labelledby={tabId(0)} role="tabpanel" className="excelx-customize-dialog__toolbars" hidden={tab !== 'Toolbars'}>{catalog.toolbars.map((toolbar) => <label key={toolbar.id}><input type="checkbox" checked={toolbar.visible} onChange={(event) => { const visible = event.currentTarget.checked; if (runtime?.dlg?.reactCustomizeToggle?.(toolbar.id, visible)) setCatalog((current) => current ? { ...current, toolbars: current.toolbars.map((item) => item.id === toolbar.id ? { ...item, visible } : item) } : current); }} />{t(toolbar.label)}</label>)}</div>
    <div id={panelId(1)} aria-labelledby={tabId(1)} role="tabpanel" className="excelx-customize-dialog__options" hidden={tab !== 'Options'}>{OPTIONS.map((label, index) => <label key={label}><input type="checkbox" checked={options[index]} onChange={(event) => setOptions((current) => current.map((value, item) => item === index ? event.currentTarget.checked : value))} />{t(label).replace('&', '')}</label>)}</div>
    <div className="excelx-customize-dialog__actions"><button type="button" className="primary" onClick={close}>{t('Close')}</button></div>
  </OfficeDialog>;
}
