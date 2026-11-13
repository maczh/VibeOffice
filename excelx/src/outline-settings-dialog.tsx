import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface OutlineSettings { below: boolean; right: boolean; applyStyles: boolean }
interface OutlineRuntime {
  dlg?: { reactOutlineSettingsCatalog?(): OutlineSettings | null; reactOutlineSettingsApply?(state: OutlineSettings): boolean; reactOutlineSettingsCreate?(): void; reactOutlineSettingsClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): OutlineRuntime | undefined { return (window as Window & { L?: OutlineRuntime }).L; }

export function ExcelxOutlineSettingsDialog() {
  const [catalog, setCatalog] = useState<OutlineSettings | null>(null);
  const [state, setState] = useState<OutlineSettings>({ below: true, right: true, applyStyles: false });
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactOutlineSettingsCatalog?.() || null; setCatalog(next); if (next) setState(next); };
    const off = runtime()?.bus?.on('excelx-outline-settings-dialog', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactOutlineSettingsCatalog?.()) runtime()?.dlg?.reactOutlineSettingsClose?.(); };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactOutlineSettingsClose?.();
  const apply = () => runtime()?.dlg?.reactOutlineSettingsApply?.(state);
  const update = (patch: Partial<OutlineSettings>) => setState((current) => ({ ...current, ...patch }));
  return <OfficeDialog title={translate('Settings')} className="excelx-outline-settings-dialog" style={{ width: 'min(92vw, 360px)' }} onRequestClose={close}>
    <div className="excelx-outline-settings-dialog__body">
      <fieldset><legend>{translate('Direction')}</legend>
        <label><input type="checkbox" checked={state.below} onChange={(event) => update({ below: event.currentTarget.checked })} />{translate('Summary rows &below detail').replace('&', '')}</label>
        <label><input type="checkbox" checked={state.right} onChange={(event) => update({ right: event.currentTarget.checked })} />{translate('Summary columns to ri&ght of detail').replace('&', '')}</label>
      </fieldset>
      <label><input type="checkbox" checked={state.applyStyles} onChange={(event) => update({ applyStyles: event.currentTarget.checked })} />{translate('&Automatic styles').replace('&', '')}</label>
    </div>
    <div className="excelx-outline-settings-dialog__actions"><button type="button" onClick={() => runtime()?.dlg?.reactOutlineSettingsCreate?.()}>{translate('&Create').replace('&', '')}</button><span /><button type="button" className="primary" onClick={apply}>{translate('OK')}</button><button type="button" onClick={close}>{translate('Cancel')}</button></div>
  </OfficeDialog>;
}
