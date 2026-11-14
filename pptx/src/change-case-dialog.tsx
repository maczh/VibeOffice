import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type CaseMode = 'sentence' | 'lower' | 'upper' | 'title' | 'toggle';
interface ChangeCaseCatalog { mode: CaseMode; options: Array<[CaseMode, string]> }
interface ChangeCaseRuntime {
  dlg?: { reactChangeCaseCatalog?(): ChangeCaseCatalog | null; reactChangeCaseAction?(action: string, mode?: CaseMode): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): ChangeCaseRuntime | undefined { return (window as Window & { L?: ChangeCaseRuntime }).L; }

/** React-owned change-case choices; text transformation remains in the presentation formatter. */
export function PptxChangeCaseDialog() {
  const [catalog, setCatalog] = useState<ChangeCaseCatalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactChangeCaseCatalog?.() || null);
    const off = runtime()?.bus?.on('pptx-change-case-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const act = (action: string, mode?: CaseMode) => runtime()?.dlg?.reactChangeCaseAction?.(action, mode);
  return <OfficeDialog title={t('Change Case')} className="pptx-change-case-dialog" style={{ width: 'min(92vw, 300px)' }} onRequestClose={() => act('close')}>
    <fieldset><legend>{t('Change case')}</legend>
      {catalog.options.map(([mode, label]) => <label key={mode}><input type="radio" name="pptx-change-case" value={mode} checked={catalog.mode === mode} onChange={() => { act('select', mode); setCatalog({ ...catalog, mode }); }} />{t(label.replace('&', ''))}</label>)}
    </fieldset>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={() => act('apply')}>{t('OK')}</button><button type="button" onClick={() => act('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
