import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState } from 'react';

type ActionKind = 'next' | 'prev' | 'first' | 'last' | 'lastViewed' | 'end' | 'slide' | 'url';
interface ActionSettingsCatalog { none: boolean; kind: ActionKind; slide: string; url: string; actions: Array<[ActionKind, string]>; slides: Array<{ id: string; label: string }> }
interface ActionSettingsRuntime {
  dlg?: { reactActionSettingsCatalog?(): ActionSettingsCatalog | null; reactActionSettingsAction?(action: string, values?: Pick<ActionSettingsCatalog, 'none' | 'kind' | 'slide' | 'url'>): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): ActionSettingsRuntime | undefined { return (window as Window & { L?: ActionSettingsRuntime }).L; }

/** React-owned mouse-action settings; selected-shape hyperlink transactions remain engine-backed. */
export function PptxActionSettingsDialog() {
  const [catalog, setCatalog] = useState<ActionSettingsCatalog | null>(null);
  const id = useId().replaceAll(':', '');
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactActionSettingsCatalog?.() || null; setCatalog(next ? { ...next, actions: [...next.actions], slides: [...next.slides] } : null); };
    const off = runtime()?.bus?.on('pptx-action-settings-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const update = (changes: Partial<ActionSettingsCatalog>) => setCatalog({ ...catalog, ...changes });
  const close = () => runtime()?.dlg?.reactActionSettingsAction?.('close');
  const apply = () => runtime()?.dlg?.reactActionSettingsAction?.('apply', { none: catalog.none, kind: catalog.kind, slide: catalog.slide, url: catalog.url });
  return <OfficeDialog title={t('Action Settings')} className="pptx-action-settings-dialog" style={{ width: 'min(92vw, 430px)' }} onRequestClose={close}>
    <div className="tabs-head"><span id={`${id}-mouse-click-title`} className="tab on">{t('Mouse Click')}</span></div>
    <section role="tabpanel" aria-labelledby={`${id}-mouse-click-title`} className="col"><strong>{t('Action on click')}</strong>
      <label><input type="radio" name="pptx-action-mode" checked={catalog.none} onChange={() => update({ none: true })} />{t('&None').replace('&', '')}</label>
      <label><input type="radio" name="pptx-action-mode" checked={!catalog.none} onChange={() => update({ none: false })} />{t('&Hyperlink to:').replace('&', '')}</label>
      <div className="col" style={{ marginLeft: 20 }}>
        <select aria-label={t('Action type')} value={catalog.kind} disabled={catalog.none} onChange={(event) => update({ kind: event.currentTarget.value as ActionKind })}>{catalog.actions.map(([value, label]) => <option key={value} value={value}>{t(label)}</option>)}</select>
        {catalog.kind === 'slide' && <select aria-label={t('Slide')} value={catalog.slide} disabled={catalog.none} translate="no" onChange={(event) => update({ slide: event.currentTarget.value })}>{catalog.slides.map((slide) => <option key={slide.id} value={slide.id}>{slide.label}</option>)}</select>}
        {catalog.kind === 'url' && <input aria-label={t('URL')} type="text" value={catalog.url} disabled={catalog.none} onChange={(event) => update({ url: event.currentTarget.value })} />}
      </div>
    </section>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={apply}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
