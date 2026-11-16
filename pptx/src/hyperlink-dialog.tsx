import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState } from 'react';

type HyperlinkKind = 'web' | 'place' | 'mail';
interface HyperlinkPlace { value: string; label: string }
interface HyperlinkCatalog { title: string; existing: boolean; editing: boolean; kind: HyperlinkKind; display: string; address: string; tip: string; email: string; subject: string; place: string; places: HyperlinkPlace[] }
interface HyperlinkRuntime {
  dlg?: { reactHyperlinkCatalog?(): HyperlinkCatalog | null; reactHyperlinkAction?(action: string, values?: Omit<HyperlinkCatalog, 'title' | 'existing' | 'editing' | 'places'>): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}
function runtime(): HyperlinkRuntime | undefined { return (window as Window & { L?: HyperlinkRuntime }).L; }

/** React-owned hyperlink destination forms; links continue through the selected text/object history APIs. */
export function PptxHyperlinkDialog() {
  const [catalog, setCatalog] = useState<HyperlinkCatalog | null>(null);
  const id = useId().replaceAll(':', '');
  const tabId = (kind: HyperlinkKind) => `${id}-hyperlink-tab-${kind}`;
  const panelId = `${id}-hyperlink-panel`;
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactHyperlinkCatalog?.() || null; setCatalog(next ? { ...next, places: [...next.places] } : null); };
    const off = runtime()?.bus?.on('pptx-hyperlink-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string, ...args: unknown[]) => runtime()?.t?.(message, ...args) || message;
  const update = (changes: Partial<HyperlinkCatalog>) => setCatalog({ ...catalog, ...changes });
  const values = { kind: catalog.kind, display: catalog.display, address: catalog.address, tip: catalog.tip, email: catalog.email, subject: catalog.subject, place: catalog.place };
  const dispatch = (action: string) => runtime()?.dlg?.reactHyperlinkAction?.(action, values);
  const close = () => runtime()?.dlg?.reactHyperlinkAction?.('close');
  const kinds: Array<[HyperlinkKind, string]> = [['web', 'Existing File or Web Page'], ['place', 'Place in This Document'], ['mail', 'E-mail Address']];
  return <OfficeDialog title={t(catalog.title)} className="pptx-hyperlink-dialog" style={{ width: 'min(94vw, 640px)' }} onRequestClose={close}>
    <div className="col">
      <label>{t('&Text to display:').replace('&', '')}<input aria-label={t('Text to display:')} type="text" value={catalog.display} disabled={!catalog.editing} onChange={(event) => update({ display: event.currentTarget.value })} /></label>
      <div className="row" style={{ alignItems: 'stretch', gap: 12 }}>
        <section className="col" style={{ flex: '0 0 150px' }}><strong>{t('Link to:')}</strong><div role="tablist" aria-label={t('Link to:')} className="col">
          {kinds.map(([kind, label], index) => <button key={kind} id={tabId(kind)} type="button" role="tab" aria-controls={panelId} aria-selected={catalog.kind === kind} tabIndex={catalog.kind === kind ? 0 : -1} onClick={() => update({ kind })} onKeyDown={(event) => {
            const nextIndex = event.key === 'ArrowRight' ? (index + 1) % kinds.length
              : event.key === 'ArrowLeft' ? (index - 1 + kinds.length) % kinds.length
                : event.key === 'Home' ? 0
                  : event.key === 'End' ? kinds.length - 1 : -1;
            if (nextIndex < 0) return;
            event.preventDefault();
            const next = kinds[nextIndex];
            if (!next) return;
            update({ kind: next[0] });
            event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[nextIndex]?.focus();
          }}>{t(label)}</button>)}
        </div></section>
        <section id={panelId} aria-labelledby={tabId(catalog.kind)} role="tabpanel" className="col" style={{ flex: '1 1 auto', minWidth: 0 }}>
          {catalog.kind === 'web' && <><label>{t('Addr&ess:').replace('&', '')}<input aria-label={t('Address:')} type="text" value={catalog.address} onChange={(event) => update({ address: event.currentTarget.value })} /></label><label>{t('ScreenTip:')}<input aria-label={t('ScreenTip:')} type="text" value={catalog.tip} onChange={(event) => update({ tip: event.currentTarget.value })} /></label></>}
          {catalog.kind === 'place' && <label>{t('Select a place in this d&ocument:').replace('&', '')}<select aria-label={t('Place in this document')} size={10} value={catalog.place} onChange={(event) => update({ place: event.currentTarget.value })}>{catalog.places.map((place) => <option key={place.value} value={place.value}>{t(place.label)}</option>)}</select></label>}
          {catalog.kind === 'mail' && <><label>{t('&E-mail address:').replace('&', '')}<input aria-label={t('E-mail address:')} type="text" value={catalog.email} onChange={(event) => update({ email: event.currentTarget.value })} /></label><label>{t('Su&bject:').replace('&', '')}<input aria-label={t('Subject:')} type="text" value={catalog.subject} onChange={(event) => update({ subject: event.currentTarget.value })} /></label></>}
        </section>
      </div>
    </div>
    <div className="vibeoffice-dialog__actions">{catalog.existing && <button type="button" onClick={() => dispatch('remove')}>{t('&Remove Link').replace('&', '')}</button>}<button className="primary" type="button" onClick={() => dispatch('apply')}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
