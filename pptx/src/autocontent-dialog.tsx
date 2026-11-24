import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

interface AutoContentValues { type: string; title: string; footer: string; design: string; number: boolean; date: boolean }
interface AutoContentCatalog { types: string[]; designs: Array<{ key: string; name: string }>; defaultDesign: string }
interface AutoContentRuntime {
  app?: { reactAutoContentCatalog?(): AutoContentCatalog | null; reactAutoContentAction?(action: string, values?: AutoContentValues): Promise<boolean> | boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

/** React-owned AutoContent form; the presentation engine retains template generation and document replacement. */
export function PptxAutoContentDialog() {
  const runtime = usePptxRuntime<AutoContentRuntime>();
  const [catalog, setCatalog] = useState<AutoContentCatalog | null>(null);
  const [values, setValues] = useState<AutoContentValues>({ type: '', title: '', footer: '', design: 'azure', number: true, date: false });
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.app?.reactAutoContentCatalog?.() || null;
      setCatalog(next);
      if (next) setValues((previous) => ({ ...previous, type: previous.type && next.types.includes(previous.type) ? previous.type : next.types[0] || '', design: next.designs.some((design) => design.key === previous.design) ? previous.design : next.defaultDesign }));
    };
    const off = runtime?.bus?.on('pptx-autocontent-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const update = <K extends keyof AutoContentValues>(key: K, value: AutoContentValues[K]) => setValues((previous) => ({ ...previous, [key]: value }));
  const close = () => runtime?.app?.reactAutoContentAction?.('close');
  const create = async () => { await runtime?.app?.reactAutoContentAction?.('create', values); };
  return <OfficeDialog title={t('AutoContent Wizard')} className="pptx-autocontent-dialog" style={{ width: 'min(94vw, 560px)' }} onRequestClose={close}>
    <div className="row pptx-autocontent-fields">
      <label>{t('Presentation type:')}<select aria-label={t('Presentation type:')} size={6} value={values.type} onChange={(event) => update('type', event.currentTarget.value)}>
        {catalog.types.map((type) => <option key={type} value={type}>{t(type)}</option>)}
      </select></label>
      <div className="col">
        <label>{t('Presentation title:')}<input aria-label={t('Presentation title:')} type="text" placeholder={t('Presentation title')} value={values.title} onChange={(event) => update('title', event.currentTarget.value)} /></label>
        <label>{t('Footer (optional):')}<input aria-label={t('Footer (optional):')} type="text" placeholder={t('Footer (optional)')} value={values.footer} onChange={(event) => update('footer', event.currentTarget.value)} /></label>
        <label>{t('Design:')}<select aria-label={t('Design:')} value={values.design} onChange={(event) => update('design', event.currentTarget.value)}>
          {catalog.designs.map((design) => <option key={design.key} value={design.key}>{t(design.name)}</option>)}
        </select></label>
        <label className="ck"><input type="checkbox" aria-label={t('Slide number')} checked={values.number} onChange={(event) => update('number', event.currentTarget.checked)} /><span>{t('Slide number')}</span></label>
        <label className="ck"><input type="checkbox" aria-label={t('Date last updated')} checked={values.date} onChange={(event) => update('date', event.currentTarget.checked)} /><span>{t('Date last updated')}</span></label>
      </div>
    </div>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={() => { void create(); }}>{t('Finish')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
