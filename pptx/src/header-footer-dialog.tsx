import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

interface DateFormat { id: string; label: string }
interface HeaderFooterState { date: boolean; automatic: boolean; format: string; fixedText: string; number: boolean; footer: boolean; footerText: string; notOnTitle: boolean; formats: DateFormat[] }
interface HeaderFooterCommit { dt: boolean; dtAuto: boolean; dtFmt: string; dtText: string; num: boolean; ftr: boolean; ftrText: string; notOnTitle: boolean }
interface HeaderFooterRuntime {
  dlg?: { reactHeaderFooterCatalog?(): HeaderFooterState | null; reactHeaderFooterAction?(action: string, values?: HeaderFooterCommit, all?: boolean): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

/** React-owned slide header/footer form; slide and presentation history transactions stay in the engine. */
export function PptxHeaderFooterDialog() {
  const runtime = usePptxRuntime<HeaderFooterRuntime>();
  const [state, setState] = useState<HeaderFooterState | null>(null);
  useEffect(() => {
    const refresh = () => setState(runtime?.dlg?.reactHeaderFooterCatalog?.() || null);
    const off = runtime?.bus?.on('pptx-header-footer-dialog', refresh);
    return () => off?.();
  }, []);
  if (!state) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const update = (changes: Partial<HeaderFooterState>) => setState({ ...state, ...changes });
  const commit: HeaderFooterCommit = { dt: state.date, dtAuto: state.automatic, dtFmt: state.format, dtText: state.fixedText, num: state.number, ftr: state.footer, ftrText: state.footerText, notOnTitle: state.notOnTitle };
  const act = (action: string, all?: boolean) => runtime?.dlg?.reactHeaderFooterAction?.(action, action === 'apply' ? commit : undefined, all);
  return <OfficeDialog title={t('Header and Footer')} className="pptx-header-footer-dialog" style={{ width: 'min(94vw, 470px)' }} onRequestClose={() => act('close')}>
    <section className="col" aria-label={t('Slide')}>
      <strong>{t('Include on slide')}</strong>
      <label><input type="checkbox" aria-label={t('&Date and time').replace('&', '')} checked={state.date} onChange={(event) => update({ date: event.currentTarget.checked })} />{t('&Date and time').replace('&', '')}</label>
      <div className="col" style={{ marginLeft: 20 }}>
        <label><input type="radio" name="pptx-hf-date" checked={state.automatic} disabled={!state.date} onChange={() => update({ automatic: true })} />{t('Update &automatically').replace('&', '')}</label>
        <select aria-label={t('Date format')} value={state.format} disabled={!state.date || !state.automatic} onChange={(event) => update({ format: event.currentTarget.value })}>{state.formats.map((format) => <option key={format.id} value={format.id}>{format.label}</option>)}</select>
        <label><input type="radio" name="pptx-hf-date" checked={!state.automatic} disabled={!state.date} onChange={() => update({ automatic: false })} />{t('Fi&xed').replace('&', '')}</label>
        <input aria-label={t('Fixed date and time')} type="text" value={state.fixedText} disabled={!state.date || state.automatic} onChange={(event) => update({ fixedText: event.currentTarget.value })} />
      </div>
      <label><input type="checkbox" aria-label={t('Slide &number').replace('&', '')} checked={state.number} onChange={(event) => update({ number: event.currentTarget.checked })} />{t('Slide &number').replace('&', '')}</label>
      <label><input type="checkbox" aria-label={t('&Footer').replace('&', '')} checked={state.footer} onChange={(event) => update({ footer: event.currentTarget.checked })} />{t('&Footer').replace('&', '')}</label>
      <div style={{ marginLeft: 20 }}><input aria-label={t('Footer text')} type="text" value={state.footerText} disabled={!state.footer} onChange={(event) => update({ footerText: event.currentTarget.value })} /></div>
      <label><input type="checkbox" aria-label={t("Don't show on title &slide").replace('&', '')} checked={state.notOnTitle} onChange={(event) => update({ notOnTitle: event.currentTarget.checked })} />{t("Don't show on title &slide").replace('&', '')}</label>
    </section>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={() => act('apply', true)}>{t('Apply to &All').replace('&', '')}</button><button type="button" onClick={() => act('apply', false)}>{t('&Apply').replace('&', '')}</button><button type="button" onClick={() => act('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
