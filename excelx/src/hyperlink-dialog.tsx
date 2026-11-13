import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Kind = 'web' | 'place' | 'mail';
type Catalog = { editing: boolean; kind: Kind; text: string; address: string; cell: string; sheet: string; email: string; subject: string; tooltip: string; sheets: { value: string; label: string }[] };
type State = Omit<Catalog, 'editing' | 'sheets'>;
type Runtime = { dlg?: { reactHyperlinkCatalog?(): Catalog | null; reactHyperlinkAction?(action: string, state?: State): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }
const KINDS: [Kind, string][] = [['web', 'Existing File or Web Page'], ['place', 'Place in This Document'], ['mail', 'E-mail Address']];

/** React-owned Insert/Edit Hyperlink form; the existing workbook link transaction remains authoritative. */
export function ExcelxHyperlinkDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [form, setForm] = useState<State>({ kind: 'web', text: '', address: '', cell: 'A1', sheet: '', email: '', subject: '', tooltip: '' });
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactHyperlinkCatalog?.() || null; setCatalog(next); if (next) setForm({ kind: next.kind, text: next.text, address: next.address, cell: next.cell, sheet: next.sheet, email: next.email, subject: next.subject, tooltip: next.tooltip }); };
    const off = runtime()?.bus?.on('excelx-hyperlink-dialog', refresh); return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactHyperlinkAction?.('close');
  const update = (key: keyof State, value: string) => setForm((old) => ({ ...old, [key]: value }));
  const field = (label: string, key: keyof State, options: { placeholder?: string; translate?: 'yes' | 'no' } = {}) => <label>{t(label).replace('&', '')}<input type="text" value={form[key]} placeholder={options.placeholder} translate={options.translate} onChange={(event) => { const value = event.currentTarget.value; update(key, value); }} /></label>;
  return <OfficeDialog title={t(catalog.editing ? 'Edit Hyperlink' : 'Insert Hyperlink')} className="excelx-hyperlink-dialog" style={{ width: 'min(94vw, 560px)' }} onRequestClose={close}>
    <div className="excelx-hyperlink-dialog__body">
      {field('&Text to display:', 'text')}
      <div className="excelx-hyperlink-dialog__main">
        <div className="excelx-hyperlink-dialog__kinds" role="tablist" aria-label={t('Link type')}>
          {KINDS.map(([kind, label]) => <button key={kind} type="button" role="tab" aria-selected={form.kind === kind} onClick={() => update('kind', kind)}>{t(label)}</button>)}
        </div>
        <div className="excelx-hyperlink-dialog__target" role="tabpanel">
          {form.kind === 'web' && field('Addr&ess:', 'address', { placeholder: 'https://', translate: 'no' })}
          {form.kind === 'place' && <>{field('Type the cell refere&nce:', 'cell')}<label>{t('Or select a place in this document:')}<select value={form.sheet} onChange={(event) => { const value = event.currentTarget.value; update('sheet', value); }} translate="no">{catalog.sheets.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}</select></label></>}
          {form.kind === 'mail' && <>{field('&E-mail address:', 'email')} {field('Su&bject:', 'subject')}</>}
        </div>
      </div>
      {field('ScreenTi&p:', 'tooltip')}
    </div>
    <div className="excelx-hyperlink-dialog__actions">{catalog.editing && <button type="button" onClick={() => runtime()?.dlg?.reactHyperlinkAction?.('remove')}>{t('&Remove Link').replace('&', '')}</button>}<span /><button type="button" className="primary" onClick={() => runtime()?.dlg?.reactHyperlinkAction?.('apply', form)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
