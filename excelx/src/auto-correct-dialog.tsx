import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Catalog = { options: Record<string, boolean>; entries: [string, string][] };
type Runtime = { dlg?: { reactAutoCorrectCatalog?(): Catalog | null; reactAutoCorrectAction?(action: string, from?: string, to?: string, options?: Record<string, boolean>): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }
const OPTIONS: [string, string][] = [
  ['acTwoCaps', 'Correct TWo INitial CApitals'], ['acSentence', 'Capitalize first letter of &sentences'],
  ['acDays', 'Capitalize &names of days'], ['acCapsLock', 'Correct accidental use of cAPS &LOCK key'],
  ['acReplace', 'Replace &text as you type'],
];

/** React-owned AutoCorrect UI; the legacy engine retains its dictionary and option persistence. */
export function ExcelxAutoCorrectDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [options, setOptions] = useState<Record<string, boolean>>({});
  const [entries, setEntries] = useState<[string, string][]>([]);
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactAutoCorrectCatalog?.() || null;
      setCatalog(next);
      if (next) { setOptions(next.options); setEntries(next.entries); }
    };
    const off = runtime()?.bus?.on('excelx-auto-correct-dialog', refresh);
    refresh();
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = runtime()?.dlg?.reactAutoCorrectAction;
  const close = () => action?.('close');
  const choose = (key: string, value: string) => { setFrom(key); setTo(value); };
  const mutateEntry = (kind: 'add' | 'delete') => {
    if (!from) return;
    action?.(kind, from, to);
    if (kind === 'delete') { setFrom(''); setTo(''); }
  };
  return <OfficeDialog title={t('AutoCorrect: English (U.S.)')} className="excelx-auto-correct-dialog" style={{ width: 'min(94vw, 460px)' }} onRequestClose={close}>
    <section className="excelx-auto-correct-dialog__body">
      {OPTIONS.map(([key, label]) => <label key={key}><input type="checkbox" checked={!!options[key]} onChange={(event) => { const checked = event.currentTarget.checked; setOptions((old) => ({ ...old, [key]: checked })); }} />{t(label).replace('&', '')}</label>)}
      <div className="excelx-auto-correct-dialog__fields">
        <label>{t('&Replace:').replace('&', '')}<input aria-label={t('&Replace:').replace('&', '')} value={from} onChange={(event) => setFrom(event.currentTarget.value)} /></label>
        <label>{t('&With:').replace('&', '')}<input aria-label={t('&With:').replace('&', '')} value={to} onChange={(event) => setTo(event.currentTarget.value)} /></label>
      </div>
      <div className="excelx-auto-correct-dialog__entries" role="listbox" aria-label={t('AutoCorrect entries')}>
        {entries.map(([key, value]) => <button type="button" role="option" aria-selected={from === key} key={key} onClick={() => choose(key, value)}>{key} → {value}</button>)}
      </div>
      <div className="excelx-auto-correct-dialog__entry-actions"><button type="button" onClick={() => mutateEntry('add')}>{t('&Add').replace('&', '')}</button><button type="button" onClick={() => mutateEntry('delete')}>{t('&Delete').replace('&', '')}</button></div>
    </section>
    <div className="excelx-auto-correct-dialog__actions"><button type="button" className="primary" onClick={() => action?.('save', '', '', options)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
