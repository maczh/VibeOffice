import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState, type KeyboardEvent } from 'react';
import { useExcelxRuntime } from './editor-host.js';

type Catalog = { options: Record<string, boolean>; entries: [string, string][] };
type Runtime = { dlg?: { reactAutoCorrectCatalog?(): Catalog | null; reactAutoCorrectAction?(action: string, from?: string, to?: string, options?: Record<string, boolean>): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
const OPTIONS: [string, string][] = [
  ['acTwoCaps', 'Correct TWo INitial CApitals'], ['acSentence', 'Capitalize first letter of &sentences'],
  ['acDays', 'Capitalize &names of days'], ['acCapsLock', 'Correct accidental use of cAPS &LOCK key'],
  ['acReplace', 'Replace &text as you type'],
];

/** React-owned AutoCorrect UI; the legacy engine retains its dictionary and option persistence. */
export function ExcelxAutoCorrectDialog() {
  const runtime = useExcelxRuntime<Runtime>();
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [options, setOptions] = useState<Record<string, boolean>>({});
  const [entries, setEntries] = useState<[string, string][]>([]);
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [activeEntry, setActiveEntry] = useState(0);
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.dlg?.reactAutoCorrectCatalog?.() || null;
      setCatalog(next);
      if (next) { setOptions(next.options); setEntries(next.entries); const index = next.entries.findIndex(([key]) => key === from); setActiveEntry(index >= 0 ? index : 0); }
    };
    const off = runtime?.bus?.on('excelx-auto-correct-dialog', refresh);
    refresh();
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const action = runtime?.dlg?.reactAutoCorrectAction;
  const close = () => action?.('close');
  const choose = (key: string, value: string) => { setFrom(key); setTo(value); const index = entries.findIndex(([item]) => item === key); if (index >= 0) setActiveEntry(index); };
  const onEntriesKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (!entries.length) return;
    let next = activeEntry;
    if (event.key === 'ArrowDown' || event.key === 'ArrowRight') next = (activeEntry + 1) % entries.length;
    else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') next = (activeEntry - 1 + entries.length) % entries.length;
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = entries.length - 1;
    else if (event.key !== 'Enter' && event.key !== ' ') return;
    event.preventDefault();
    event.stopPropagation();
    const [key, value] = entries[next] || [];
    if (key !== undefined && value !== undefined) choose(key, value);
  };
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
      <div id="excelx-auto-correct-entries" className="excelx-auto-correct-dialog__entries" role="listbox" aria-label={t('AutoCorrect entries')} aria-activedescendant={entries[activeEntry] ? `excelx-auto-correct-entry-${activeEntry}` : undefined} tabIndex={0} onKeyDown={onEntriesKeyDown}>
        {entries.map(([key, value], index) => <button id={`excelx-auto-correct-entry-${index}`} type="button" role="option" aria-label={`${key} → ${value}`} aria-selected={from === key} tabIndex={-1} key={key} onPointerEnter={() => setActiveEntry(index)} onClick={() => choose(key, value)}>{key} → {value}</button>)}
      </div>
      <div className="excelx-auto-correct-dialog__entry-actions"><button type="button" onClick={() => mutateEntry('add')}>{t('&Add').replace('&', '')}</button><button type="button" onClick={() => mutateEntry('delete')}>{t('&Delete').replace('&', '')}</button></div>
    </section>
    <div className="excelx-auto-correct-dialog__actions"><button type="button" className="primary" onClick={() => action?.('save', '', '', options)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
