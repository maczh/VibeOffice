import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useRef, useState, type KeyboardEvent } from 'react';

type AutoCorrectOption = 'autocorrect' | 'capSentence' | 'smartQuotes';
interface AutoCorrectCatalog { options: Record<AutoCorrectOption, boolean>; entries: Array<[string, string]>; selected: string; replacement: string }
interface AutoCorrectRuntime {
  dlg?: { reactAutoCorrectCatalog?(): AutoCorrectCatalog | null; reactAutoCorrectAction?(action: string, value?: string, secondary?: string | boolean): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): AutoCorrectRuntime | undefined { return (window as Window & { L?: AutoCorrectRuntime }).L; }
const options: Array<[AutoCorrectOption, string]> = [['autocorrect', 'Correct TWo INitial CApitals'], ['capSentence', 'Capitalize first letter of sentences'], ['smartQuotes', 'Replace straight quotes with smart quotes']];

/** React-owned AutoCorrect dictionary and typing preferences; the text engine keeps the live dictionary. */
export function PptxAutoCorrectDialog() {
  const [catalog, setCatalog] = useState<AutoCorrectCatalog | null>(null);
  const [entry, setEntry] = useState('');
  const [replacement, setReplacement] = useState('');
  const [activeEntry, setActiveEntry] = useState(0);
  const listRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactAutoCorrectCatalog?.() || null;
      setCatalog(next ? { ...next, options: { ...next.options }, entries: [...next.entries] } : null);
      setEntry(next?.selected || '');
      setReplacement(next?.replacement || '');
      if (next) { const index = next.entries.findIndex(([key]) => key === next.selected); setActiveEntry(index >= 0 ? index : 0); }
    };
    const off = runtime()?.bus?.on('pptx-autocorrect-dialog', refresh);
    return () => off?.();
  }, []);
  useEffect(() => {
    const list = listRef.current;
    const option = list?.querySelector<HTMLElement>(`#pptx-autocorrect-entry-${activeEntry}`);
    if (!list || !option) return;
    const listBounds = list.getBoundingClientRect();
    const optionBounds = option.getBoundingClientRect();
    const visibleTop = listBounds.top + list.clientTop;
    const visibleBottom = visibleTop + list.clientHeight;
    if (optionBounds.top < visibleTop) list.scrollTop -= visibleTop - optionBounds.top;
    else if (optionBounds.bottom > visibleBottom) list.scrollTop += optionBounds.bottom - visibleBottom;
  }, [activeEntry, catalog]);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const dispatch = (action: string, value?: string, secondary?: string | boolean) => runtime()?.dlg?.reactAutoCorrectAction?.(action, value, secondary);
  const select = (key: string, value: string) => { dispatch('select', key); setEntry(key); setReplacement(value); const index = catalog.entries.findIndex(([item]) => item === key); if (index >= 0) setActiveEntry(index); };
  const onEntriesKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (!catalog.entries.length) return;
    let next = activeEntry;
    if (event.key === 'ArrowDown' || event.key === 'ArrowRight') next = (activeEntry + 1) % catalog.entries.length;
    else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') next = (activeEntry - 1 + catalog.entries.length) % catalog.entries.length;
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = catalog.entries.length - 1;
    else if (event.key !== 'Enter' && event.key !== ' ') return;
    event.preventDefault();
    event.stopPropagation();
    const [key, value] = catalog.entries[next] || [];
    if (key !== undefined && value !== undefined) select(key, value);
  };
  const save = () => dispatch('save');
  return <OfficeDialog title={t('AutoCorrect: English (U.S.)')} className="pptx-autocorrect-dialog" style={{ width: 'min(94vw, 480px)' }} onRequestClose={() => dispatch('close')}>
    <section className="col" aria-label={t('AutoCorrect')}>
      {options.map(([key, label]) => <label key={key}><input type="checkbox" checked={catalog.options[key]} onChange={(event) => {
        const checked = event.currentTarget.checked;
        dispatch('option', key, checked);
        setCatalog({ ...catalog, options: { ...catalog.options, [key]: checked } });
      }} />{t(label)}</label>)}
      <div className="row">
        <label>{t('&Replace:').replace('&', '')}<input aria-label={t('&Replace:').replace('&', '')} type="text" value={entry} onChange={(event) => setEntry(event.currentTarget.value)} /></label>
        <label>{t('&With:').replace('&', '')}<input aria-label={t('&With:').replace('&', '')} type="text" value={replacement} onChange={(event) => setReplacement(event.currentTarget.value)} /></label>
      </div>
      <div ref={listRef} id="pptx-autocorrect-entry-list" role="listbox" aria-label={t('AutoCorrect replacements')} aria-activedescendant={catalog.entries[activeEntry] ? `pptx-autocorrect-entry-${activeEntry}` : undefined} tabIndex={0} onKeyDown={onEntriesKeyDown} className="list-box" style={{ height: 150, overflow: 'auto' }}>
        {catalog.entries.map(([key, value], index) => <button id={`pptx-autocorrect-entry-${index}`} key={key} type="button" role="option" aria-label={`${key} → ${value}`} aria-selected={entry === key} tabIndex={-1} translate="no" onPointerEnter={() => setActiveEntry(index)} onClick={() => select(key, value)}>{key}<span aria-hidden="true"> → </span>{value}</button>)}
      </div>
      <div className="row"><button type="button" onClick={() => {
        if (!entry) return;
        dispatch('add', entry, replacement);
        const next = runtime()?.dlg?.reactAutoCorrectCatalog?.();
        if (next) setCatalog({ ...next, options: { ...next.options }, entries: [...next.entries] });
      }}>{t('&Add').replace('&', '')}</button><button type="button" disabled={!catalog.selected && !catalog.entries.some(([key]) => key === entry)} onClick={() => {
        dispatch('select', entry);
        dispatch('delete');
        const next = runtime()?.dlg?.reactAutoCorrectCatalog?.();
        if (next) { setCatalog({ ...next, options: { ...next.options }, entries: [...next.entries] }); setEntry(''); setReplacement(''); }
      }}>{t('&Delete').replace('&', '')}</button></div>
    </section>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={save}>{t('OK')}</button></div>
  </OfficeDialog>;
}
