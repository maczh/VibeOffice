import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

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
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactAutoCorrectCatalog?.() || null;
      setCatalog(next ? { ...next, options: { ...next.options }, entries: [...next.entries] } : null);
      setEntry(next?.selected || '');
      setReplacement(next?.replacement || '');
    };
    const off = runtime()?.bus?.on('pptx-autocorrect-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const dispatch = (action: string, value?: string, secondary?: string | boolean) => runtime()?.dlg?.reactAutoCorrectAction?.(action, value, secondary);
  const select = (key: string, value: string) => { dispatch('select', key); setEntry(key); setReplacement(value); };
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
      <div role="listbox" aria-label={t('AutoCorrect replacements')} className="list-box" style={{ height: 150, overflow: 'auto' }}>
        {catalog.entries.map(([key, value]) => <button key={key} type="button" role="option" aria-selected={entry === key} translate="no" onClick={() => select(key, value)}>{key}<span aria-hidden="true"> → </span>{value}</button>)}
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
