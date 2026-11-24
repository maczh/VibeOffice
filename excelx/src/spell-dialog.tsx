import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import './spell-dialog.css';
import { useExcelxRuntime } from './editor-host';

type Catalog = { title: string; word: string; before: string; after: string; suggestions: string[]; language: string; languages: { value: string; label: string }[]; replacement: string };
type Runtime = { spell?: { reactSpellCatalog?(): Catalog | null; reactSpellAction?(action: string, patch?: Record<string, string>): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
/** React owns the spelling review dialog; dictionary lookup and cell edits stay in the spelling service. */
export function ExcelxSpellDialog() {
  const runtime = useExcelxRuntime<Runtime>();
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [value, setValue] = useState('');
  const [suggestion, setSuggestion] = useState('');
  useEffect(() => { const refresh = () => { const next = runtime?.spell?.reactSpellCatalog?.() || null; setCatalog(next); if (next) { setValue(next.word); setSuggestion(next.suggestions[0] || ''); } }; const off = runtime?.bus?.on('excelx-spell-dialog', refresh); return () => { off?.(); runtime?.spell?.reactSpellAction?.('unmount'); }; }, []);
  if (!catalog) return null;
  const t = (text: string) => runtime?.t?.(text) || text;
  const action = (name: string, patch?: Record<string, string>) => runtime?.spell?.reactSpellAction?.(name, patch);
  const close = () => { action('cancel'); setCatalog(null); };
  const choose = (name: string) => action(name, { value: value !== catalog.word ? value : suggestion });
  return <OfficeDialog title={t(catalog.title)} className="excelx-spell-dialog" style={{ width: 'min(94vw, 540px)' }} onRequestClose={close}>
    <div className="excelx-spell-dialog__grid"><section><label>{t('Not in Dictionary:')}</label><div className="excelx-spell-dialog__context"><span>{catalog.before}</span><strong>{catalog.word}</strong><span>{catalog.after}</span></div><input aria-label={t('Replace with')} value={value} onChange={(event) => { setValue(event.currentTarget.value); action('value', { value: event.currentTarget.value }); }} /><label>{t('Suggestions:')}</label><select size={6} value={suggestion} disabled={!catalog.suggestions.length} onChange={(event) => { setSuggestion(event.currentTarget.value); action('suggestion', { value: event.currentTarget.value }); }} onDoubleClick={() => choose('change')}>{catalog.suggestions.map((item) => <option translate="no" key={item}>{item}</option>)}</select><label>{t('Dictionary &language:')}<select value={catalog.language} onChange={(event) => action('language', { value: event.currentTarget.value })}>{catalog.languages.map((item) => <option key={item.value} value={item.value}>{t(item.label)}</option>)}</select></label></section><section className="excelx-spell-dialog__buttons"><button type="button" onClick={() => action('ignore')}>{t('&Ignore Once').replace('&','')}</button><button type="button" onClick={() => action('ignoreAll')}>{t('I&gnore All').replace('&','')}</button><button type="button" onClick={() => action('add')}>{t('&Add to Dictionary').replace('&','')}</button><button type="button" onClick={() => choose('change')}>{t('&Change').replace('&','')}</button><button type="button" onClick={() => choose('changeAll')}>{t('Change A&ll').replace('&','')}</button><button type="button" onClick={() => choose('autoCorrect')}>{t('Auto&Correct').replace('&','')}</button></section></div>
    <div className="excelx-spell-dialog__actions"><button type="button" onClick={() => { action('options'); setCatalog(null); }}>{t('&Options...').replace('&','')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
