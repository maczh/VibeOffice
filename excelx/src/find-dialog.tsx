import { useEffect, useState } from 'react';
import './find-dialog.css';

type SearchState = { what: string; with: string; within: 'sheet' | 'workbook'; by: 'rows' | 'cols'; lookIn: 'formulas' | 'values' | 'comments'; matchCase: boolean; whole: boolean; more: boolean };
type Result = { index: number; book: string; sheet: string; name: string; cell: string; value: string; formula: string };
type Catalog = { mode: number; state: SearchState; results: Result[]; status: string };
type Action = (action: string, key?: string, value?: string | number | boolean) => boolean;
type Runtime = { dlg?: { reactFindCatalog?(): Catalog | null; reactFindAction?: Action }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string, ...args: unknown[]): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** Modeless React Find/Replace window; workbook matching and replacement remain in Excelx operations. */
export function ExcelxFindDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactFindCatalog?.() || null);
    const off = runtime()?.bus?.on('excelx-find-dialog', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactFindCatalog?.()) runtime()?.dlg?.reactFindAction?.('close'); };
  }, []);
  if (!catalog) return null;
  const t = (message: string, ...args: unknown[]) => runtime()?.t?.(message, ...args) || message;
  const action = runtime()?.dlg?.reactFindAction;
  const set = (key: keyof SearchState, value: string | boolean) => action?.('set', key, value);
  const mnemonic = (value: string) => t(value).replace('&', '');
  return <section className="dlg excelx-find-dialog" role="dialog" aria-label={t('Find and Replace')}>
    <div className="dlg-ttl">{t('Find and Replace')}</div>
    <div className="dlg-body">
      <div className="excelx-find-dialog__tabs" role="tablist">
        {[t('Fi&nd').replace('&', ''), t('Re&place').replace('&', '')].map((label, index) => <button type="button" role="tab" aria-selected={catalog.mode === index} key={label} onClick={() => action?.('tab', undefined, index)}>{label}</button>)}
      </div>
      <div className="excelx-find-dialog__form">
        <label>{mnemonic('Fi&nd what:')}<input autoFocus value={catalog.state.what} onChange={(event) => set('what', event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); action?.(event.shiftKey ? 'previous' : 'next'); } }} /></label>
        {catalog.mode === 1 && <label>{mnemonic('R&eplace with:')}<input value={catalog.state.with} onChange={(event) => set('with', event.currentTarget.value)} /></label>}
        <button type="button" aria-expanded={catalog.state.more} onClick={() => set('more', !catalog.state.more)}>{t(catalog.state.more ? 'Options <<' : 'Op&tions >>').replace('&', '')}</button>
        {catalog.state.more && <div className="excelx-find-dialog__options">
          <label>{mnemonic('Wit&hin:')}<select value={catalog.state.within} onChange={(event) => set('within', event.currentTarget.value)}><option value="sheet">{t('Sheet')}</option><option value="workbook">{t('Workbook')}</option></select></label>
          <label>{mnemonic('&Search:')}<select value={catalog.state.by} onChange={(event) => set('by', event.currentTarget.value)}><option value="rows">{t('By Rows')}</option><option value="cols">{t('By Columns')}</option></select></label>
          <label>{mnemonic('&Look in:')}<select value={catalog.state.lookIn} onChange={(event) => set('lookIn', event.currentTarget.value)}><option value="formulas">{t('Formulas')}</option><option value="values">{t('Values')}</option><option value="comments">{t('Comments')}</option></select></label>
          <label><input type="checkbox" checked={catalog.state.matchCase} onChange={(event) => set('matchCase', event.currentTarget.checked)} />{mnemonic('Match &case')}</label>
          <label><input type="checkbox" checked={catalog.state.whole} onChange={(event) => set('whole', event.currentTarget.checked)} />{mnemonic('Match entire cell c&ontents')}</label>
        </div>}
      </div>
      {catalog.results.length > 0 && <div className="excelx-find-dialog__results"><table className="fr-table"><thead><tr>{['Book|workbook', 'Sheet', 'Name', 'Cell', 'Value', 'Formula'].map((heading) => <th key={heading}>{t(heading)}</th>)}</tr></thead><tbody>{catalog.results.map((result) => <tr key={`${result.book}-${result.sheet}-${result.cell}`} onClick={() => action?.('select', undefined, result.index)}><td translate="no">{result.book}</td><td translate="no">{result.sheet}</td><td translate="no">{result.name}</td><td translate="no">{result.cell}</td><td translate="no">{result.value}</td><td translate="no">{result.formula}</td></tr>)}</tbody></table></div>}
      <div className="hint" aria-live="polite">{catalog.status}</div>
    </div>
    <div className="dlg-foot">
      {catalog.mode === 1 && <><button type="button" onClick={() => action?.('replaceAll')}>{mnemonic('Replace &All')}</button><button type="button" onClick={() => action?.('replace')}>{mnemonic('&Replace')}</button></>}
      <button type="button" onClick={() => action?.('all')}>{mnemonic('F&ind All')}</button>
      <button type="button" className="primary" onClick={() => action?.('next')}>{mnemonic('&Find Next')}</button>
      <button type="button" onClick={() => action?.('close')}>{t('Close')}</button>
    </div>
  </section>;
}
