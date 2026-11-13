import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import './insert-function-dialog.css';

type Catalog = { categories: string[]; category: string; search: string; names: string[]; selected: string | null; syntax: string; description: string };
type Runtime = { dlg?: { reactFunctionCatalog?(): Catalog | null; reactFunctionAction?(action: string, value?: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React function picker; formula syntax, evaluation, MRU storage and argument editing remain engine-backed. */
export function ExcelxInsertFunctionDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [search, setSearch] = useState('');
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactFunctionCatalog?.() || null; setCatalog(next); if (next) setSearch(next.search); };
    const off = runtime()?.bus?.on('excelx-insert-function-dialog', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactFunctionCatalog?.()) runtime()?.dlg?.reactFunctionAction?.('close'); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = runtime()?.dlg?.reactFunctionAction;
  const open = () => action?.('open');
  const searchNow = () => action?.('search', search);
  return <OfficeDialog title={t('Insert Function')} className="excelx-insert-function-dialog" style={{ width: 'min(94vw, 520px)' }} onRequestClose={() => action?.('close')}>
    <div className="excelx-insert-function-dialog__search">
      <label>{t('Search for a function:')}<div><input value={search} placeholder={t('Type a brief description of what you want to do and then click Go')} onChange={(event) => setSearch(event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); searchNow(); } }} /><button type="button" onClick={searchNow}>{t('&Go').replace('&', '')}</button></div></label>
      <label>{t('Or select a &category:').replace('&', '')}<select value={catalog.category} onChange={(event) => action?.('category', event.currentTarget.value)}>{catalog.categories.map((category) => <option key={category} value={category}>{t(category)}</option>)}{catalog.category === 'Recommended' && <option value="Recommended">{t('Recommended')}</option>}</select></label>
    </div>
    <label className="excelx-insert-function-dialog__list-label">{t('Select a functio&n:').replace('&', '')}<select size={9} aria-label={t('Select a function:')} value={catalog.selected || ''} onChange={(event) => action?.('select', event.currentTarget.value)} onDoubleClick={open}>{catalog.names.map((name) => <option key={name} value={name} translate="no">{name}</option>)}</select></label>
    <div className="if-syntax" translate="no">{catalog.syntax}</div>
    <div className="if-desc">{t(catalog.description)}</div>
    <div className="excelx-insert-function-dialog__actions"><button type="button" className="primary" disabled={!catalog.selected} onClick={open}>{t('OK')}</button><button type="button" onClick={() => action?.('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
