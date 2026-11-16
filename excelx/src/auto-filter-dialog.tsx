import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import './auto-filter-dialog.css';

type Condition = { op: string; value: string };
type Catalog = { kind: 'custom' | 'top10' | 'list'; title: string; header?: string; operations?: [string, string][]; entries?: string[]; items?: { value: string; label: string }[]; first?: Condition; second?: Condition; and?: boolean; left?: number; top?: number; minWidth?: number; current?: string };
type Result = { first: Condition; second: Condition; and: boolean } | { top: boolean; percent: boolean; val: number } | string;
type Runtime = { filter?: { reactDialogCatalog?(): Catalog | null; reactDialogAction?(action: string, result?: Result): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** AutoFilter dialogs own only their form state; filtering and criterion conversion stay in the Excelx service. */
export function ExcelxAutoFilterDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [first, setFirst] = useState<Condition>({ op: 'equal', value: '' });
  const [second, setSecond] = useState<Condition>({ op: '', value: '' });
  const [and, setAnd] = useState(true);
  const [top, setTop] = useState(true);
  const [kind, setKind] = useState<'items' | 'percent'>('items');
  const [value, setValue] = useState(10);
  const [active, setActive] = useState(0);
  const listRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.filter?.reactDialogCatalog?.() || null; setCatalog(next); if (next) { setFirst(next.first || { op: 'equal', value: '' }); setSecond(next.second || { op: '', value: '' }); setAnd(next.and !== false); setTop(true); setKind('items'); setValue(10); setActive(Math.max(0, next.items?.findIndex((item) => item.value === next.current) ?? 0)); } };
    const off = runtime()?.bus?.on('excelx-auto-filter-dialog', refresh); return () => off?.();
  }, []);
  useLayoutEffect(() => { if (catalog?.kind === 'list') listRef.current?.focus(); }, [catalog]);
  useEffect(() => {
    if (catalog?.kind !== 'list') return;
    const close = () => runtime()?.filter?.reactDialogAction?.('cancel');
    const pointer = (event: PointerEvent) => { if (listRef.current && !listRef.current.contains(event.target as Node)) close(); };
    const key = (event: KeyboardEvent) => {
      if (!listRef.current?.contains(event.target as Node)) return;
      const items = catalog.items || [];
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(); }
      else if (event.key === 'ArrowDown') { event.preventDefault(); event.stopPropagation(); setActive((index) => Math.min(items.length - 1, index + 1)); }
      else if (event.key === 'ArrowUp') { event.preventDefault(); event.stopPropagation(); setActive((index) => Math.max(0, index - 1)); }
      else if (event.key === 'Enter') { event.preventDefault(); event.stopPropagation(); const item = items[active]; if (item) runtime()?.filter?.reactDialogAction?.('apply', item.value); }
      else if (event.key.length === 1) { const char = event.key.toLowerCase(); const index = items.findIndex((item, i) => i > active && item.label.toLowerCase().startsWith(char)); const next = index >= 0 ? index : items.findIndex((item) => item.label.toLowerCase().startsWith(char)); if (next >= 0) setActive(next); event.preventDefault(); event.stopPropagation(); }
    };
    document.addEventListener('pointerdown', pointer, true); document.addEventListener('keydown', key, true);
    return () => { document.removeEventListener('pointerdown', pointer, true); document.removeEventListener('keydown', key, true); };
  }, [catalog, active]);
  if (!catalog) return null;
  const t = (text: string) => runtime()?.t?.(text) || text;
  const close = () => runtime()?.filter?.reactDialogAction?.('cancel');
  if (catalog.kind === 'list') return <div ref={listRef} id="excelx-auto-filter-list" className="excelx-auto-filter-list" role="listbox" aria-activedescendant={catalog.items?.[active] ? `excelx-auto-filter-list-option-${active}` : undefined} tabIndex={-1} style={{ left: catalog.left, top: catalog.top, minWidth: catalog.minWidth }}>
    {(catalog.items || []).map((item, index) => <div id={`excelx-auto-filter-list-option-${index}`} key={`${index}-${item.value}`} className={`excelx-auto-filter-list__item${item.value === catalog.current ? ' cur' : ''}${index === active ? ' act' : ''}`} role="option" aria-selected={item.value === catalog.current} translate="no" onPointerEnter={() => setActive(index)} onClick={() => runtime()?.filter?.reactDialogAction?.('apply', item.value)}>{item.label}</div>)}
  </div>;
  const submit = () => {
    const result: Result = catalog.kind === 'custom' ? { first, second, and } : { top, percent: kind === 'percent', val: Math.max(1, Math.min(500, Math.round(value) || 1)) };
    runtime()?.filter?.reactDialogAction?.('apply', result);
  };
  const condition = (item: Condition, update: (value: Condition) => void, listId: string) => <div className="excelx-auto-filter-dialog__condition"><select aria-label={t('Comparison')} value={item.op} onChange={(event) => update({ ...item, op: event.currentTarget.value })}><option value=""></option>{catalog.operations?.map(([op, label]) => <option key={op} value={op}>{t(label)}</option>)}</select><input aria-label={t('Value')} list={listId} value={item.value} disabled={!item.op} onChange={(event) => update({ ...item, value: event.currentTarget.value })} /></div>;
  return <OfficeDialog title={t(catalog.title)} className="excelx-auto-filter-dialog" style={{ width: `min(94vw, ${catalog.kind === 'custom' ? 500 : 360}px)` }} onRequestClose={close}>
    {catalog.kind === 'custom' ? <div className="excelx-auto-filter-dialog__body"><div>{t('Show rows where:')}</div><strong translate="no">{catalog.header}</strong>{condition(first, setFirst, 'excelx-filter-values')}
      <div className="excelx-auto-filter-dialog__join">{(['and','or'] as const).map((join) => <label key={join}><input type="radio" name="excelx-filter-join" checked={(join === 'and') === and} onChange={() => setAnd(join === 'and')} />{t(join === 'and' ? '&And' : '&Or').replace('&', '')}</label>)}</div>
      {condition(second, setSecond, 'excelx-filter-values')}<datalist id="excelx-filter-values">{catalog.entries?.map((entry, index) => <option key={`${index}-${entry}`} value={entry} />)}</datalist><p className="hint">{t('Use ? to represent any single character')}{'\n'}{t('Use * to represent any series of characters')}</p>
    </div> : <div className="excelx-auto-filter-dialog__body"><div>{t('Show')}</div><div className="excelx-auto-filter-dialog__top10"><select aria-label={t('Top or bottom')} value={top ? 'top' : 'bottom'} onChange={(event) => setTop(event.currentTarget.value === 'top')}><option value="top">{t('Top')}</option><option value="bottom">{t('Bottom')}</option></select><input aria-label={t('Count')} type="number" min="1" max="500" value={value} onChange={(event) => setValue(Number(event.currentTarget.value))} /><select aria-label={t('Items or percent')} value={kind} onChange={(event) => setKind(event.currentTarget.value as 'items' | 'percent')}><option value="items">{t('Items')}</option><option value="percent">{t('Percent')}</option></select></div></div>}
    <div className="excelx-auto-filter-dialog__actions"><button type="button" className="primary" onClick={submit}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
