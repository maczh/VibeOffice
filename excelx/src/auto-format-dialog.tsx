import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import type { CSSProperties } from 'react';

type Cell = { value: string; style: string };
type Item = { name: string; rows: Cell[][] };
type Parts = Record<string, boolean>;
type Catalog = { items: Item[]; parts: Parts };
type Runtime = { dlg?: { reactAutoFormatCatalog?(): Catalog | null; reactAutoFormatApply?(name: string, parts: Parts): boolean; reactAutoFormatClose?(): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }
const optionLabels: [string, string][] = [['nf', '&Number'], ['border', '&Border'], ['font', '&Font'], ['fill', '&Patterns'], ['align', '&Alignment'], ['size', '&Width/Height']];
function previewStyle(value: string): CSSProperties { return Object.fromEntries(value.split(';').filter(Boolean).map((rule) => { const [key = '', ...rest] = rule.split(':'); return [key.replace(/-([a-z])/g, (_, letter: string) => letter.toUpperCase()), rest.join(':')]; })); }

export function ExcelxAutoFormatDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [selected, setSelected] = useState('Simple');
  const [parts, setParts] = useState<Parts>({});
  const [showOptions, setShowOptions] = useState(false);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactAutoFormatCatalog?.() || null; setCatalog(next); if (next) { setSelected(next.items.some((item) => item.name === 'Simple') ? 'Simple' : next.items[0]?.name || ''); setParts(next.parts); setShowOptions(false); } };
    const off = runtime()?.bus?.on('excelx-auto-format-dialog', refresh);
    return () => { off?.(); runtime()?.dlg?.reactAutoFormatClose?.(); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactAutoFormatClose?.();
  return <OfficeDialog title={t('AutoFormat')} className="excelx-auto-format-dialog" style={{ width: 'min(94vw, 600px)' }} onRequestClose={close}>
    <div className="afmt-grid">{catalog.items.map((item) => <button type="button" key={item.name} className={`afmt-item${selected === item.name ? ' on' : ''}`} aria-pressed={selected === item.name} onClick={() => setSelected(item.name)}><table className="afmt-sample"><tbody>{item.rows.map((row, r) => <tr key={r}>{row.map((cell, c) => <td key={c} style={previewStyle(cell.style)}>{cell.value}</td>)}</tr>)}</tbody></table><span>{t(item.name)}</span></button>)}</div>
    {showOptions && <div className="excelx-auto-format-dialog__options">{optionLabels.map(([key, label]) => <label key={key}><input type="checkbox" checked={!!parts[key]} onChange={(event) => setParts((old) => ({ ...old, [key]: event.currentTarget.checked }))} />{t(label)}</label>)}</div>}
    <div className="excelx-auto-format-dialog__actions"><button type="button" onClick={() => setShowOptions((shown) => !shown)}>{t('&Options...').replace('&', '')}</button><span /><button type="button" className="primary" onClick={() => runtime()?.dlg?.reactAutoFormatApply?.(selected, parts)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
