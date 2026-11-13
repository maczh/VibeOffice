import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type State = { fixed: boolean; delims: Record<string, boolean | string>; merge: boolean; quote: string; breaks: number[]; formats: string[]; dest: string };
type Catalog = { state: State; delimiterLabels: [string, string][]; formatChoices: [string, string][] };
type Runtime = { dlg?: { reactTextColumnsCatalog?(): Catalog | null; reactTextColumnsPreview?(state: State): string[][]; reactTextColumnsApply?(state: State): boolean; reactTextColumnsClose?(): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

export function ExcelxTextToColumnsDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [state, setState] = useState<State>({ fixed: false, delims: {}, merge: false, quote: '"', breaks: [], formats: [], dest: '' });
  const [preview, setPreview] = useState<string[][]>([]);
  const [selected, setSelected] = useState(0);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactTextColumnsCatalog?.() || null; setCatalog(next); if (next) { setState(next.state); setSelected(0); setPreview(runtime()?.dlg?.reactTextColumnsPreview?.(next.state) || []); } };
    const off = runtime()?.bus?.on('excelx-text-columns-dialog', refresh);
    return () => { off?.(); runtime()?.dlg?.reactTextColumnsClose?.(); };
  }, []);
  if (!catalog) return null;
  const t = (s: string) => runtime()?.t?.(s) || s;
  const update = (patch: Partial<State>) => setState((old) => { const next = { ...old, ...patch }; setPreview(runtime()?.dlg?.reactTextColumnsPreview?.(next) || []); return next; });
  const close = () => runtime()?.dlg?.reactTextColumnsClose?.();
  const columns = Math.max(1, ...preview.map((row) => row.length));
  return <OfficeDialog title={t('Convert Text to Columns Wizard')} className="excelx-text-columns-dialog" style={{ width: 'min(94vw, 680px)' }} onRequestClose={close}>
    <div className="excelx-text-columns-dialog__body">
      <fieldset><legend>{t('Original data type')}</legend><label><input type="radio" checked={!state.fixed} onChange={() => update({ fixed: false })} />{t('&Delimited - Characters such as commas or tabs separate each field.')}</label><label><input type="radio" checked={state.fixed} onChange={() => update({ fixed: true, breaks: state.breaks.length ? state.breaks : [] })} />{t('Fixed &width - Fields are aligned in columns with spaces between each field.')}</label></fieldset>
      {!state.fixed ? <div className="excelx-text-columns-dialog__delimiters">{catalog.delimiterLabels.map(([key, label]) => <label key={key}><input type="checkbox" checked={!!state.delims[key]} onChange={(e) => update({ delims: { ...state.delims, [key]: e.currentTarget.checked } })} />{t(label)}</label>)}<label>{t('&Other:')}<input maxLength={1} value={String(state.delims.other || '')} onChange={(e) => update({ delims: { ...state.delims, other: e.currentTarget.value } })} /></label></div> : <label>{t('Column &breaks at characters:')}<input value={state.breaks.join(', ')} placeholder="e.g. 10, 25, 40" onChange={(e) => update({ breaks: e.currentTarget.value.split(/[ ,]+/).map(Number).filter((x) => x > 0).sort((a, b) => a - b) })} /></label>}
      <div className="excelx-text-columns-dialog__options"><label><input type="checkbox" checked={state.merge} onChange={(e) => update({ merge: e.currentTarget.checked })} />{t('Treat consecutive delimiters as &one')}</label><label>{t('Text &qualifier:')}<select value={state.quote} onChange={(e) => update({ quote: e.currentTarget.value })}><option value={'"'}>{'"'}</option><option value="'">{'\''}</option><option value="">{t('{none}')}</option></select></label></div>
      <fieldset><legend>{t('Column data format')}</legend><div className="ttc-fmts excelx-text-columns-dialog__formats">{catalog.formatChoices.map(([key, label]) => <label key={key}><input type="radio" name="excelx-text-column-format" checked={(state.formats[selected] || 'general') === key} onChange={() => update({ formats: Object.assign([...state.formats], { [selected]: key }) })} />{t(label)}</label>)}</div></fieldset>
      <label>{t('D&estination:')}<input value={state.dest} onChange={(e) => update({ dest: e.currentTarget.value })} /></label>
      <div className="tiw-cap excelx-text-columns-dialog__caption">{t('Data preview:')}</div><div className="tiw-preview excelx-text-columns-dialog__preview"><table className="tiw-table"><thead><tr>{Array.from({ length: columns }, (_, col) => <th key={col} className={selected === col ? 'on' : ''} onClick={() => setSelected(col)}>{t(catalog.formatChoices.find(([key]) => key === (state.formats[col] || 'general'))?.[1] || 'General')}</th>)}</tr></thead><tbody>{preview.map((row, i) => <tr key={i}>{Array.from({ length: columns }, (_, col) => <td key={col} className={selected === col ? 'on' : ''}>{row[col] || ''}</td>)}</tr>)}</tbody></table></div>
    </div><div className="excelx-text-columns-dialog__actions"><span /><button type="button" className="primary" onClick={() => runtime()?.dlg?.reactTextColumnsApply?.(state)}>{t('&Finish').replace('&', '')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
