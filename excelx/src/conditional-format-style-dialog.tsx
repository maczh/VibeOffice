import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type State = { fontStyle: string; underline: string; fontColor: string; strike: boolean; fillColor: string; border: Record<string, boolean> };
type Catalog = { source: Record<string, unknown>; state: State };
type Runtime = { dlg?: { reactCFFormatCatalog?(): Catalog | null; reactCFFormatApply?(state: State): boolean; reactCFFormatClear?(): void; reactCFFormatClose?(): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }
const sides = [['top', 'Top'], ['bottom', 'Bottom'], ['left', 'Left'], ['right', 'Right']] as const;

export function ExcelxConditionalFormatStyleDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [state, setState] = useState<State>({ fontStyle: '', underline: '', fontColor: '', strike: false, fillColor: '', border: {} });
  const [tab, setTab] = useState(0);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactCFFormatCatalog?.() || null; setCatalog(next); if (next) { setState(next.state); setTab(0); } };
    const off = runtime()?.bus?.on('excelx-cf-format-dialog', refresh);
    return () => { off?.(); runtime()?.dlg?.reactCFFormatClose?.(); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const update = (patch: Partial<State>) => setState((old) => ({ ...old, ...patch }));
  const color = (label: string, value: string, set: (value: string) => void) => <label>{t(label)}<input aria-label={label} type="color" value={value || '#000000'} onChange={(event) => set(event.currentTarget.value)} /><button type="button" onClick={() => set('')}>{t('Automatic')}</button></label>;
  const close = () => runtime()?.dlg?.reactCFFormatClose?.();
  return <OfficeDialog title={t('Format Cells')} className="excelx-cf-format-dialog" style={{ width: 'min(94vw, 440px)' }} onRequestClose={close}>
    <div className="excelx-cf-format-dialog__tabs" role="tablist">{['Font', 'Border', 'Patterns'].map((label, index) => <button type="button" role="tab" aria-selected={tab === index} key={label} onClick={() => setTab(index)}>{t(label)}</button>)}</div>
    <div className="excelx-cf-format-dialog__body">
      {tab === 0 && <div className="excelx-cf-format-dialog__font"><label>{t('Font st&yle:')}<select value={state.fontStyle} onChange={(event) => update({ fontStyle: event.currentTarget.value })}><option value="">{t('(unchanged)')}</option><option value="r">{t('Regular')}</option><option value="i">{t('Italic')}</option><option value="b">{t('Bold')}</option><option value="bi">{t('Bold Italic')}</option></select></label><label>{t('&Underline:')}<select value={state.underline} onChange={(event) => update({ underline: event.currentTarget.value })}><option value="">{t('(unchanged)')}</option><option value="none">{t('None')}</option><option value="single">{t('Single')}</option><option value="double">{t('Double')}</option></select></label>{color('&Color:', state.fontColor, (fontColor) => update({ fontColor }))}<label><input type="checkbox" checked={state.strike} onChange={(event) => update({ strike: event.currentTarget.checked })} />{t('Stri&kethrough')}</label></div>}
      {tab === 1 && <div className="excelx-cf-format-dialog__border"><label><input type="checkbox" checked={!!state.border.outline} onChange={(event) => update({ border: { outline: event.currentTarget.checked, top: event.currentTarget.checked, bottom: event.currentTarget.checked, left: event.currentTarget.checked, right: event.currentTarget.checked } })} />{t('&Outline')}</label>{sides.map(([key, label]) => <label key={key}><input type="checkbox" checked={!!state.border[key]} onChange={(event) => { const border = { ...state.border, [key]: event.currentTarget.checked }; border.outline = !!(border.top && border.bottom && border.left && border.right); update({ border }); }} />{t(label)}</label>)}</div>}
      {tab === 2 && <div className="excelx-cf-format-dialog__patterns">{color('Cell shading color:', state.fillColor, (fillColor) => update({ fillColor }))}</div>}
    </div><div className="excelx-cf-format-dialog__actions"><button type="button" onClick={() => runtime()?.dlg?.reactCFFormatClear?.()}>{t('&Clear').replace('&', '')}</button><span /><button type="button" className="primary" onClick={() => runtime()?.dlg?.reactCFFormatApply?.(state)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
