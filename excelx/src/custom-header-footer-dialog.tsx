import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useRef, useState } from 'react';

type Sections = { l: string; c: string; r: string };
type Catalog = { kind: string; sections: Sections };
type Runtime = { dlg?: { reactCustomHFCatalog?(): Catalog | null; reactCustomHFApply?(sections: Sections): boolean; reactCustomHFClose?(): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }
const fields: [string, string, string][] = [['Page #', '&P', 'page'], ['Pages', '&N', 'pages'], ['Date', '&D', 'date'], ['Time', '&T', 'time'], ['Path & File', '&Z&F', 'path and file'], ['File', '&F', 'file'], ['Tab', '&A', 'tab']];

export function ExcelxCustomHeaderFooterDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [sections, setSections] = useState<Sections>({ l: '', c: '', r: '' });
  const [active, setActive] = useState<keyof Sections>('c');
  const [fontOpen, setFontOpen] = useState(false);
  const [font, setFont] = useState({ name: 'Arial', style: 'Regular', size: '10' });
  const areas = useRef<Partial<Record<keyof Sections, HTMLTextAreaElement | null>>>({});
  const selection = useRef<Record<keyof Sections, [number, number]>>({ l: [0, 0], c: [0, 0], r: [0, 0] });
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactCustomHFCatalog?.() || null; setCatalog(next); if (next) { setSections(next.sections); setActive('c'); setFontOpen(false); } };
    const off = runtime()?.bus?.on('excelx-custom-hf-dialog', refresh);
    return () => { off?.(); runtime()?.dlg?.reactCustomHFClose?.(); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactCustomHFClose?.();
  const insert = (token: string) => {
    const area = areas.current[active]; if (!area) return;
    const [start, end] = selection.current[active], value = sections[active];
    const next = { ...sections, [active]: value.slice(0, start) + token + value.slice(end) };
    selection.current[active] = [start + token.length, start + token.length];
    setSections(next);
    requestAnimationFrame(() => { const input = areas.current[active]; input?.focus(); input?.setSelectionRange(start + token.length, start + token.length); });
  };
  const formatFont = () => { const size = Number.parseInt(font.size, 10) || 10; insert(`&"${font.name},${font.style}"&${size}`); setFontOpen(false); };
  return <><OfficeDialog title={t(catalog.kind)} className="excelx-custom-hf-dialog" style={{ width: 'min(94vw, 600px)' }} onRequestClose={close}>
    <div className="excelx-custom-hf-dialog__body"><p className="hint">{t('To format text: select the text, then choose the font button. To insert a page number, date, time, file path, filename, or tab name: position the insertion point in the edit box, then choose the appropriate button.')}</p><div className="excelx-custom-hf-dialog__fields"><button type="button" className="btn small" onClick={() => setFontOpen(true)}>{t('A (Font)')}</button>{fields.map(([label, token]) => <button type="button" className="btn small" key={token} onClick={() => insert(token)}>{t(label)}</button>)}</div><div className="hf-secs">{([['l', '&Left section:'], ['c', '&Center section:'], ['r', '&Right section:']] as const).map(([key, label]) => <label key={key}>{t(label)}<textarea ref={(node) => { areas.current[key] = node; }} rows={5} value={sections[key]} onFocus={(event) => { setActive(key); selection.current[key] = [event.currentTarget.selectionStart, event.currentTarget.selectionEnd]; }} onSelect={(event) => { selection.current[key] = [event.currentTarget.selectionStart, event.currentTarget.selectionEnd]; }} onChange={(event) => { const value = event.currentTarget.value, start = event.currentTarget.selectionStart, end = event.currentTarget.selectionEnd; selection.current[key] = [start, end]; setSections((old) => ({ ...old, [key]: value })); }} /></label>)}</div><div className="excelx-custom-hf-dialog__actions"><span /><button type="button" className="primary" onClick={() => runtime()?.dlg?.reactCustomHFApply?.(sections)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div></div>
  </OfficeDialog>{fontOpen && <OfficeDialog title={t('Font')} className="excelx-custom-hf-font-dialog" style={{ width: 'min(94vw, 360px)' }} onRequestClose={() => setFontOpen(false)}><div className="excelx-custom-hf-dialog__font"><label>{t('Font:')}<input value={font.name} onChange={(event) => { const name = event.currentTarget.value; setFont((old) => ({ ...old, name })); }} /></label><label>{t('Style:')}<select value={font.style} onChange={(event) => { const style = event.currentTarget.value; setFont((old) => ({ ...old, style })); }}><option>Regular</option><option>Bold</option><option>Italic</option><option>Bold Italic</option></select></label><label>{t('Size:')}<input type="number" value={font.size} onChange={(event) => { const size = event.currentTarget.value; setFont((old) => ({ ...old, size })); }} /></label><div className="excelx-custom-hf-dialog__actions"><button type="button" onClick={() => setFontOpen(false)}>{t('Cancel')}</button><span /><button type="button" className="primary" onClick={formatFont}>{t('OK')}</button></div></div></OfficeDialog>}</>;
}
