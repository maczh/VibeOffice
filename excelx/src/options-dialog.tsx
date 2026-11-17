import { AppearanceFields, OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState } from 'react';
import './options-dialog.css';

type Catalog = {
  tab: string; tabs: string[]; options: Record<string, boolean | number | string>;
  view: Record<string, boolean>; comments: 'none' | 'indicator' | 'comments';
  calc: { mode: string; iterate: boolean; iterateCount: number; iterateDelta: number; fullPrecision: boolean };
  workbook: { date1904: boolean; r1c1: boolean; hideHScroll: boolean; hideVScroll: boolean; hideTabs: boolean; privacy: boolean };
  general: { sheetsInNew: number; defaultFont: string; defaultSize: number; userName: string; password: string };
  appearance: { look: string; looks: { id: string; name: string }[]; language: string; languages: { id: string; name: string }[] };
  fonts: string[]; sizes: number[]; customLists: string[][];
};
type Runtime = { dlg?: { reactOptionsCatalog?(): Catalog | null; reactOptionsAction?(action: string, catalog?: Catalog): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }
const TABS = ['View', 'Calculation', 'Edit', 'General', 'Custom Lists', 'Error Checking', 'Spelling', 'Security'];
const VIEW_FIELDS: [keyof Catalog['view'], string][] = [['pages','Page brea&ks'],['formulas','Fo&rmulas'],['grid','&Gridlines'],['headings','Row && column h&eaders'],['zeros','&Zero values'],['outlineSymbols','Outline s&ymbols']];

/** Excelx preferences UI; workbook and application settings are committed by the existing app controller. */
export function ExcelxOptionsDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [tab, setTab] = useState('View');
  const id = useId().replaceAll(':', '');
  const [draft, setDraft] = useState<Catalog | null>(null);
  const [listIndex, setListIndex] = useState(-1);
  const [listText, setListText] = useState('');
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactOptionsCatalog?.() || null; setCatalog(next); setDraft(next); if (next) { setTab(next.tab); setListIndex(-1); setListText(''); } };
    const off = runtime()?.bus?.on('excelx-options-dialog', refresh); return () => off?.();
  }, []);
  if (!catalog || !draft) return null;
  const t = (text: string) => runtime()?.t?.(text) || text;
  const action = runtime()?.dlg?.reactOptionsAction;
  const change = (patch: Partial<Catalog>) => setDraft((current) => current ? { ...current, ...patch } : current);
  const setOption = (key: string, value: boolean | number | string) => change({ options: { ...draft.options, [key]: value } });
  const setView = (key: string, value: boolean) => change({ view: { ...draft.view, [key]: value } });
  const setCalc = (key: keyof Catalog['calc'], value: string | number | boolean) => change({ calc: { ...draft.calc, [key]: value } });
  const setWorkbook = (key: keyof Catalog['workbook'], value: boolean) => change({ workbook: { ...draft.workbook, [key]: value } });
  const setGeneral = (key: keyof Catalog['general'], value: string | number) => change({ general: { ...draft.general, [key]: value } });
  const setCheck = (label: string, checked: boolean, onChange: (value: boolean) => void, disabled = false) => <label key={label}><input type="checkbox" checked={checked} disabled={disabled} onChange={(event) => onChange(event.currentTarget.checked)} />{t(label).replace('&', '')}</label>;
  const field = (label: string, control: React.ReactNode) => <label className="excelx-options-dialog__field">{t(label).replace('&', '')}{control}</label>;
  const group = (label: string, children: React.ReactNode) => <fieldset><legend>{t(label)}</legend>{children}</fieldset>;
  const radio = (name: string, value: string, current: string, label: string, onChange: () => void) => <label key={value}><input type="radio" name={name} value={value} checked={current === value} onChange={onChange} />{t(label).replace('&', '')}</label>;
  const select = (label: string, value: string | number, items: [string | number, string][], onChange: (value: string) => void) => field(label, <select aria-label={t(label).replace('&', '')} value={value} onChange={(event) => onChange(event.currentTarget.value)}>{items.map(([itemValue, itemLabel]) => <option key={itemValue} value={itemValue}>{t(itemLabel)}</option>)}</select>);
  const viewBody = <div className="excelx-options-dialog__columns"><div>
    {group('Show', <>{setCheck('&Startup Task Pane', !!draft.options.taskOpen, (value) => setOption('taskOpen', value))}{setCheck('Task pane f&ollows the selected object', !!draft.options.paneFollows, (value) => setOption('paneFollows', value))}{setCheck('&Formula bar', !!draft.options.formulaBar, (value) => setOption('formulaBar', value))}{setCheck('St&atus bar', !!draft.options.statusBar, (value) => setOption('statusBar', value))}</>)}
    {group('Comments', <>{radio('excelx-options-comments', 'none', draft.comments, '&None', () => change({ comments: 'none' }))}{radio('excelx-options-comments', 'indicator', draft.comments, 'Comment &indicator only', () => change({ comments: 'indicator' }))}{radio('excelx-options-comments', 'comments', draft.comments, 'Comment && indi&cator', () => change({ comments: 'comments' }))}</>)}
  </div><div>{group('Window options', <div className="excelx-options-dialog__window">{VIEW_FIELDS.map(([key,label]) => setCheck(label, !!draft.view[key], (value) => setView(key, value)))}{setCheck('Hori&zontal scroll bar', !draft.workbook.hideHScroll, (value) => setWorkbook('hideHScroll', !value))}{setCheck('&Vertical scroll bar', !draft.workbook.hideVScroll, (value) => setWorkbook('hideVScroll', !value))}{setCheck('She&et tabs', !draft.workbook.hideTabs, (value) => setWorkbook('hideTabs', !value))}</div>)}</div></div>;
  const calcBody = <div className="excelx-options-dialog__columns"><div>{group('Calculation', <>{radio('excelx-options-calc', 'auto', draft.calc.mode === 'autoNoTable' ? 'autoNoTable' : draft.calc.mode, '&Automatic', () => setCalc('mode', 'auto'))}{radio('excelx-options-calc', 'autoNoTable', draft.calc.mode, 'Automatic except &tables', () => setCalc('mode', 'autoNoTable'))}{radio('excelx-options-calc', 'manual', draft.calc.mode, '&Manual', () => setCalc('mode', 'manual'))}<button type="button" onClick={() => action?.('calculate')}>{t('Calc &Now (F9)').replace('&', '')}</button></>)}</div><div>
    {group('Iteration', <>{setCheck('&Iteration', draft.calc.iterate, (value) => setCalc('iterate', value))}{field('Maximum iteratio&ns:', <input type="number" min="1" max="32767" value={draft.calc.iterateCount} onChange={(event) => setCalc('iterateCount', Number(event.currentTarget.value))} />)}{field('Maximum &change:', <input type="number" min="0" step="any" value={draft.calc.iterateDelta} onChange={(event) => setCalc('iterateDelta', Number(event.currentTarget.value))} />)}</>)}
    {group('Workbook options', <>{setCheck('&Precision as displayed', !draft.calc.fullPrecision, (value) => setCalc('fullPrecision', !value))}{setCheck('&1904 date system', draft.workbook.date1904, (value) => setWorkbook('date1904', value))}{setCheck('Accept la&bels in formulas', false, () => {}, true)}</>)}
  </div></div>;
  const editBody = <div className="excelx-options-dialog__column">{setCheck('Edit directly in c&ell', !!draft.options.editInCell, (value) => setOption('editInCell', value))}{setCheck('Allow cell &drag and drop', !!draft.options.dragDrop, (value) => setOption('dragDrop', value))}<div className="excelx-options-dialog__row">{setCheck('&Move selection after Enter', !!draft.options.moveAfterEnter, (value) => setOption('moveAfterEnter', value))}{select('D&irection:', String(draft.options.enterDir || 'down'), [['down','Down'],['right','Right'],['up','Up'],['left','Left']], (value) => setOption('enterDir', value))}</div>{setCheck('Enable Auto&Complete for cell values', !!draft.options.autoComplete, (value) => setOption('autoComplete', value))}{setCheck('Extend data range &formats and formulas', !!draft.options.extendFormats, (value) => setOption('extendFormats', value))}{setCheck('Enable automatic percent entry', !!draft.options.autoPercent, (value) => setOption('autoPercent', value))}</div>;
  const generalBody = <div className="excelx-options-dialog__column"><AppearanceFields className="excelx-options-dialog__field" look={draft.appearance.look} looks={draft.appearance.looks.map((item) => ({ value: item.id, label: item.name }))} language={draft.appearance.language} languages={draft.appearance.languages.map((item) => ({ value: item.id, label: item.name }))} onLookChange={(value) => change({ appearance: { ...draft.appearance, look: value } })} onLanguageChange={(value) => change({ appearance: { ...draft.appearance, language: value } })} translate={t} />{setCheck('R1C1 reference st&yle', draft.workbook.r1c1, (value) => setWorkbook('r1c1', value))}{field('Sheets in new work&book:', <input type="number" min="1" max="255" value={draft.general.sheetsInNew} onChange={(event) => setGeneral('sheetsInNew', Number(event.currentTarget.value))} />)}<div className="excelx-options-dialog__row">{select('Standard fon&t:', draft.general.defaultFont, draft.fonts.map((font) => [font,font]), (value) => setGeneral('defaultFont', value))}{select('Si&ze:', draft.general.defaultSize, draft.sizes.map((size) => [size,String(size)]), (value) => setGeneral('defaultSize', Number(value)))}</div>{field('User &name:', <input value={draft.general.userName} onChange={(event) => setGeneral('userName', event.currentTarget.value)} />)}</div>;
  const customListBody = <div className="excelx-options-dialog__custom"><label>{t('Custom lists:')}<select size={8} aria-label={t('Custom lists:')} value={listIndex} onChange={(event) => { const index = Number(event.currentTarget.value); setListIndex(index); setListText(index >= 0 ? draft.customLists[index]?.join('\n') || '' : ''); }}>{<option value={-1}>{t('NEW LIST')}</option>}{draft.customLists.map((items, index) => <option key={index} value={index}>{items.slice(0,4).join(', ')}{items.length > 4 ? ', …' : ''}</option>)}</select></label><label>{t('List entries:')}<textarea rows={8} value={listText} readOnly={listIndex >= 0 && listIndex < 4} onChange={(event) => setListText(event.currentTarget.value)} /></label><div><button type="button" onClick={() => { const items = listText.split(/[\n,]+/).map((item) => item.trim()).filter(Boolean); if (!items.length) return; const lists = draft.customLists.map((entry) => [...entry]); if (listIndex >= 4) lists[listIndex] = items; else lists.push(items); change({ customLists: lists }); setListIndex(-1); setListText(''); }}>{t('&Add').replace('&', '')}</button><button type="button" disabled={listIndex < 4} onClick={() => { const lists = draft.customLists.filter((_, index) => index !== listIndex).map((entry) => [...entry]); change({ customLists: lists }); setListIndex(-1); setListText(''); }}>{t('&Delete').replace('&', '')}</button></div></div>;
  const errorBody = <div className="excelx-options-dialog__column">{setCheck('&Enable background error checking', !!draft.options.bgErrors, (value) => setOption('bgErrors', value))}{setCheck('Evaluates to error &value', true, () => {}, true)}{setCheck('Number stored as &text', !!draft.options.errNumText, (value) => setOption('errNumText', value))}{setCheck('&Inconsistent formula in region', !!draft.options.errInconsistent, (value) => setOption('errInconsistent', value))}</div>;
  const spellingBody = <div className="excelx-options-dialog__column">{select('&Dictionary language:', String(draft.options.dictLang || 'en_US'), [['en_US','English (U.S.)'],['en_GB','English (U.K.)']], (value) => setOption('dictLang', value))}{setCheck('Ignore words in &UPPERCASE', !!draft.options.ignoreUpper, (value) => setOption('ignoreUpper', value))}{setCheck('Ignore words with &numbers', !!draft.options.ignoreNum, (value) => setOption('ignoreNum', value))}{setCheck('Ignore &Internet and file addresses', !!draft.options.ignoreInternet, (value) => setOption('ignoreInternet', value))}<button type="button" onClick={() => action?.('autocorrect')}>{t('&AutoCorrect Options...').replace('&', '')}</button></div>;
  const securityBody = <div className="excelx-options-dialog__column">{group('File encryption settings for this workbook', <>{field('Password to &open:', <input type="password" autoComplete="new-password" value={draft.general.password} onChange={(event) => setGeneral('password', event.currentTarget.value)} />)}<p className="hint">{t('The workbook is encrypted with AES-256 the next time you save it as an Excel workbook, the way Excel 2007 and later do. Leave the box empty to save it without a password. If you lose or forget the password, it cannot be recovered.')}</p></>)}{group('Privacy options', setCheck('&Remove personal information from file properties on save', draft.workbook.privacy, (value) => setWorkbook('privacy', value)))}</div>;
  const bodies: Record<string, React.ReactNode> = { View: viewBody, Calculation: calcBody, Edit: editBody, General: generalBody, 'Custom Lists': customListBody, 'Error Checking': errorBody, Spelling: spellingBody, Security: securityBody };
  const tabId = (index: number) => `${id}-options-tab-${index}`;
  const panelId = (index: number) => `${id}-options-panel-${index}`;
  return <OfficeDialog title={t('Options')} className="excelx-options-dialog" style={{ width: 'min(96vw, 620px)' }} onRequestClose={() => action?.('cancel')}>
    <div className="excelx-options-dialog__tabs" role="tablist" aria-label={t('Options')}>{TABS.map((item, index) => <button key={item} id={tabId(index)} type="button" role="tab" aria-controls={panelId(index)} aria-selected={tab === item} tabIndex={tab === item ? 0 : -1} onClick={() => setTab(item)} onKeyDown={(event) => {
      const current = TABS.indexOf(item);
      const next = event.key === 'ArrowRight' ? (current + 1) % TABS.length
        : event.key === 'ArrowLeft' ? (current - 1 + TABS.length) % TABS.length
          : event.key === 'Home' ? 0
            : event.key === 'End' ? TABS.length - 1 : -1;
      if (next < 0) return;
      event.preventDefault();
      const nextTab = TABS[next];
      if (nextTab) setTab(nextTab);
      event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
    }}>{t(item)}</button>)}</div>
    {TABS.map((item, index) => <div key={item} id={panelId(index)} aria-labelledby={tabId(index)} role="tabpanel" className="excelx-options-dialog__panel" hidden={tab !== item}>{bodies[item]}</div>)}
    <div className="excelx-options-dialog__actions"><button type="button" className="primary" onClick={() => action?.('save', draft)}>{t('OK')}</button><button type="button" onClick={() => action?.('cancel')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
