import { OfficeDialog } from '@vibeoffice/core';
import { createPortal } from 'react-dom';
import { useEffect, useRef, useState, type KeyboardEvent as ReactKeyboardEvent, type MouseEvent } from 'react';
import './format-cells-dialog.css';

interface FontColor { hex: string; automatic: boolean }
interface NumberOptions { kind: 'number' | 'currency' | 'accounting' | 'percentage' | 'scientific' | 'catalog' | 'description' | 'custom'; dec?: number; sep?: boolean; neg?: number; sym?: string; formats?: { value: string; label: string }[]; selected?: string; description?: string; locale?: boolean; code?: string; canDelete?: boolean }
interface Fill { fillColor: string | null; pattern: string | null; patColor: string }
interface BorderEdge { style: string; color: string }
interface BorderState { edges: Record<string, BorderEdge | null>; style: string; color: string | null; multiR: boolean; multiC: boolean }
interface Catalog { title: string; content: HTMLElement; protection: { locked: boolean; hidden: boolean }; alignment: Record<string, unknown>; merged: boolean; numberCategories: { value: string; label: string }[]; numberCategory: string; numberOptions: NumberOptions | null; symbols: { label: string; value: string }[]; font: Record<string, unknown>; fontName: string; fontColors: FontColor; fonts: string[]; fontSizes: number[]; fill: Fill; fillColors: { value: string; label: string }[]; patterns: { value: string; label: string; preview: string }[]; border: BorderState; borderLines: { style: string; label: string }[] }
interface Runtime {
  dlg?: { reactFormatCellsCatalog?(): Catalog | null; reactFormatCellsSetProtection?(locked: boolean, hidden: boolean): boolean; reactFormatCellsSetAlignment?(patch: Record<string, unknown>): boolean; reactFormatCellsSetNumberCategory?(category: string): NumberOptions | null | false; reactFormatCellsSetNumberOptions?(patch: Record<string, unknown>): NumberOptions | false; reactFormatCellsDeleteNumberFormat?(code: string): boolean; reactFormatCellsSetFont?(patch: Record<string, unknown>): false | { font: Record<string, unknown>; fontName: string; fontColors: FontColor }; reactFormatCellsChooseFontColor?(anchor: HTMLElement): boolean; reactFormatCellsSetFill?(patch: Partial<Fill>): Fill | false; reactFormatCellsPatternURL?(pattern: string, foreground: string, background: string): string; reactFormatCellsSetBorder?(patch: Record<string, unknown>): BorderState | false; reactFormatCellsChooseBorderColor?(anchor: HTMLElement): boolean; reactFormatCellsApply?(): boolean; reactFormatCellsClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React-owned lifecycle and dialog shell for Excelx's full six-tab cell formatter. */
export function ExcelxFormatCellsDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [protectionHost, setProtectionHost] = useState<HTMLElement | null>(null);
  const [alignmentHost, setAlignmentHost] = useState<HTMLElement | null>(null);
  const [numberHost, setNumberHost] = useState<HTMLElement | null>(null);
  const [fontHost, setFontHost] = useState<HTMLElement | null>(null);
  const [numberOptionsHost, setNumberOptionsHost] = useState<HTMLElement | null>(null);
  const [fillHost, setFillHost] = useState<HTMLElement | null>(null);
  const [borderHost, setBorderHost] = useState<HTMLElement | null>(null);
  const [protection, setProtection] = useState({ locked: true, hidden: false });
  const [alignment, setAlignment] = useState<Record<string, unknown>>({});
  const [merged, setMerged] = useState(false);
  const [numberCategory, setNumberCategory] = useState('General');
  const [numberOptions, setNumberOptions] = useState<NumberOptions | null>(null);
  const [customFormatInput, setCustomFormatInput] = useState('');
  const [fill, setFill] = useState<Fill>({ fillColor: null, pattern: null, patColor: '#000000' });
  const [patternMenuOpen, setPatternMenuOpen] = useState(false);
  const [border, setBorder] = useState<BorderState>({ edges: {}, style: 'thin', color: null, multiR: false, multiC: false });
  const borderColorButton = useRef<HTMLButtonElement>(null);
  const [font, setFont] = useState<Record<string, unknown>>({});
  const [fontName, setFontName] = useState('Arial');
  const [fontColors, setFontColors] = useState<FontColor>({ hex: '#000000', automatic: true });
  const [fontNameInput, setFontNameInput] = useState('Arial');
  const [fontSizeInput, setFontSizeInput] = useState('10');
  const colorButton = useRef<HTMLButtonElement>(null);
  const contentHost = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactFormatCellsCatalog?.() || null;
      setCatalog(next);
      if (next) { setProtection(next.protection); setAlignment(next.alignment); setMerged(next.merged); setNumberCategory(next.numberCategory); setNumberOptions(next.numberOptions); setCustomFormatInput(next.numberOptions?.kind === 'custom' ? String(next.numberOptions.code || '') : ''); setFont(next.font); setFontName(next.fontName); setFontNameInput(next.fontName); setFontSizeInput(String(next.font.sz || 10)); setFontColors(next.fontColors); setFill(next.fill); setPatternMenuOpen(false); setBorder(next.border); }
    };
    const off = runtime()?.bus?.on('excelx-format-cells-dialog', refresh);
    const offColor = runtime()?.bus?.on('excelx-format-cells-font-color', () => {
      const current = runtime()?.dlg?.reactFormatCellsCatalog?.();
      if (current) { setFont(current.font); setFontColors(current.fontColors); }
    });
    const offNumber = runtime()?.bus?.on('excelx-format-cells-number-options', () => {
      const current = runtime()?.dlg?.reactFormatCellsCatalog?.();
      if (current) { setNumberOptions(current.numberOptions); setCustomFormatInput(current.numberOptions?.kind === 'custom' ? String(current.numberOptions.code || '') : ''); }
    });
    const offBorder = runtime()?.bus?.on('excelx-format-cells-border-color', () => { const current = runtime()?.dlg?.reactFormatCellsCatalog?.(); if (current) setBorder(current.border); });
    return () => { off?.(); offColor?.(); offNumber?.(); offBorder?.(); runtime()?.dlg?.reactFormatCellsClose?.(); };
  }, []);
  useEffect(() => {
    if (!catalog || !contentHost.current) return;
    contentHost.current.replaceChildren(catalog.content);
    return () => { catalog.content.remove(); };
  }, [catalog]);
  useEffect(() => {
    if (!catalog) { setProtectionHost(null); setAlignmentHost(null); setNumberHost(null); setFontHost(null); setNumberOptionsHost(null); setFillHost(null); setBorderHost(null); return; }
    setProtectionHost(catalog.content.querySelector<HTMLElement>('.fc-prot'));
    setAlignmentHost(catalog.content.querySelector<HTMLElement>('.fc-align'));
    setNumberHost(catalog.content.querySelector<HTMLElement>('.fc-cats'));
    setFontHost(catalog.content.querySelector<HTMLElement>('.fc-font'));
    setNumberOptionsHost(catalog.content.querySelector<HTMLElement>('.fc-react-options'));
    setFillHost(catalog.content.querySelector<HTMLElement>('.fc-pat'));
    setBorderHost(catalog.content.querySelector<HTMLElement>('.fc-border'));
  }, [catalog]);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const setAlignmentValue = (key: string, value: unknown) => {
    const patch = { [key]: value };
    setAlignment((current) => ({ ...current, ...patch }));
    runtime()?.dlg?.reactFormatCellsSetAlignment?.(patch);
  };
  const rotation = alignment.rot === 255 ? 0 : typeof alignment.rot === 'number' && alignment.rot > 90 ? 90 - alignment.rot : Number(alignment.rot || 0);
  const setNumberCategoryValue = (value: string) => { setNumberCategory(value); const next = runtime()?.dlg?.reactFormatCellsSetNumberCategory?.(value); setNumberOptions(next === false ? null : next || null); };
  const updateNumberOptions = (patch: Record<string, unknown>) => { const next = runtime()?.dlg?.reactFormatCellsSetNumberOptions?.(patch); if (next !== false && next) { setNumberOptions(next); if (next.kind === 'custom') setCustomFormatInput(String(next.code || '')); } };
  const updateFont = (patch: Record<string, unknown>) => {
    const updated = runtime()?.dlg?.reactFormatCellsSetFont?.(patch);
    if (!updated) return;
    setFont(updated.font); setFontName(updated.fontName); setFontNameInput(updated.fontName); setFontSizeInput(String(updated.font.sz || 10)); setFontColors(updated.fontColors);
  };
  const updateFill = (patch: Partial<Fill>) => {
    const next = runtime()?.dlg?.reactFormatCellsSetFill?.(patch);
    if (next && typeof next === 'object') setFill(next);
  };
  const patternPreview = runtime()?.dlg?.reactFormatCellsPatternURL;
  const updateBorder = (patch: Record<string, unknown>) => { const next = runtime()?.dlg?.reactFormatCellsSetBorder?.(patch); if (next && typeof next === 'object') setBorder(next); };
  const borderStroke = (edge: BorderEdge | null | undefined, x1: number, y1: number, x2: number, y2: number) => {
    if (!edge) return null;
    const width = /medium|slant/i.test(edge.style) ? 2 : edge.style === 'thick' ? 3 : edge.style === 'hair' ? 0.5 : 1;
    const dash = ({ hair: '1 1', dotted: '1 2', dashDotDot: '6 2 1 2 1 2', dashDot: '6 2 1 2', dashed: '4 2', mediumDashDotDot: '6 2 1 2 1 2', slantDashDot: '7 2 2 2', mediumDashDot: '6 2 2 2', mediumDashed: '6 3' } as Record<string, string>)[edge.style];
    if (edge.style === 'double') return <g><line x1={x1} y1={y1} x2={x2} y2={y2} stroke={edge.color} strokeWidth="3" /><line x1={x1} y1={y1} x2={x2} y2={y2} stroke="#fff" strokeWidth="1" /></g>;
    return <line x1={x1} y1={y1} x2={x2} y2={y2} stroke={edge.color} strokeWidth={width} {...(dash ? { strokeDasharray: dash } : {})} />;
  };
  const toggleBorderAt = (event: MouseEvent<SVGSVGElement>) => {
    const rect = event.currentTarget.getBoundingClientRect(); const x = (event.clientX - rect.left) / rect.width * 150; const y = (event.clientY - rect.top) / rect.height * 110;
    const distances: Record<string, number> = { top: Math.abs(y - 10), bottom: Math.abs(y - 100), left: Math.abs(x - 15), right: Math.abs(x - 135) };
    if (border.multiR) distances.insideH = Math.abs(y - 55);
    if (border.multiC) distances.insideV = Math.abs(x - 75);
    const edge = Object.keys(distances).sort((a, b) => (distances[a] ?? Infinity) - (distances[b] ?? Infinity))[0];
    if (edge && (distances[edge] ?? Infinity) < 12) updateBorder({ type: 'toggle', edge });
  };
  const selectNumberCategory = (offset: number) => {
    const current = catalog.numberCategories.findIndex((item) => item.value === numberCategory);
    const next = Math.max(0, Math.min(catalog.numberCategories.length - 1, current + offset));
    const item = catalog.numberCategories[next];
    if (item) setNumberCategoryValue(item.value);
  };
  const activeNumberCategoryIndex = Math.max(0, catalog.numberCategories.findIndex((item) => item.value === numberCategory));
  const onNumberCategoryKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
    event.preventDefault();
    event.stopPropagation();
    if (event.key === 'Home') { const first = catalog.numberCategories[0]; if (first) setNumberCategoryValue(first.value); }
    else if (event.key === 'End') { const last = catalog.numberCategories.at(-1); if (last) setNumberCategoryValue(last.value); }
    else selectNumberCategory(event.key === 'ArrowDown' ? 1 : -1);
  };
  const onNumberFormatKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const formats = numberOptions?.formats || [];
    if (!formats.length || !['ArrowDown', 'ArrowUp', 'Home', 'End', 'Enter', ' '].includes(event.key)) return;
    event.preventDefault();
    event.stopPropagation();
    const current = Math.max(0, formats.findIndex((format) => format.value === numberOptions?.selected));
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? formats.length - 1 : event.key === 'ArrowDown' ? Math.min(formats.length - 1, current + 1) : event.key === 'ArrowUp' ? Math.max(0, current - 1) : current;
    const format = formats[next];
    if (format) updateNumberOptions({ code: format.value });
  };
  const close = () => runtime()?.dlg?.reactFormatCellsClose?.();
  return <OfficeDialog title={translate(catalog.title)} className="excelx-format-cells-dialog" style={{ width: 'min(96vw, 520px)' }} onRequestClose={close}>
    <div ref={contentHost} className="excelx-format-cells-dialog__content" />
    {numberHost && createPortal(<div className="excelx-format-cells-dialog__number-category">
      <label>{translate('Category:')}</label>
      <div className="lbox excelx-format-cells-dialog__category-list" role="listbox" aria-label={translate('Category:')} aria-activedescendant={`excelx-format-cells-category-${activeNumberCategoryIndex}`} tabIndex={0} onKeyDown={onNumberCategoryKeyDown}>
        {catalog.numberCategories.map((item, index) => <div id={`excelx-format-cells-category-${index}`} key={item.value} className={`lbox-i${numberCategory === item.value ? ' on' : ''}`} role="option" aria-selected={numberCategory === item.value} onClick={() => setNumberCategoryValue(item.value)}>{translate(item.label)}</div>)}
      </div>
    </div>, numberHost)}
    {numberOptionsHost && createPortal(<div className="excelx-format-cells-dialog__number-options" hidden={!numberOptions}>
      {numberOptions && <>
      {['number','currency','accounting','percentage','scientific'].includes(numberOptions.kind) && <label>{translate('&Decimal places:').replace('&', '')}<input aria-label={translate('&Decimal places:').replace('&', '')} type="number" min="0" max="30" step="1" value={numberOptions.dec || 0} onChange={(event) => updateNumberOptions({ dec: Number(event.currentTarget.value) })} /></label>}
      {numberOptions.kind === 'catalog' && <>
        <label>{translate('Type:')}</label>
        <div className="lbox short excelx-format-cells-dialog__number-format-list" role="listbox" aria-label={translate('Type:')} aria-activedescendant={`excelx-format-cells-number-format-${Math.max(0, (numberOptions.formats || []).findIndex((format) => format.value === numberOptions.selected))}`} tabIndex={0} onKeyDown={onNumberFormatKeyDown}>{(numberOptions.formats || []).map((format, index) => <div id={`excelx-format-cells-number-format-${index}`} key={format.value} data-format={format.value} className={`lbox-i${numberOptions.selected === format.value ? ' on' : ''}`} role="option" aria-selected={numberOptions.selected === format.value} onClick={() => updateNumberOptions({ code: format.value })}>{translate(format.label)}</div>)}</div>
        {numberOptions.locale && <label>{translate('&Locale (location):').replace('&', '')}<select aria-label={translate('&Locale (location):').replace('&', '')}><option value="en-US">{translate('English (United States)')}</option></select></label>}
        {numberOptions.description && <p className="hint">{translate(numberOptions.description)}</p>}
      </>}
      {numberOptions.kind === 'description' && numberOptions.description && <p className="hint">{translate(numberOptions.description)}</p>}
      {numberOptions.kind === 'custom' && <>
        <label>{translate('&Type:').replace('&', '')}<input aria-label={translate('&Type:').replace('&', '')} type="text" value={customFormatInput} onChange={(event) => { setCustomFormatInput(event.currentTarget.value); updateNumberOptions({ code: event.currentTarget.value }); }} /></label>
        <div className="lbox short excelx-format-cells-dialog__custom-format-list" role="listbox" aria-label={translate('Existing format codes')}>{(numberOptions.formats || []).map((format) => <div key={format.value} data-format={format.value} translate="no" className={`lbox-i${numberOptions.selected === format.value ? ' on' : ''}`} role="option" aria-selected={numberOptions.selected === format.value} onClick={() => updateNumberOptions({ code: format.value })}>{format.label}</div>)}</div>
        <div className="row"><button type="button" disabled={!numberOptions.canDelete} onClick={() => { if (numberOptions.selected) runtime()?.dlg?.reactFormatCellsDeleteNumberFormat?.(numberOptions.selected); }}>{translate('&Delete').replace('&', '')}</button></div>
        {numberOptions.description && <p className="hint">{translate(numberOptions.description)}</p>}
      </>}
      {(numberOptions.kind === 'number' || numberOptions.kind === 'currency') && <>
        {numberOptions.kind === 'number' && <label><input type="checkbox" aria-label={translate('&Use 1000 Separator (,)').replace('&', '')} checked={!!numberOptions.sep} onChange={(event) => updateNumberOptions({ sep: event.currentTarget.checked })} />{translate('&Use 1000 Separator (,)').replace('&', '')}</label>}
        {numberOptions.kind === 'currency' && <label>{translate('&Symbol:').replace('&', '')}<select aria-label={translate('&Symbol:').replace('&', '')} value={numberOptions.sym || ''} onChange={(event) => updateNumberOptions({ sym: event.currentTarget.value })}>{catalog.symbols.map((symbol) => <option key={symbol.value} value={symbol.value}>{translate(symbol.label)}</option>)}</select></label>}
        <label>{translate('Negative numbers:')}</label>
        <div className="lbox short excelx-format-cells-dialog__negative-list" role="listbox" aria-label={translate('Negative numbers:')}>{[0,1,2,3].map((value) => { const currency = numberOptions.kind === 'currency' ? (numberOptions.sym || '').replaceAll('"', '') : ''; const sample = `${currency}${numberOptions.sep || numberOptions.kind === 'currency' ? '1,234' : '1234'}${numberOptions.dec ? '.' + '1'.repeat(Math.min(numberOptions.dec, 3)) : ''}`; const label = value === 2 ? `(${sample})` : value === 3 ? `(${sample}) Red` : `-${sample}${value === 1 ? ' Red' : ''}`; return <div key={value} className={`lbox-i${numberOptions.neg === value ? ' on' : ''}`} role="option" aria-selected={numberOptions.neg === value} onClick={() => updateNumberOptions({ neg: value })}>{label}</div>; })}</div>
      </>}
      {numberOptions.kind === 'accounting' && <label>{translate('&Symbol:').replace('&', '')}<select aria-label={translate('&Symbol:').replace('&', '')} value={numberOptions.sym || ''} onChange={(event) => updateNumberOptions({ sym: event.currentTarget.value })}>{catalog.symbols.map((symbol) => <option key={symbol.value} value={symbol.value}>{translate(symbol.label)}</option>)}</select></label>}
      </>}
    </div>, numberOptionsHost)}
    {fontHost && createPortal(<div className="excelx-format-cells-dialog__font">
      <div className="fc-f3">
        <div className="col"><label>{translate('Font:')}</label><input aria-label={translate('Font:')} type="text" value={fontNameInput} onChange={(event) => { setFontNameInput(event.currentTarget.value); updateFont({ name: event.currentTarget.value }); }} onBlur={() => setFontNameInput(fontName)} />
          <div className="lbox short" role="listbox" aria-label={translate('Font:')}>{catalog.fonts.map((name) => <div className={`lbox-i${fontName === name ? ' on' : ''}`} role="option" aria-selected={fontName === name} key={name} onClick={() => updateFont({ name })}>{name}</div>)}</div>
        </div>
        <div className="col"><label>{translate('Font style:')}</label><div className="lbox short" role="listbox" aria-label={translate('Font style:')}>{([['r','Regular'],['i','Italic'],['b','Bold'],['bi','Bold Italic']] as const).map(([key,label]) => { const selected = (font.b ? 'b' : '') + (font.i ? 'i' : '') || 'r'; return <div className={`lbox-i${selected === key ? ' on' : ''}`} role="option" aria-selected={selected === key} key={key} onClick={() => updateFont({ b: /b/.test(key), i: /i/.test(key) })}>{translate(label)}</div>; })}</div></div>
        <div className="col"><label>{translate('Size:')}</label><input aria-label={translate('Size:')} type="text" value={fontSizeInput} onChange={(event) => setFontSizeInput(event.currentTarget.value)} onBlur={() => updateFont({ sz: fontSizeInput })} onKeyDown={(event) => { if (event.key === 'Enter') updateFont({ sz: fontSizeInput }); }} /><div className="lbox short" role="listbox" aria-label={translate('Size:')}>{catalog.fontSizes.map((size) => <div className={`lbox-i${Number(font.sz || 10) === size ? ' on' : ''}`} role="option" aria-selected={Number(font.sz || 10) === size} key={size} onClick={() => updateFont({ sz: size })}>{size}</div>)}</div></div>
      </div>
      <div className="excelx-format-cells-dialog__font-row"><label>{translate('&Underline:').replace('&', '')}<select aria-label={translate('&Underline:').replace('&', '')} value={String(font.u || 'none')} onChange={(event) => updateFont({ u: event.currentTarget.value })}>{([['none','None'],['single','Single'],['double','Double'],['singleAccounting','Single Accounting'],['doubleAccounting','Double Accounting']] as const).map(([value,label]) => <option value={value} key={value}>{translate(label)}</option>)}</select></label>
        <label>{translate('&Color:').replace('&', '')}<button ref={colorButton} type="button" className="swatch-btn wide" style={{ background: fontColors.hex }} onClick={() => { if (colorButton.current) runtime()?.dlg?.reactFormatCellsChooseFontColor?.(colorButton.current); }}><span>{fontColors.automatic ? translate('Automatic') : ''}</span><span className="dd-arrow" /></button></label>
        <label><input type="checkbox" aria-label={translate('&Normal font').replace('&', '')} onChange={(event) => { if (event.currentTarget.checked) updateFont({ normal: true }); }} />{translate('&Normal font').replace('&', '')}</label></div>
      <div className="excelx-format-cells-dialog__font-row"><fieldset><legend>{translate('Effects')}</legend>
        <label><input type="checkbox" aria-label={translate('Stri&kethrough').replace('&', '')} checked={!!font.strike} onChange={(event) => updateFont({ strike: event.currentTarget.checked })} />{translate('Stri&kethrough').replace('&', '')}</label>
        <label><input type="checkbox" aria-label={translate('Sup&erscript').replace('&', '')} checked={font.vert === 'superscript'} onChange={(event) => updateFont({ vert: event.currentTarget.checked ? 'superscript' : '' })} />{translate('Sup&erscript').replace('&', '')}</label>
        <label><input type="checkbox" aria-label={translate('Su&bscript').replace('&', '')} checked={font.vert === 'subscript'} onChange={(event) => updateFont({ vert: event.currentTarget.checked ? 'subscript' : '' })} />{translate('Su&bscript').replace('&', '')}</label>
      </fieldset><fieldset><legend>{translate('Preview')}</legend><div className="excelx-format-cells-dialog__font-preview" style={{ fontFamily: String(fontName), fontSize: `${Number(font.sz || 10) * 1.3}px`, fontWeight: font.b ? 'bold' : undefined, fontStyle: font.i ? 'italic' : undefined, color: fontColors.hex, textDecoration: [font.u && font.u !== 'none' ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ') || 'none' }}>{translate('AaBbCcYyZz')}</div></fieldset></div>
    </div>, fontHost)}
    {fillHost && createPortal(<div className="excelx-format-cells-dialog__patterns">
      <fieldset><legend>{translate('Cell shading')}</legend>
        <label>{translate('Color:')}</label>
        <button type="button" className="cp-wide" onClick={() => updateFill({ fillColor: null })}>{translate('No Color')}</button>
        <div className="cp-grid">{catalog.fillColors.map((color) => <button key={color.value} type="button" className="cp-sw" data-fill-color={color.value} style={{ backgroundColor: color.value }} aria-label={translate(color.label)} title={translate(color.label)} onClick={() => updateFill({ fillColor: color.value })} />)}</div>
        <div className="excelx-format-cells-dialog__pattern-picker">
          <span>{translate('Pattern:')}</span>
          <button type="button" className="swatch-btn wide" aria-expanded={patternMenuOpen} onClick={() => setPatternMenuOpen((open) => !open)}>{translate(fill.pattern || 'Pattern')}<span className="dd-arrow" /></button>
          {patternMenuOpen && <div className="excelx-format-cells-dialog__pattern-menu">
            <div className="pat-grid">{catalog.patterns.map((pattern) => <button key={pattern.value} type="button" data-pattern={pattern.value} aria-label={translate(pattern.label)} title={translate(pattern.label)} aria-pressed={(fill.pattern || 'solid') === pattern.value} style={{ backgroundColor: '#fff', backgroundImage: `url("${pattern.preview}")` }} onClick={() => { updateFill({ pattern: pattern.value === 'solid' ? null : pattern.value }); setPatternMenuOpen(false); }} />)}</div>
            <div className="cp-sep" /><label>{translate('Pattern color:')}</label>
            <div className="cp-grid">{catalog.fillColors.map((color) => <button key={color.value} type="button" className="cp-sw" data-pattern-color={color.value} style={{ backgroundColor: color.value }} aria-label={translate(color.label)} title={translate(color.label)} onClick={() => updateFill({ patColor: color.value })} />)}</div>
          </div>}
        </div>
      </fieldset>
      <fieldset><legend>{translate('Sample')}</legend><div className="fc-psample" aria-label={translate('Fill preview')} style={{ backgroundColor: fill.fillColor || '#fff', backgroundImage: fill.pattern ? `url("${patternPreview?.(fill.pattern, fill.patColor, fill.fillColor || '#FFFFFF')}")` : 'none' }} /></fieldset>
    </div>, fillHost)}
    {borderHost && createPortal(<div className="excelx-format-cells-dialog__border">
      <div className="excelx-format-cells-dialog__border-main">
        <fieldset><legend>{translate('Presets')}</legend><div className="excelx-format-cells-dialog__border-presets">
          <button type="button" onClick={() => updateBorder({ type: 'preset', value: 'none' })}>{translate('None')}</button>
          <button type="button" onClick={() => updateBorder({ type: 'preset', value: 'outline' })}>{translate('Outline|border')}</button>
          <button type="button" disabled={!border.multiR && !border.multiC} onClick={() => updateBorder({ type: 'preset', value: 'inside' })}>{translate('Inside')}</button>
        </div></fieldset>
        <fieldset><legend>{translate('Border')}</legend><div className="excelx-format-cells-dialog__border-tools">
          <div className="excelx-format-cells-dialog__border-buttons">{([['top','Top'],['insideH','Inside horizontal'],['bottom','Bottom'],['diagUp','Diagonal up'],['left','Left'],['insideV','Inside vertical'],['right','Right'],['diagDown','Diagonal down']] as const).map(([key,label]) => <button key={key} type="button" aria-label={translate(label)} aria-pressed={!!border.edges[key]} disabled={(key==='insideH'&&!border.multiR)||(key==='insideV'&&!border.multiC)} onClick={() => updateBorder({ type: 'toggle', edge: key })}>{translate(label)}</button>)}</div>
          <svg className="excelx-format-cells-dialog__border-preview" viewBox="0 0 150 110" role="img" aria-label={translate('Border preview')} onClick={toggleBorderAt}>
            <rect x="15" y="10" width="120" height="90" fill="#fff" />
            <text x={border.multiC?45:75} y={border.multiR?33:58} textAnchor="middle">{translate('Text')}</text>{border.multiC&&<text x="105" y={border.multiR?33:58} textAnchor="middle">{translate('Text')}</text>}{border.multiR&&<text x={border.multiC?45:75} y="80" textAnchor="middle">{translate('Text')}</text>}
            {borderStroke(border.edges.top,15,10,135,10)}{borderStroke(border.edges.bottom,15,100,135,100)}{borderStroke(border.edges.left,15,10,15,100)}{borderStroke(border.edges.right,135,10,135,100)}{border.multiR&&borderStroke(border.edges.insideH,15,55,135,55)}{border.multiC&&borderStroke(border.edges.insideV,75,10,75,100)}{borderStroke(border.edges.diagUp,15,100,135,10)}{borderStroke(border.edges.diagDown,15,10,135,100)}
          </svg>
        </div></fieldset>
      </div>
      <fieldset><legend>{translate('Line')}</legend><div className="excelx-format-cells-dialog__border-line"><label>{translate('Style:')}</label><div className="fc-lines" role="listbox" aria-label={translate('Style:')}>{catalog.borderLines.map((item) => <button key={item.style} type="button" className={`fc-line${border.style===item.style?' on':''}`} role="option" aria-selected={border.style===item.style} aria-label={translate(item.label)} onClick={() => updateBorder({ type: 'style', style: item.style })}>{item.style==='none'?<span>{translate('None')}</span>:<svg width="52" height="14" viewBox="0 0 52 14">{borderStroke({style:item.style,color:'#000'},2,7,50,7)}</svg>}</button>)}</div>
        <label>{translate('&Color:').replace('&','')}<button ref={borderColorButton} type="button" className="swatch-btn wide" style={{background:border.color||'#000'}} onClick={() => { if(borderColorButton.current) runtime()?.dlg?.reactFormatCellsChooseBorderColor?.(borderColorButton.current); }}><span>{border.color?translate(''):translate('Automatic')}</span><span className="dd-arrow" /></button></label>
      </div></fieldset>
    </div>, borderHost)}
    {alignmentHost && createPortal(<div className="excelx-format-cells-dialog__alignment">
      <fieldset><legend>{translate('Text alignment')}</legend>
        <label>{translate('&Horizontal:').replace('&', '')}<select aria-label="Horizontal alignment" value={String(alignment.h || 'general')} onChange={(event) => setAlignmentValue('h', event.currentTarget.value)}>{([['general','General'],['left','Left (Indent)'],['center','Center'],['right','Right (Indent)'],['fill','Fill'],['justify','Justify'],['centerContinuous','Center Across Selection'],['distributed','Distributed (Indent)']] as const).map(([value,label]) => <option key={value} value={value}>{translate(label)}</option>)}</select></label>
        <label>{translate('&Indent:').replace('&', '')}<input aria-label="Indent" type="number" min="0" max="15" step="1" value={Number(alignment.indent || 0)} disabled={!/left|right|distributed/.test(String(alignment.h || ''))} onChange={(event) => setAlignmentValue('indent', event.currentTarget.value)} /></label>
        <label>{translate('&Vertical:').replace('&', '')}<select aria-label="Vertical alignment" value={String(alignment.v || 'bottom')} onChange={(event) => setAlignmentValue('v', event.currentTarget.value)}>{([['top','Top'],['center','Center'],['bottom','Bottom'],['justify','Justify'],['distributed','Distributed']] as const).map(([value,label]) => <option key={value} value={value}>{translate(label)}</option>)}</select></label>
        <label><input aria-label="Justify distributed" type="checkbox" checked={!!alignment.justLast} onChange={(event) => setAlignmentValue('justLast', event.currentTarget.checked)} />{translate('&Justify distributed').replace('&', '')}</label>
      </fieldset>
      <fieldset><legend>{translate('Text control')}</legend>
        {([['wrap','&Wrap text'],['shrink','Shrin&k to fit']] as const).map(([key,label]) => <label key={key}><input type="checkbox" aria-label={translate(label).replace('&', '')} checked={!!alignment[key]} onChange={(event) => setAlignmentValue(key, event.currentTarget.checked)} />{translate(label).replace('&', '')}</label>)}
        <label><input type="checkbox" aria-label={translate('&Merge cells').replace('&', '')} checked={merged} onChange={(event) => { setMerged(event.currentTarget.checked); setAlignmentValue('merge', event.currentTarget.checked); }} />{translate('&Merge cells').replace('&', '')}</label>
      </fieldset>
      <fieldset><legend>{translate('Right-to-left')}</legend><label>{translate('&Text direction:').replace('&', '')}<select aria-label="Text direction" value={alignment.rtl ? 'rtl' : alignment.ltr ? 'ltr' : 'ctx'} onChange={(event) => { const value = event.currentTarget.value; setAlignment((current) => ({ ...current, rtl: value === 'rtl', ltr: value === 'ltr' })); runtime()?.dlg?.reactFormatCellsSetAlignment?.({ rtl: value === 'rtl', ltr: value === 'ltr' }); }}><option value="ctx">{translate('Context')}</option><option value="ltr">{translate('Left-to-Right')}</option><option value="rtl">{translate('Right-to-Left')}</option></select></label></fieldset>
      <fieldset><legend>{translate('Orientation')}</legend><label>{translate('Rotation')} <input aria-label="Text rotation" type="range" min="-90" max="90" step="1" value={rotation} onChange={(event) => setAlignmentValue('rot', Number(event.currentTarget.value))} /></label><output>{rotation}°</output><label><input type="checkbox" aria-label={translate('Vertical text')} checked={alignment.rot === 255} onChange={(event) => setAlignmentValue('rot', event.currentTarget.checked ? 255 : 0)} />{translate('Vertical text')}</label></fieldset>
    </div>, alignmentHost)}
    {protectionHost && createPortal(<div className="excelx-format-cells-dialog__protection">
      <label><input type="checkbox" aria-label={translate('&Locked').replace('&', '')} checked={protection.locked} onChange={(event) => { const next = { ...protection, locked: event.currentTarget.checked }; setProtection(next); runtime()?.dlg?.reactFormatCellsSetProtection?.(next.locked, next.hidden); }} />{translate('&Locked').replace('&', '')}</label>
      <label><input type="checkbox" aria-label={translate('H&idden').replace('&', '')} checked={protection.hidden} onChange={(event) => { const next = { ...protection, hidden: event.currentTarget.checked }; setProtection(next); runtime()?.dlg?.reactFormatCellsSetProtection?.(next.locked, next.hidden); }} />{translate('H&idden').replace('&', '')}</label>
      <p className="hint">{translate('Locking cells or hiding formulas has no effect unless the worksheet is protected. To protect the worksheet, choose Protection from the Tools menu, and then choose Protect Sheet. A password is optional.')}</p>
    </div>, protectionHost)}
    <div className="excelx-format-cells-dialog__actions">
      <span />
      <button type="button" className="primary" onClick={() => runtime()?.dlg?.reactFormatCellsApply?.()}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
