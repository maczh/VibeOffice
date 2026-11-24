import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useMemo, useState, type KeyboardEvent } from 'react';
import { useExcelxRuntime } from './editor-host';

interface SymbolRange { index: number; label: string; start: number; end: number }
interface SymbolOption { value: string; label: string }
interface SymbolCatalog { ranges: SymbolRange[]; special: SymbolOption[] }
interface SymbolRuntime {
  dlg?: { reactSymbolCatalog?(): SymbolCatalog | null; reactSymbolInsert?(character: string): boolean; reactSymbolClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
  layout?: { fontStack?(font: string): string };
}

export function ExcelxSymbolDialog() {
  const runtime = useExcelxRuntime<SymbolRuntime>();
  const [catalog, setCatalog] = useState<SymbolCatalog | null>(null);
  const [tab, setTab] = useState<'symbols' | 'special'>('symbols');
  const id = useId().replaceAll(':', '');
  const [subset, setSubset] = useState(1);
  const [font, setFont] = useState('Arial');
  const [selected, setSelected] = useState<string | null>(null);
  const [activeSpecial, setActiveSpecial] = useState(0);
  const [codePoint, setCodePoint] = useState<number | null>(null);
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.dlg?.reactSymbolCatalog?.() || null;
      setCatalog(next);
      if (next) { setTab('symbols'); setSubset(1); setFont('Arial'); setSelected(null); setCodePoint(null); setActiveSpecial(0); }
    };
    const off = runtime?.bus?.on('excelx-symbol-dialog', refresh);
    return () => {
      off?.();
      if (runtime?.dlg?.reactSymbolCatalog?.()) runtime?.dlg?.reactSymbolClose?.();
    };
  }, []);
  const range = catalog?.ranges.find((item) => item.index === subset);
  const characters = useMemo(() => range ? Array.from({ length: range.end - range.start + 1 }, (_, index) => range.start + index) : [], [range]);
  if (!catalog) return null;
  const translate = (message: string, ...args: unknown[]) => runtime?.t?.(message, ...args) || message;
  const close = () => runtime?.dlg?.reactSymbolClose?.();
  const insert = (character: string | null = selected) => { if (character) runtime?.dlg?.reactSymbolInsert?.(character); };
  const fontStack = runtime?.layout?.fontStack?.(font) || `${font}, sans-serif`;
  const tabId = (key: 'symbols' | 'special') => `${id}-symbol-tab-${key}`;
  const panelId = (key: 'symbols' | 'special') => `${id}-symbol-panel-${key}`;
  const selectSpecial = (index: number) => {
    const character = catalog.special[index]?.value;
    if (character === undefined) return;
    setActiveSpecial(index);
    setSelected(character);
  };
  const onSpecialKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (!catalog.special.length) return;
    let next = activeSpecial;
    if (event.key === 'ArrowDown' || event.key === 'ArrowRight') next = (activeSpecial + 1) % catalog.special.length;
    else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') next = (activeSpecial - 1 + catalog.special.length) % catalog.special.length;
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = catalog.special.length - 1;
    else if (event.key === 'Enter' || event.key === ' ') next = activeSpecial;
    else return;
    event.preventDefault();
    event.stopPropagation();
    selectSpecial(next);
  };
  return <OfficeDialog title={translate('Symbol')} className="excelx-symbol-dialog" style={{ width: 'min(94vw, 560px)' }} onRequestClose={close}>
    <div className="excelx-symbol-dialog__tabs" role="tablist" aria-label={translate('Symbol')}>
      {(['symbols', 'special'] as const).map((key, index) => <button key={key} type="button" id={tabId(key)} role="tab" aria-controls={panelId(key)} aria-selected={tab === key} tabIndex={tab === key ? 0 : -1} onClick={() => setTab(key)} onKeyDown={(event) => {
        const next = event.key === 'ArrowRight' ? (index + 1) % 2
          : event.key === 'ArrowLeft' ? (index - 1 + 2) % 2
            : event.key === 'Home' ? 0
              : event.key === 'End' ? 1 : -1;
        if (next < 0) return;
        event.preventDefault();
        const nextTab = next === 0 ? 'symbols' : 'special';
        setTab(nextTab);
        event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
      }}>{translate(key === 'symbols' ? '&Symbols' : 'S&pecial Characters').replace('&', '')}</button>)}
    </div>
    <div id={panelId('symbols')} aria-labelledby={tabId('symbols')} role="tabpanel" className="excelx-symbol-dialog__body" hidden={tab !== 'symbols'}>
      <div className="excelx-symbol-dialog__selectors">
        <label>{translate('&Font:').replace('&', '')}<input value={font} onChange={(event) => setFont(event.currentTarget.value)} /></label>
        <label>{translate('Su&bset:').replace('&', '')}<select value={subset} onChange={(event) => { setSubset(Number(event.currentTarget.value)); setSelected(null); setCodePoint(null); }}>{catalog.ranges.map(({ index, label }) => <option key={index} value={index}>{translate(label)}</option>)}</select></label>
      </div>
      <div className="sym-grid" translate="no" role="grid" aria-label={range ? translate(range.label) : ''} style={{ display: 'grid', gridTemplateColumns: 'repeat(16, minmax(24px, 1fr))', gap: 2, maxHeight: 260, overflow: 'auto' }}>
        {characters.map((code) => { const character = String.fromCodePoint(code); return <button key={code} type="button" aria-label={`U+${code.toString(16).toUpperCase()}`} aria-pressed={selected === character} translate="no" style={{ fontFamily: fontStack }} onClick={() => { setSelected(character); setCodePoint(code); }} onDoubleClick={() => { setSelected(character); insert(character); }}>{character}</button>; })}
      </div>
      <div className="hint" aria-live="polite">{codePoint == null ? '' : translate('Character code: {0} (Unicode hex)', codePoint.toString(16).toUpperCase().padStart(4, '0'))}</div>
    </div>
    <div id={panelId('special')} aria-labelledby={tabId('special')} role="tabpanel" className="excelx-symbol-dialog__special-panel" hidden={tab !== 'special'}>
      <div className="excelx-symbol-dialog__special" id={`${id}-special-list`} role="listbox" aria-label={translate('S&pecial Characters').replace('&', '')} aria-activedescendant={catalog.special[activeSpecial] ? `${id}-special-option-${activeSpecial}` : undefined} tabIndex={0} onKeyDown={onSpecialKeyDown}>
      {catalog.special.map(({ value, label }, index) => <button id={`${id}-special-option-${index}`} key={label} type="button" role="option" aria-selected={selected === value} tabIndex={-1} onPointerEnter={() => setActiveSpecial(index)} onClick={() => selectSpecial(index)} onDoubleClick={() => { selectSpecial(index); insert(value); }}><span translate="no">{value}</span>{translate(label)}</button>)}
      </div>
    </div>
    <div className="excelx-symbol-dialog__actions">
      <button type="button" className="primary" disabled={!selected} onClick={() => insert()}>{translate('&Insert').replace('&', '')}</button>
      <button type="button" onClick={close}>{translate('Close')}</button>
    </div>
  </OfficeDialog>;
}
