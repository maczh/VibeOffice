import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useMemo, useState } from 'react';

interface SymbolRange { index: number; label: string; start: number; end: number }
interface SymbolOption { value: string; label: string }
interface SymbolCatalog { ranges: SymbolRange[]; special: SymbolOption[] }
interface SymbolRuntime {
  dlg?: { reactSymbolCatalog?(): SymbolCatalog | null; reactSymbolInsert?(character: string): boolean; reactSymbolClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
  layout?: { fontStack?(font: string): string };
}

function runtime(): SymbolRuntime | undefined {
  return (window as Window & { L?: SymbolRuntime }).L;
}

export function ExcelxSymbolDialog() {
  const [catalog, setCatalog] = useState<SymbolCatalog | null>(null);
  const [tab, setTab] = useState<'symbols' | 'special'>('symbols');
  const [subset, setSubset] = useState(1);
  const [font, setFont] = useState('Arial');
  const [selected, setSelected] = useState<string | null>(null);
  const [codePoint, setCodePoint] = useState<number | null>(null);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactSymbolCatalog?.() || null;
      setCatalog(next);
      if (next) { setTab('symbols'); setSubset(1); setFont('Arial'); setSelected(null); setCodePoint(null); }
    };
    const off = runtime()?.bus?.on('excelx-symbol-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactSymbolCatalog?.()) runtime()?.dlg?.reactSymbolClose?.();
    };
  }, []);
  const range = catalog?.ranges.find((item) => item.index === subset);
  const characters = useMemo(() => range ? Array.from({ length: range.end - range.start + 1 }, (_, index) => range.start + index) : [], [range]);
  if (!catalog) return null;
  const translate = (message: string, ...args: unknown[]) => runtime()?.t?.(message, ...args) || message;
  const close = () => runtime()?.dlg?.reactSymbolClose?.();
  const insert = (character: string | null = selected) => { if (character) runtime()?.dlg?.reactSymbolInsert?.(character); };
  const fontStack = runtime()?.layout?.fontStack?.(font) || `${font}, sans-serif`;
  return <OfficeDialog title={translate('Symbol')} className="excelx-symbol-dialog" style={{ width: 'min(94vw, 560px)' }} onRequestClose={close}>
    <div className="excelx-symbol-dialog__tabs" role="tablist" aria-label={translate('Symbol')}>
      <button type="button" role="tab" aria-selected={tab === 'symbols'} onClick={() => setTab('symbols')}>{translate('&Symbols').replace('&', '')}</button>
      <button type="button" role="tab" aria-selected={tab === 'special'} onClick={() => setTab('special')}>{translate('S&pecial Characters').replace('&', '')}</button>
    </div>
    {tab === 'symbols' ? <div className="excelx-symbol-dialog__body">
      <div className="excelx-symbol-dialog__selectors">
        <label>{translate('&Font:').replace('&', '')}<input value={font} onChange={(event) => setFont(event.currentTarget.value)} /></label>
        <label>{translate('Su&bset:').replace('&', '')}<select value={subset} onChange={(event) => { setSubset(Number(event.currentTarget.value)); setSelected(null); setCodePoint(null); }}>{catalog.ranges.map(({ index, label }) => <option key={index} value={index}>{translate(label)}</option>)}</select></label>
      </div>
      <div className="sym-grid" translate="no" role="grid" aria-label={range ? translate(range.label) : ''} style={{ display: 'grid', gridTemplateColumns: 'repeat(16, minmax(24px, 1fr))', gap: 2, maxHeight: 260, overflow: 'auto' }}>
        {characters.map((code) => { const character = String.fromCodePoint(code); return <button key={code} type="button" aria-label={`U+${code.toString(16).toUpperCase()}`} aria-pressed={selected === character} translate="no" style={{ fontFamily: fontStack }} onClick={() => { setSelected(character); setCodePoint(code); }} onDoubleClick={() => { setSelected(character); insert(character); }}>{character}</button>; })}
      </div>
      <div className="hint" aria-live="polite">{codePoint == null ? '' : translate('Character code: {0} (Unicode hex)', codePoint.toString(16).toUpperCase().padStart(4, '0'))}</div>
    </div> : <div className="excelx-symbol-dialog__special" role="listbox" aria-label={translate('S&pecial Characters').replace('&', '')}>
      {catalog.special.map(({ value, label }) => <button key={label} type="button" role="option" aria-selected={selected === value} onClick={() => setSelected(value)} onDoubleClick={() => { setSelected(value); insert(value); }}><span translate="no">{value}</span>{translate(label)}</button>)}
    </div>}
    <div className="excelx-symbol-dialog__actions">
      <button type="button" className="primary" disabled={!selected} onClick={() => insert()}>{translate('&Insert').replace('&', '')}</button>
      <button type="button" onClick={close}>{translate('Close')}</button>
    </div>
  </OfficeDialog>;
}
