import { createPortal } from 'react-dom';
import { useEffect, useState } from 'react';

interface WordxMenuRuntime {
  app?: {
    reactTableGridAction?(rows: number, columns: number): void;
    reactColumnsPickerAction?(columns: number): void;
    reactHighlightMenuCatalog?(): Array<{ kind: string; name: string; color: string }>;
    reactHighlightMenuAction?(kind: string): void;
    reactHighlightNone?(): void;
    reactBorderMenuCatalog?(): Array<{ kind: string; icon: string; label: string }>;
    reactBorderMenuAction?(kind: string): void;
    reactBorderStyleCatalog?(): Array<{ kind: string; label: string; svg: string }>;
    reactBorderStyleAction?(kind: string): void;
    reactBorderWeightCatalog?(): Array<{ weight: number; label: string; svg: string }>;
    reactBorderWeightAction?(weight: number): void;
    reactObjectWeightCatalog?(): Array<{ weight: number; label: string; svg: string }>;
    reactObjectWeightAction?(weight: number | 'more'): void;
    reactObjectDashCatalog?(): Array<{ kind: string; label: string; svg: string }>;
    reactObjectDashAction?(kind: string): void;
    reactObjectArrowCatalog?(): Array<{ index: number; label: string; svg: string }>;
    reactObjectArrowAction?(index: number): void;
    reactBackgroundCatalog?(): Array<{ color: string; name: string }>;
    reactBackgroundAction?(color: string): void;
  };
  find?: {
    reactFindFormatMenuCatalog?(): Array<{ id: string; label: string }>;
    reactFindFormatMenuAction?(id: string): void;
    reactFindSpecialMenuCatalog?(): Array<{ code: string; label: string }>;
    reactFindSpecialMenuAction?(code: string): void;
  };
  panes?: { preview?: { reactMultiPageAction?(rows: number, columns: number): void } };
  t?(text: string, ...args: unknown[]): string;
}

function runtime(): WordxMenuRuntime | undefined {
  return (window as Window & { L?: WordxMenuRuntime }).L;
}

function TableGrid({ L }: { L: WordxMenuRuntime }) {
  const [size, setSize] = useState<[number, number]>([0, 0]);
  const [bounds, setBounds] = useState<[number, number]>([4, 5]);
  const [rows, columns] = size;
  const [maxRows, maxColumns] = bounds;
  const label = rows ? L.t?.('{0} x {1} Table', rows, columns) || `${rows} x ${columns} Table` : L.t?.('Cancel') || 'Cancel';
  return <div className="tbl-grid-pick" id="react-wordx-table-grid">
    <div className="tgp-cells" style={{ gridTemplateColumns: `repeat(${maxColumns}, 18px)` }} onPointerLeave={() => setSize([0, 0])}>
      {Array.from({ length: maxRows }, (_, rowIndex) => Array.from({ length: maxColumns }, (_, columnIndex) => {
        const row = rowIndex + 1;
        const column = columnIndex + 1;
        return <span key={`${row}:${column}`} role="button" tabIndex={0} className={row <= rows && column <= columns ? 'on' : ''}
          aria-label={L.t?.('{0} x {1} Table', row, column) || `${row} x ${column} Table`}
          onPointerEnter={() => {
            setSize([row, column]);
            setBounds(([r, c]) => [row >= r - 1 && r < 12 ? r + 1 : r, column >= c - 1 && c < 10 ? c + 1 : c]);
          }}
          onClick={() => { if (rows && columns) L.app?.reactTableGridAction?.(rows, columns); }}
          onKeyDown={(event) => { if ((event.key === 'Enter' || event.key === ' ') && rows && columns) { event.preventDefault(); L.app?.reactTableGridAction?.(rows, columns); } }} />;
      }))}
    </div>
    <div className="tgp-label" aria-live="polite">{label}</div>
  </div>;
}

function ColumnsPicker({ L }: { L: WordxMenuRuntime }) {
  const [columns, setColumns] = useState(0);
  return <div className="tbl-grid-pick" id="react-wordx-columns-picker">
    <div className="tgp-cells" style={{ gridTemplateColumns: 'repeat(4,22px)' }}>
      {Array.from({ length: 4 }, (_, index) => {
        const value = index + 1;
        const label = L.t?.(value === 1 ? '{0} Column' : '{0} Columns', value) || `${value} ${value === 1 ? 'Column' : 'Columns'}`;
        return <span key={value} role="button" tabIndex={0} style={{ height: 30 }} className={value <= columns ? 'on' : ''}
          aria-label={label} onPointerEnter={() => setColumns(value)} onFocus={() => setColumns(value)}
          onClick={() => L.app?.reactColumnsPickerAction?.(value)}
          onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); L.app?.reactColumnsPickerAction?.(value); } }} />;
      })}
    </div>
    <div className="tgp-label" aria-live="polite">{columns ? L.t?.(columns === 1 ? '{0} Column' : '{0} Columns', columns) : L.t?.('Cancel') || 'Cancel'}</div>
  </div>;
}

function HighlightPalette({ L }: { L: WordxMenuRuntime }) {
  const colors = L.app?.reactHighlightMenuCatalog?.() || [];
  return <div className="color-panel" id="react-wordx-highlight-menu">
    <div className="cp-grid hl">{colors.map(({ kind, name, color }) => <button key={kind} type="button" className="cp-sw"
      style={{ background: color }} aria-label={L.t?.(name) || name} title={L.t?.(name) || name}
      onClick={() => L.app?.reactHighlightMenuAction?.(kind)} />)}</div>
    <button type="button" className="cp-wide" onClick={() => L.app?.reactHighlightNone?.()}>{L.t?.('None') || 'None'}</button>
  </div>;
}

function BorderMenu({ L }: { L: WordxMenuRuntime }) {
  const items = L.app?.reactBorderMenuCatalog?.() || [];
  return <div className="border-grid" id="react-wordx-border-menu">{items.map(({ kind, icon, label }) => <button key={kind} type="button"
    data-tip={label} aria-label={label} dangerouslySetInnerHTML={{ __html: icon }}
    onClick={() => L.app?.reactBorderMenuAction?.(kind)} />)}</div>;
}

function BorderStyleMenu({ L }: { L: WordxMenuRuntime }) {
  const items = L.app?.reactBorderStyleCatalog?.() || [];
  return <div className="style-list" id="react-wordx-border-style-menu">{items.map(({ kind, label, svg }) => <button key={kind} type="button" aria-label={label} title={label}
    dangerouslySetInnerHTML={{ __html: svg }} onClick={() => L.app?.reactBorderStyleAction?.(kind)} />)}
    <button type="button" onClick={() => L.app?.reactBorderStyleAction?.('nil')}>{L.t?.('No Border') || 'No Border'}</button>
  </div>;
}

function BorderWeightMenu({ L }: { L: WordxMenuRuntime }) {
  const items = L.app?.reactBorderWeightCatalog?.() || [];
  return <div className="style-list" id="react-wordx-border-weight-menu">{items.map(({ weight, label, svg }) => <button key={weight} type="button" aria-label={label} title={label} onClick={() => L.app?.reactBorderWeightAction?.(weight)}>
    <span dangerouslySetInnerHTML={{ __html: svg }} /><span>{label}</span></button>)}</div>;
}

function ObjectWeightMenu({ L }: { L: WordxMenuRuntime }) {
  const items = L.app?.reactObjectWeightCatalog?.() || [];
  return <div className="style-list" id="react-wordx-object-weight-menu">{items.map(({ weight, label, svg }) => <button key={weight} type="button" aria-label={label} title={label}
    onClick={() => L.app?.reactObjectWeightAction?.(weight)}><span dangerouslySetInnerHTML={{ __html: svg }} /><span>{label}</span></button>)}
    <button type="button" onClick={() => L.app?.reactObjectWeightAction?.('more')}>{L.t?.('More Lines...') || 'More Lines...'}</button>
  </div>;
}

function ObjectDashMenu({ L }: { L: WordxMenuRuntime }) {
  const items = L.app?.reactObjectDashCatalog?.() || [];
  return <div className="style-list" id="react-wordx-object-dash-menu">{items.map(({ kind, label, svg }) => <button key={kind} type="button" aria-label={label} title={label}
    onClick={() => L.app?.reactObjectDashAction?.(kind)}><span dangerouslySetInnerHTML={{ __html: svg }} /></button>)}</div>;
}

function ObjectArrowMenu({ L }: { L: WordxMenuRuntime }) {
  const items = L.app?.reactObjectArrowCatalog?.() || [];
  return <div className="style-list" id="react-wordx-object-arrow-menu">{items.map(({ index, label, svg }) => <button key={index} type="button" aria-label={label} title={label}
    onClick={() => L.app?.reactObjectArrowAction?.(index)}><span dangerouslySetInnerHTML={{ __html: svg }} /></button>)}</div>;
}

function BackgroundPalette({ L }: { L: WordxMenuRuntime }) {
  const colors = L.app?.reactBackgroundCatalog?.() || [];
  return <div className="color-panel" id="react-wordx-background-menu">
    <div className="cp-grid">{colors.map(({ color, name }) => <button key={color} type="button" className="cp-sw" style={{ background: color }}
      aria-label={name} title={name} onClick={() => L.app?.reactBackgroundAction?.(color)} />)}</div>
  </div>;
}

function MultiPagePicker({ L }: { L: WordxMenuRuntime }) {
  const [selection, setSelection] = useState<[number, number]>([0, 0]);
  const [rows, columns] = selection;
  const label = rows ? L.t?.('{0} x {1} Pages', rows, columns) || `${rows} x ${columns} Pages` : L.t?.('Cancel') || 'Cancel';
  return <div className="tbl-grid-pick" id="react-wordx-multipage-menu">
    <div className="tgp-cells" style={{ gridTemplateColumns: 'repeat(3,24px)' }} onPointerLeave={() => setSelection([0, 0])}>
      {Array.from({ length: 6 }, (_, index) => {
        const row = Math.floor(index / 3) + 1;
        const column = index % 3 + 1;
        return <span key={`${row}:${column}`} role="button" tabIndex={0} style={{ height: 30, width: 24 }}
          className={row <= rows && column <= columns ? 'on' : ''} aria-label={L.t?.('{0} x {1} Pages', row, column) || `${row} x ${column} Pages`}
          onPointerEnter={() => setSelection([row, column])} onFocus={() => setSelection([row, column])}
          onClick={() => { if (rows && columns) L.panes?.preview?.reactMultiPageAction?.(rows, columns); }}
          onKeyDown={(event) => { if ((event.key === 'Enter' || event.key === ' ') && rows && columns) { event.preventDefault(); L.panes?.preview?.reactMultiPageAction?.(rows, columns); } }} />;
      })}
    </div>
    <div className="tgp-label" aria-live="polite">{label}</div>
  </div>;
}

function FindFormatMenu({ L }: { L: WordxMenuRuntime }) {
  const items = L.find?.reactFindFormatMenuCatalog?.() || [];
  useEffect(() => () => {
    if (L.find) { delete L.find.reactFindFormatMenuCatalog; delete L.find.reactFindFormatMenuAction; }
  }, [L]);
  return <div id="react-wordx-find-format-menu" role="group">{items.map(({ id, label }) => <button key={id} type="button" role="menuitem" className="m-item" onClick={() => L.find?.reactFindFormatMenuAction?.(id)}>
    <span className="m-ico" /><span className="m-lbl">{label}</span><span /><span />
  </button>)}</div>;
}

function FindSpecialMenu({ L }: { L: WordxMenuRuntime }) {
  const items = L.find?.reactFindSpecialMenuCatalog?.() || [];
  useEffect(() => () => {
    if (L.find) { delete L.find.reactFindSpecialMenuCatalog; delete L.find.reactFindSpecialMenuAction; }
  }, [L]);
  return <div id="react-wordx-find-special-menu" role="group">{items.map(({ code, label }) => <button key={code} type="button" role="menuitem" className="m-item" onClick={() => L.find?.reactFindSpecialMenuAction?.(code)}>
    <span className="m-ico" /><span className="m-lbl">{label}</span><span /><span />
  </button>)}</div>;
}

/** React owns the contents of Wordx's legacy menu slots; the editor retains only commit operations. */
export function WordxMenuContentLayer({ host }: { host: HTMLElement }) {
  const [targets, setTargets] = useState<{ table: HTMLElement | null; columns: HTMLElement | null; highlight: HTMLElement | null; borders: HTMLElement | null; borderStyle: HTMLElement | null; borderWeight: HTMLElement | null; objectWeight: HTMLElement | null; objectDash: HTMLElement | null; objectArrow: HTMLElement | null; background: HTMLElement | null; multipage: HTMLElement | null; findFormat: HTMLElement | null; findSpecial: HTMLElement | null }>({ table: null, columns: null, highlight: null, borders: null, borderStyle: null, borderWeight: null, objectWeight: null, objectDash: null, objectArrow: null, background: null, multipage: null, findFormat: null, findSpecial: null });
  useEffect(() => {
    const scan = () => setTargets({
      table: host.querySelector<HTMLElement>('#react-wordx-table-grid-menu-root'),
      columns: host.querySelector<HTMLElement>('#react-wordx-columns-picker-menu-root'),
      highlight: host.querySelector<HTMLElement>('#react-wordx-highlight-menu-root'),
      borders: host.querySelector<HTMLElement>('#react-wordx-border-menu-root'),
      borderStyle: host.querySelector<HTMLElement>('#react-wordx-border-style-menu-root'),
      borderWeight: host.querySelector<HTMLElement>('#react-wordx-border-weight-menu-root'),
      objectWeight: host.querySelector<HTMLElement>('#react-wordx-object-weight-menu-root'),
      objectDash: host.querySelector<HTMLElement>('#react-wordx-object-dash-menu-root'),
      objectArrow: host.querySelector<HTMLElement>('#react-wordx-object-arrow-menu-root'),
      background: host.querySelector<HTMLElement>('#react-wordx-background-menu-root'),
      multipage: host.querySelector<HTMLElement>('#react-wordx-multipage-menu-root'),
      findFormat: host.querySelector<HTMLElement>('#react-wordx-find-format-menu-root'),
      findSpecial: host.querySelector<HTMLElement>('#react-wordx-find-special-menu-root'),
    });
    const observer = new MutationObserver(scan);
    observer.observe(host, { childList: true, subtree: true });
    scan();
    return () => observer.disconnect();
  }, [host]);
  const L = runtime();
  if (!L) return null;
  return <>
    {targets.table ? createPortal(<TableGrid L={L} />, targets.table) : null}
    {targets.columns ? createPortal(<ColumnsPicker L={L} />, targets.columns) : null}
    {targets.highlight ? createPortal(<HighlightPalette L={L} />, targets.highlight) : null}
    {targets.borders ? createPortal(<BorderMenu L={L} />, targets.borders) : null}
    {targets.borderStyle ? createPortal(<BorderStyleMenu L={L} />, targets.borderStyle) : null}
    {targets.borderWeight ? createPortal(<BorderWeightMenu L={L} />, targets.borderWeight) : null}
    {targets.objectWeight ? createPortal(<ObjectWeightMenu L={L} />, targets.objectWeight) : null}
    {targets.objectDash ? createPortal(<ObjectDashMenu L={L} />, targets.objectDash) : null}
    {targets.objectArrow ? createPortal(<ObjectArrowMenu L={L} />, targets.objectArrow) : null}
    {targets.background ? createPortal(<BackgroundPalette L={L} />, targets.background) : null}
    {targets.multipage ? createPortal(<MultiPagePicker L={L} />, targets.multipage) : null}
    {targets.findFormat ? createPortal(<FindFormatMenu L={L} />, targets.findFormat) : null}
    {targets.findSpecial ? createPortal(<FindSpecialMenu L={L} />, targets.findSpecial) : null}
  </>;
}
