import { useEffect, useRef, useState, type KeyboardEvent as ReactKeyboardEvent } from 'react';
import { SvgArtwork } from '@vibeoffice/core';
import { useExcelxRuntime } from './editor-host.js';
import './chart-type-menu.css';

type Catalog = { left: number; top: number; anchor?: HTMLElement | null; items: { kind: string; label: string; icon: string }[] };
type Runtime = { chartWizard?: { typeMenuCatalog?(): Catalog | null; typeMenuAction?(kind?: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };

/** Keyboard- and pointer-accessible React chart-type picker for selected chart drawings. */
export function ExcelxChartTypeMenu() {
  const runtime = useExcelxRuntime<Runtime>();
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [active, setActive] = useState(0);
  const rootRef = useRef<HTMLDivElement>(null);
  const returnFocus = useRef<HTMLElement | null>(null);
  useEffect(() => { const refresh = () => { const next = runtime?.chartWizard?.typeMenuCatalog?.() || null; setCatalog(next); if (next) setActive(0); }; const off = runtime?.bus?.on('excelx-chart-type-menu', refresh); return () => { off?.(); runtime?.chartWizard?.typeMenuAction?.(); }; }, [runtime]);
  useEffect(() => { if (catalog) { returnFocus.current = catalog.anchor?.isConnected ? catalog.anchor : null; rootRef.current?.focus({ preventScroll: true }); } }, [catalog]);
  const dismiss = (kind?: string) => {
    const anchor = returnFocus.current;
    runtime?.chartWizard?.typeMenuAction?.(kind);
    if (anchor?.isConnected) anchor.focus({ preventScroll: true });
  };
  useEffect(() => {
    if (!catalog) return;
    const close = () => dismiss();
    const pointer = (event: PointerEvent) => { if (rootRef.current && !rootRef.current.contains(event.target as Node)) close(); };
    const key = (event: KeyboardEvent) => { if (rootRef.current?.contains(event.target as Node) && event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(); } };
    const ownerDocument = rootRef.current?.ownerDocument;
    if (!ownerDocument) return;
    ownerDocument.addEventListener('pointerdown', pointer, true); ownerDocument.addEventListener('keydown', key, true);
    return () => { ownerDocument.removeEventListener('pointerdown', pointer, true); ownerDocument.removeEventListener('keydown', key, true); };
  }, [catalog, runtime]);
  if (!catalog) return null;
  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (!catalog.items.length) return;
    if (event.key === 'ArrowDown' || event.key === 'ArrowRight') setActive((index) => (index + 1) % catalog.items.length);
    else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') setActive((index) => (index - 1 + catalog.items.length) % catalog.items.length);
    else if (event.key === 'Home') setActive(0);
    else if (event.key === 'End') setActive(catalog.items.length - 1);
    else if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); event.stopPropagation(); dismiss(catalog.items[active]?.kind); return; }
    else return;
    event.preventDefault();
    event.stopPropagation();
  };
  return <div ref={rootRef} className="excelx-chart-type-menu" role="listbox" aria-label={runtime?.t?.('Chart type') || 'Chart type'} aria-activedescendant={catalog.items[active] ? `excelx-chart-type-option-${active}` : undefined} tabIndex={0} onKeyDown={onKeyDown} style={{ left: catalog.left, top: catalog.top }}>
    {catalog.items.map((item, index) => <button type="button" id={`excelx-chart-type-option-${index}`} role="option" aria-label={runtime?.t?.(item.label) || item.label} aria-selected={active === index} tabIndex={-1} key={item.kind} onPointerEnter={() => setActive(index)} onClick={() => dismiss(item.kind)}><SvgArtwork markup={item.icon} owner="chart-type-option" />{runtime?.t?.(item.label) || item.label}</button>)}
  </div>;
}
