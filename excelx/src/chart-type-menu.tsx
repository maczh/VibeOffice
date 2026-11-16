import { useEffect, useRef, useState, type KeyboardEvent as ReactKeyboardEvent } from 'react';
import './chart-type-menu.css';

type Catalog = { left: number; top: number; items: { kind: string; label: string; icon: string }[] };
type Runtime = { chartWizard?: { typeMenuCatalog?(): Catalog | null; typeMenuAction?(kind?: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** Keyboard- and pointer-accessible React chart-type picker for selected chart drawings. */
export function ExcelxChartTypeMenu() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [active, setActive] = useState(0);
  const rootRef = useRef<HTMLDivElement>(null);
  useEffect(() => { const refresh = () => { const next = runtime()?.chartWizard?.typeMenuCatalog?.() || null; setCatalog(next); if (next) setActive(0); }; const off = runtime()?.bus?.on('excelx-chart-type-menu', refresh); return () => { off?.(); runtime()?.chartWizard?.typeMenuAction?.(); }; }, []);
  useEffect(() => { if (catalog) rootRef.current?.focus({ preventScroll: true }); }, [catalog]);
  useEffect(() => {
    if (!catalog) return;
    const close = () => runtime()?.chartWizard?.typeMenuAction?.();
    const pointer = (event: PointerEvent) => { if (rootRef.current && !rootRef.current.contains(event.target as Node)) close(); };
    const key = (event: KeyboardEvent) => { if (rootRef.current?.contains(event.target as Node) && event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(); } };
    document.addEventListener('pointerdown', pointer, true); document.addEventListener('keydown', key, true);
    return () => { document.removeEventListener('pointerdown', pointer, true); document.removeEventListener('keydown', key, true); };
  }, [catalog]);
  if (!catalog) return null;
  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (!catalog.items.length) return;
    if (event.key === 'ArrowDown' || event.key === 'ArrowRight') setActive((index) => (index + 1) % catalog.items.length);
    else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') setActive((index) => (index - 1 + catalog.items.length) % catalog.items.length);
    else if (event.key === 'Home') setActive(0);
    else if (event.key === 'End') setActive(catalog.items.length - 1);
    else if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); event.stopPropagation(); runtime()?.chartWizard?.typeMenuAction?.(catalog.items[active]?.kind); return; }
    else return;
    event.preventDefault();
    event.stopPropagation();
  };
  return <div ref={rootRef} className="excelx-chart-type-menu" role="listbox" aria-label={runtime()?.t?.('Chart type') || 'Chart type'} aria-activedescendant={catalog.items[active] ? `excelx-chart-type-option-${active}` : undefined} tabIndex={0} onKeyDown={onKeyDown} style={{ left: catalog.left, top: catalog.top }}>
    {catalog.items.map((item, index) => <button type="button" id={`excelx-chart-type-option-${index}`} role="option" aria-label={runtime()?.t?.(item.label) || item.label} aria-selected={active === index} tabIndex={-1} key={item.kind} onPointerEnter={() => setActive(index)} onClick={() => runtime()?.chartWizard?.typeMenuAction?.(item.kind)}><span dangerouslySetInnerHTML={{ __html: item.icon }} />{runtime()?.t?.(item.label) || item.label}</button>)}
  </div>;
}
