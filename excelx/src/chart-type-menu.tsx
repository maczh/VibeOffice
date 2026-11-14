import { useEffect, useRef, useState } from 'react';
import './chart-type-menu.css';

type Catalog = { left: number; top: number; items: { kind: string; label: string; icon: string }[] };
type Runtime = { chartWizard?: { typeMenuCatalog?(): Catalog | null; typeMenuAction?(kind?: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** Keyboard- and pointer-accessible React chart-type picker for selected chart drawings. */
export function ExcelxChartTypeMenu() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const rootRef = useRef<HTMLDivElement>(null);
  useEffect(() => { const refresh = () => setCatalog(runtime()?.chartWizard?.typeMenuCatalog?.() || null); const off = runtime()?.bus?.on('excelx-chart-type-menu', refresh); return () => { off?.(); runtime()?.chartWizard?.typeMenuAction?.(); }; }, []);
  useEffect(() => {
    if (!catalog) return;
    const close = () => runtime()?.chartWizard?.typeMenuAction?.();
    const pointer = (event: PointerEvent) => { if (rootRef.current && !rootRef.current.contains(event.target as Node)) close(); };
    const key = (event: KeyboardEvent) => { if (rootRef.current?.contains(event.target as Node) && event.key === 'Escape') { event.preventDefault(); close(); } };
    document.addEventListener('pointerdown', pointer, true); document.addEventListener('keydown', key, true);
    return () => { document.removeEventListener('pointerdown', pointer, true); document.removeEventListener('keydown', key, true); };
  }, [catalog]);
  if (!catalog) return null;
  return <div ref={rootRef} className="excelx-chart-type-menu" role="listbox" aria-label={runtime()?.t?.('Chart type') || 'Chart type'} style={{ left: catalog.left, top: catalog.top }}>
    {catalog.items.map((item) => <button type="button" role="option" aria-label={runtime()?.t?.(item.label) || item.label} key={item.kind} onClick={() => runtime()?.chartWizard?.typeMenuAction?.(item.kind)}><span dangerouslySetInnerHTML={{ __html: item.icon }} />{runtime()?.t?.(item.label) || item.label}</button>)}
  </div>;
}
