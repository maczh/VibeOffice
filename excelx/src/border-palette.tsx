import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import { SvgArtwork } from '@vibeoffice/core';
import { useExcelxRuntime } from './editor-host.js';
import './border-palette.css';

type Catalog = { left: number; top: number; items: { id: string; icon: string; label: string }[] };
type Runtime = { app?: { reactBorderPaletteCatalog?(): Catalog | null; reactBorderPaletteAction?(id?: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };

/** React-owned Excelx border presets; workbook border transactions remain in the app service. */
export function ExcelxBorderPalette() {
  const runtime = useExcelxRuntime<Runtime>();
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const rootRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime?.app?.reactBorderPaletteCatalog?.() || null);
    const off = runtime?.bus?.on('excelx-border-palette', refresh);
    return () => { off?.(); runtime?.app?.reactBorderPaletteAction?.(); };
  }, [runtime]);
  useEffect(() => {
    if (!catalog) return;
    rootRef.current?.querySelector<HTMLButtonElement>('button')?.focus({ preventScroll: true });
    const close = () => runtime?.app?.reactBorderPaletteAction?.();
    const pointer = (event: PointerEvent) => { if (rootRef.current && !rootRef.current.contains(event.target as Node)) close(); };
    const key = (event: globalThis.KeyboardEvent) => { if (event.key === 'Escape' && rootRef.current?.contains(event.target as Node)) { event.preventDefault(); event.stopPropagation(); close(); } };
    const ownerDocument = rootRef.current?.ownerDocument;
    if (!ownerDocument) return;
    ownerDocument.addEventListener('pointerdown', pointer, true);
    ownerDocument.addEventListener('keydown', key, true);
    return () => { ownerDocument.removeEventListener('pointerdown', pointer, true); ownerDocument.removeEventListener('keydown', key, true); };
  }, [catalog, runtime]);
  if (!catalog) return null;
  const translate = runtime?.t || ((value: string) => value);
  const choose = (id: string) => runtime?.app?.reactBorderPaletteAction?.(id);
  const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    const buttons = [...(rootRef.current?.querySelectorAll<HTMLButtonElement>('button') || [])];
    const current = buttons.indexOf(rootRef.current?.ownerDocument.activeElement as HTMLButtonElement);
    if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); choose(''); return; }
    const delta = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : event.key === 'ArrowDown' ? 4 : event.key === 'ArrowUp' ? -4 : 0;
    if (!delta) return;
    event.preventDefault(); event.stopPropagation();
    buttons[(current + delta + buttons.length) % buttons.length]?.focus();
  };
  return <div ref={rootRef} className="excelx-border-palette" role="group" aria-label={translate('Borders')} onKeyDown={onKeyDown} style={{ left: catalog.left, top: catalog.top }}>
    <div className="excelx-border-palette__grid">{catalog.items.map((item) => <button key={item.id} type="button" aria-label={translate(item.label)} title={translate(item.label)} onClick={() => choose(item.id)}><SvgArtwork markup={item.icon} owner="border-preset" /></button>)}</div>
    <button className="excelx-border-palette__draw" type="button" onClick={() => choose('draw')}>{translate('Draw Borders...')}</button>
  </div>;
}
