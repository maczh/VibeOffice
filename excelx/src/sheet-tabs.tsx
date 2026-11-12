import { useEffect, useRef, useState, type CSSProperties, type KeyboardEvent, type PointerEvent } from 'react';
import { createPortal } from 'react-dom';

interface SheetTab { index: number; name: string; color: string | null; active: boolean; selected: boolean }
interface SheetsApi {
  sheetTabs(): SheetTab[];
  activateSheet(index: number): void;
  toggleSheetSel(index: number): void;
  rangeSheetSel(index: number): void;
  openSheetTabMenu(event: { clientX: number; clientY: number }, index: number): void;
  openSheetNavigationMenu(event: { clientX: number; clientY: number }): void;
  scrollSheetTabs(destination: 'first' | 'prev' | 'next' | 'last'): void;
  renameSheet(index: number, name: string): boolean;
  reorderSheet(from: number, to: number | null, copy: boolean): void;
}
interface SheetRuntime { app?: SheetsApi; bus?: { on(event: string, listener: () => void): () => void } }
interface Gesture { pointerId: number; from: number; startX: number; moved: boolean; target: number | null }
function runtime(): SheetRuntime | undefined { return (window as Window & { L?: SheetRuntime }).L; }

/** React-owned worksheet tabs; workbook operations and tab context menus remain engine commands. */
export function ExcelxSheetTabsLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [navMount, setNavMount] = useState<HTMLElement | null>(null);
  const [tabs, setTabs] = useState<SheetTab[]>([]);
  const [renaming, setRenaming] = useState<number | null>(null);
  const [nameDraft, setNameDraft] = useState('');
  const [drag, setDrag] = useState<{ from: number; target: number | null } | null>(null);
  const gesture = useRef<Gesture | null>(null);
  useEffect(() => {
    let off: (() => void) | undefined;
    let attached: HTMLElement | null = null;
    const attach = () => {
      const target = host.querySelector<HTMLElement>('#tabs');
      if (!target || target === attached) return;
      attached = target;
      setMount(target);
      setNavMount(host.querySelector<HTMLElement>('#tabnav'));
      off?.();
      off = runtime()?.bus?.on('sheet-tabs', () => setTabs(runtime()?.app?.sheetTabs?.() || []));
      setTabs(runtime()?.app?.sheetTabs?.() || []);
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => { observer.disconnect(); off?.(); };
  }, [host]);
  useEffect(() => {
    const active = mount?.querySelector<HTMLElement>('.stab.on');
    if (!active || !mount) return;
    const item = active.getBoundingClientRect(), viewport = mount.getBoundingClientRect();
    if (item.right > viewport.right || item.left < viewport.left) mount.scrollLeft += item.left - viewport.left - 20;
  }, [mount, tabs]);
  if (!mount) return null;

  const finishRename = (tab: SheetTab, save: boolean) => {
    if (save && !runtime()?.app?.renameSheet(tab.index, nameDraft)) {
      const input = document.querySelector<HTMLInputElement>(`#tabs input[data-sheet-index="${tab.index}"]`);
      input?.focus(); input?.select();
      return;
    }
    setRenaming(null);
  };
  const pointerDown = (event: PointerEvent<HTMLDivElement>, tab: SheetTab) => {
    if (event.target instanceof HTMLInputElement || event.button === 2) {
      if (event.button === 2 && !tab.selected && !tab.active) runtime()?.app?.activateSheet(tab.index);
      return;
    }
    event.preventDefault();
    if (event.button !== 0) return;
    if (event.ctrlKey || event.metaKey) { runtime()?.app?.toggleSheetSel(tab.index); return; }
    if (event.shiftKey) { runtime()?.app?.rangeSheetSel(tab.index); return; }
    runtime()?.app?.activateSheet(tab.index);
    setDrag({ from: tab.index, target: null });
    gesture.current = { pointerId: event.pointerId, from: tab.index, startX: event.clientX, moved: false, target: null };
  };
  const move = (event: PointerEvent<HTMLDivElement>) => {
    const active = gesture.current;
    if (!active || active.pointerId !== event.pointerId) return;
    if (!active.moved && Math.abs(event.clientX - active.startX) < 6) return;
    active.moved = true;
    const element = document.elementFromPoint(event.clientX, event.clientY)?.closest<HTMLElement>('.stab');
    const rect = element?.getBoundingClientRect();
    active.target = element && rect && event.clientX < rect.left + rect.width / 2 ? Number(element.dataset.sheetIndex) : null;
    setDrag({ from: active.from, target: active.target });
  };
  const up = (event: PointerEvent<HTMLDivElement>) => {
    const active = gesture.current;
    if (!active || active.pointerId !== event.pointerId) return;
    gesture.current = null;
    setDrag(null);
    if (active.moved) runtime()?.app?.reorderSheet(active.from, active.target, event.ctrlKey || event.metaKey);
  };
  return <>
    {createPortal(tabs.map((tab) => {
    const style = tab.color ? { '--tabc': tab.color } as CSSProperties : undefined;
    const tabClass = `stab${tab.active ? ' on' : ''}${tab.selected && !tab.active ? ' sel' : ''}${tab.color ? ' colored' : ''}${drag?.target === tab.index ? ' drop-before' : ''}`;
    return <div key={tab.index} className={tabClass} role="tab" aria-selected={tab.active} draggable={false} style={style} data-sheet-index={tab.index} onPointerDown={(event) => pointerDown(event, tab)} onPointerMove={move} onPointerUp={up} onPointerCancel={() => { gesture.current = null; setDrag(null); }} onDoubleClick={() => { setRenaming(tab.index); setNameDraft(tab.name); }} onContextMenu={(event) => { event.preventDefault(); runtime()?.app?.openSheetTabMenu(event, tab.index); }}>
      {renaming === tab.index
        ? <input autoFocus data-sheet-index={tab.index} value={nameDraft} style={{ position: 'relative', zIndex: 2, width: '100%', minWidth: 60, height: 15, font: '11px Tahoma,sans-serif', border: '1px solid #000', padding: '0 2px' }} onChange={(event) => setNameDraft(event.currentTarget.value)} onKeyDown={(event: KeyboardEvent<HTMLInputElement>) => { event.stopPropagation(); if (event.key === 'Enter') finishRename(tab, true); if (event.key === 'Escape') finishRename(tab, false); }} onBlur={() => finishRename(tab, true)} />
        : <span translate="no">{tab.name}</span>}
    </div>;
    }), mount, 'excelx-sheet-tabs')}
    {navMount && createPortal(<>
      {(['first', 'prev', 'next', 'last'] as const).map((destination) => <button key={destination} type="button" tabIndex={-1} aria-label={`${destination.charAt(0).toUpperCase()}${destination.slice(1)} sheet`} data-tip={`${destination.charAt(0).toUpperCase()}${destination.slice(1)} sheet`} data-react-sheet-navigation={destination} onClick={() => runtime()?.app?.scrollSheetTabs(destination)} onContextMenu={(event) => { event.preventDefault(); runtime()?.app?.openSheetNavigationMenu(event); }}>
        <svg width="9" height="9" viewBox="0 0 9 9" aria-hidden="true">{destination === 'first' ? <><path d="M1.5 1v7" stroke="#000" /><path d="M7 1.5L3 4.5 7 7.5z" fill="#000" /></> : destination === 'prev' ? <path d="M6.5 1.5L2.5 4.5 6.5 7.5z" fill="#000" /> : destination === 'next' ? <path d="M2.5 1.5L6.5 4.5 2.5 7.5z" fill="#000" /> : <><path d="M7.5 1v7" stroke="#000" /><path d="M2 1.5L6 4.5 2 7.5z" fill="#000" /></>}</svg>
      </button>)}
    </>, navMount, 'excelx-sheet-tab-navigation')}
  </>;
}
