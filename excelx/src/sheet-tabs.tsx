import { useEffect, useRef, useState, type CSSProperties, type KeyboardEvent, type MouseEvent as ReactMouseEvent, type PointerEvent } from 'react';
import { createPortal } from 'react-dom';

interface SheetTab { index: number; name: string; color: string | null; active: boolean; selected: boolean }
interface SheetsApi {
  sheetTabs(): SheetTab[];
  activateSheet(index: number): void;
  selectedSheetsIdx(): number[];
  toggleSheetSel(index: number): void;
  rangeSheetSel(index: number): void;
  scrollSheetTabs(destination: 'first' | 'prev' | 'next' | 'last'): void;
  renameSheet(index: number, name: string): boolean;
  reorderSheet(from: number, to: number | null, copy: boolean): void;
  deleteSheets?(): void;
  renameSheetInline?(index: number): void;
  selectAllSheets?(): void;
  ungroupSheets?(): void;
}
interface SheetRuntime { app?: SheetsApi; dlg?: { insertSheetDlg?(): void; moveCopySheet?(): void; tabColor?(): void }; grid?: { paint(): void }; ui?: { openMenu(items: unknown[], at: { x: number; y: number }): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string, ...args: unknown[]): string }
interface Gesture { pointerId: number; from: number; startX: number; moved: boolean; target: number | null }
interface ResizeGesture { pointerId: number; startX: number; startWidth: number; before: string }
function runtime(): SheetRuntime | undefined { return (window as Window & { L?: SheetRuntime }).L; }

/** React-owned worksheet tabs; workbook operations and tab context menus remain engine commands. */
export function ExcelxSheetTabsLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [navMount, setNavMount] = useState<HTMLElement | null>(null);
  const [splitMount, setSplitMount] = useState<HTMLElement | null>(null);
  const [tabs, setTabs] = useState<SheetTab[]>([]);
  const [renaming, setRenaming] = useState<number | null>(null);
  const [nameDraft, setNameDraft] = useState('');
  const [drag, setDrag] = useState<{ from: number; target: number | null } | null>(null);
  const gesture = useRef<Gesture | null>(null);
  const resizeGesture = useRef<ResizeGesture | null>(null);
  const [resizing, setResizing] = useState(false);
  useEffect(() => {
    let off: (() => void) | undefined;
    let attached: HTMLElement | null = null;
    let previousRole: string | null = null;
    let previousLabel: string | null = null;
    const attach = () => {
      const target = host.querySelector<HTMLElement>('#tabs');
      if (!target || target === attached) return;
      attached = target;
      setMount(target);
      previousRole = target.getAttribute('role');
      previousLabel = target.getAttribute('aria-label');
      target.setAttribute('role', 'tablist');
      target.setAttribute('aria-label', runtime()?.t?.('Worksheets') || 'Worksheets');
      setNavMount(host.querySelector<HTMLElement>('#tabnav'));
      setSplitMount(host.querySelector<HTMLElement>('#tabsplit'));
      off?.();
      off = runtime()?.bus?.on('sheet-tabs', () => setTabs(runtime()?.app?.sheetTabs?.() || []));
      setTabs(runtime()?.app?.sheetTabs?.() || []);
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer.disconnect(); off?.();
      if (attached) {
        if (previousRole === null) attached.removeAttribute('role'); else attached.setAttribute('role', previousRole);
        if (previousLabel === null) attached.removeAttribute('aria-label'); else attached.setAttribute('aria-label', previousLabel);
      }
    };
  }, [host]);
  useEffect(() => {
    const active = mount?.querySelector<HTMLElement>('.stab.on');
    if (!active || !mount) return;
    const item = active.getBoundingClientRect(), viewport = mount.getBoundingClientRect();
    if (item.right > viewport.right || item.left < viewport.left) mount.scrollLeft += item.left - viewport.left - 20;
  }, [mount, tabs]);
  useEffect(() => {
    if (!resizing) return;
    const escape = (event: globalThis.KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      event.preventDefault(); event.stopPropagation();
      const active = resizeGesture.current;
      if (!active) return;
      resizeGesture.current = null;
      setResizing(false);
      const grip = splitMount?.querySelector<HTMLElement>('.tabsplit-react-grip');
      if (grip?.hasPointerCapture(active.pointerId)) grip.releasePointerCapture(active.pointerId);
      if (mount) { mount.style.width = active.before; runtime()?.grid?.paint(); }
    };
    window.addEventListener('keydown', escape, true);
    return () => window.removeEventListener('keydown', escape, true);
  }, [resizing, mount, splitMount]);
  if (!mount) return null;

  const finishRename = (tab: SheetTab, save: boolean) => {
    if (save && !runtime()?.app?.renameSheet(tab.index, nameDraft)) {
      const input = mount?.querySelector<HTMLInputElement>(`input[data-sheet-index="${tab.index}"]`);
      input?.focus(); input?.select();
      return;
    }
    setRenaming(null);
  };
  const handleTabKeyDown = (event: KeyboardEvent<HTMLDivElement>, tab: SheetTab) => {
    if (renaming === tab.index) return;
    if ((event.ctrlKey || event.metaKey) && event.key === ' ') {
      event.preventDefault();
      runtime()?.app?.toggleSheetSel(tab.index);
      return;
    }
    if (event.shiftKey && event.key === ' ') {
      event.preventDefault();
      runtime()?.app?.rangeSheetSel(tab.index);
      return;
    }
    const current = tabs.findIndex((item) => item.index === tab.index);
    const destination = event.key === 'ArrowRight' ? current + 1
      : event.key === 'ArrowLeft' ? current - 1
        : event.key === 'Home' ? 0
          : event.key === 'End' ? tabs.length - 1 : -1;
    if (destination >= 0) {
      event.preventDefault();
      const next = tabs[destination];
      if (!next) return;
      runtime()?.app?.activateSheet(next.index);
      requestAnimationFrame(() => mount?.querySelector<HTMLElement>(`[data-sheet-index="${next.index}"]`)?.focus());
    } else if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      runtime()?.app?.activateSheet(tab.index);
    }
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
  const resizeMove = (event: PointerEvent<HTMLDivElement>) => {
    const active = resizeGesture.current;
    if (!active || active.pointerId !== event.pointerId || !mount) return;
    const rowWidth = host.querySelector<HTMLElement>('#tabrow')?.offsetWidth || active.startWidth + 120;
    const maxWidth = Math.max(40, rowWidth - 120);
    mount.style.width = `${Math.max(40, Math.min(maxWidth, active.startWidth + event.clientX - active.startX))}px`;
    runtime()?.grid?.paint();
  };
  const stopResize = (cancel: boolean) => {
    const active = resizeGesture.current;
    if (!active) return;
    resizeGesture.current = null;
    setResizing(false);
    if (cancel && mount) { mount.style.width = active.before; runtime()?.grid?.paint(); }
  };
  const resizeDown = (event: PointerEvent<HTMLDivElement>) => {
    if (event.button !== 0 || !mount) return;
    event.preventDefault();
    resizeGesture.current = { pointerId: event.pointerId, startX: event.clientX, startWidth: mount.offsetWidth, before: mount.style.width };
    setResizing(true);
    event.currentTarget.setPointerCapture(event.pointerId);
  };
  const openNavigationMenu = (event: ReactMouseEvent<HTMLButtonElement>) => {
    event.preventDefault();
    runtime()?.ui?.openMenu(tabs.map((tab) => ({ label: tab.name.replace(/&/g, '&&'), raw: true, checked: () => tab.active, radio: true, run: () => runtime()?.app?.activateSheet(tab.index) })), { x: event.clientX, y: event.clientY });
  };
  const openTabMenu = (event: ReactMouseEvent<HTMLDivElement>, index: number) => {
    event.preventDefault();
    const L = runtime();
    const app = L?.app;
    if (!L?.ui || !app) return;
    L.ui.openMenu([
      { label: '&Insert...', run: () => L.dlg?.insertSheetDlg?.() },
      { label: '&Delete', run: () => app.deleteSheets?.() },
      { label: '&Rename', run: () => app.renameSheetInline?.(index) },
      { label: '&Move or Copy...', run: () => L.dlg?.moveCopySheet?.() },
      { label: '&Select All Sheets', run: () => app.selectAllSheets?.() },
      { label: '&Ungroup Sheets', hidden: app.selectedSheetsIdx().length <= 1, run: () => app.ungroupSheets?.() },
      { label: '&Tab Color...', run: () => L.dlg?.tabColor?.() },
      '-',
      { label: '&View Code', disabled: true },
    ], { x: event.clientX, y: event.clientY });
  };
  return <>
    {createPortal(tabs.map((tab) => {
    const style = tab.color ? { '--tabc': tab.color } as CSSProperties : undefined;
    const tabClass = `stab${tab.active ? ' on' : ''}${tab.selected && !tab.active ? ' sel' : ''}${tab.color ? ' colored' : ''}${drag?.target === tab.index ? ' drop-before' : ''}`;
    return <div key={tab.index} className={tabClass} role="tab" aria-selected={tab.active} aria-label={tab.name} tabIndex={tab.active ? 0 : -1} draggable={false} style={style} data-sheet-index={tab.index} onKeyDown={(event) => handleTabKeyDown(event, tab)} onPointerDown={(event) => pointerDown(event, tab)} onPointerMove={move} onPointerUp={up} onPointerCancel={() => { gesture.current = null; setDrag(null); }} onDoubleClick={() => { setRenaming(tab.index); setNameDraft(tab.name); }} onContextMenu={(event) => openTabMenu(event, tab.index)}>
      {renaming === tab.index
        ? <input autoFocus data-sheet-index={tab.index} value={nameDraft} style={{ position: 'relative', zIndex: 2, width: '100%', minWidth: 60, height: 15, font: '11px Tahoma,sans-serif', border: '1px solid #000', padding: '0 2px' }} onChange={(event) => setNameDraft(event.currentTarget.value)} onKeyDown={(event: KeyboardEvent<HTMLInputElement>) => { event.stopPropagation(); if (event.key === 'Enter') finishRename(tab, true); if (event.key === 'Escape') finishRename(tab, false); }} onBlur={() => finishRename(tab, true)} />
        : <span translate="no">{tab.name}</span>}
    </div>;
    }), mount, 'excelx-sheet-tabs')}
    {navMount && createPortal(<>
      {(['first', 'prev', 'next', 'last'] as const).map((destination) => <button key={destination} type="button" tabIndex={-1} aria-label={`${destination.charAt(0).toUpperCase()}${destination.slice(1)} sheet`} data-tip={`${destination.charAt(0).toUpperCase()}${destination.slice(1)} sheet`} data-react-sheet-navigation={destination} onClick={() => runtime()?.app?.scrollSheetTabs(destination)} onContextMenu={openNavigationMenu}>
        <svg width="9" height="9" viewBox="0 0 9 9" aria-hidden="true">{destination === 'first' ? <><path d="M1.5 1v7" stroke="#000" /><path d="M7 1.5L3 4.5 7 7.5z" fill="#000" /></> : destination === 'prev' ? <path d="M6.5 1.5L2.5 4.5 6.5 7.5z" fill="#000" /> : destination === 'next' ? <path d="M2.5 1.5L6.5 4.5 2.5 7.5z" fill="#000" /> : <><path d="M7.5 1v7" stroke="#000" /><path d="M2 1.5L6 4.5 2 7.5z" fill="#000" /></>}</svg>
      </button>)}
    </>, navMount, 'excelx-sheet-tab-navigation')}
    {splitMount && createPortal(<div className="tabsplit-react-grip" style={{ width: '100%', height: '100%', minHeight: 1, cursor: 'col-resize' }} role="separator" aria-orientation="vertical" aria-label={runtime()?.t?.('Resize') || 'Resize'} onPointerDown={resizeDown} onPointerMove={resizeMove} onPointerUp={() => stopResize(false)} onPointerCancel={() => stopResize(true)} onLostPointerCapture={() => stopResize(true)} />, splitMount, 'excelx-sheet-tab-resize')}
  </>;
}
