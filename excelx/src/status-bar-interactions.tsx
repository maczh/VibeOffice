import { useEffect, useState, type MouseEvent as ReactMouseEvent } from 'react';
import { createPortal } from 'react-dom';

interface StatusSnapshot { mode: string; text: string; calculation: string; end: boolean }
interface ExcelxRuntime {
  app?: { opts?: { autoCalcFn?: string }; saveOpts?(): void; updateStatus?(): void };
  ui?: { openMenu(items: unknown[], at: { x: number; y: number }): void };
  statusSnapshot?(): StatusSnapshot;
  grid?: { scrollLock?: boolean };
  bus?: { on(event: string, listener: (snapshot?: StatusSnapshot) => void): () => void };
  N?(value: string): string;
}

function runtime(): ExcelxRuntime | undefined {
  return (window as Window & { L?: ExcelxRuntime }).L;
}

/** Attach the status-bar calculation menu through React's editor-mount lifetime. */
export function ExcelxStatusBarInteractions({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [snapshot, setSnapshot] = useState<StatusSnapshot>({ mode: 'Ready', text: '', calculation: '', end: false });
  const [locks, setLocks] = useState({ caps: false, num: false, scroll: false });
  useEffect(() => {
    const attach = () => setMount(host.querySelector<HTMLElement>('#statusbar'));
    const refresh = (next?: StatusSnapshot) => setSnapshot(next || runtime()?.statusSnapshot?.() || { mode: 'Ready', text: '', calculation: '', end: false });
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    refresh();
    const bus = runtime()?.bus;
    const offStatus = bus?.on('excelx-statusbar', refresh);
    const offReady = bus?.on('app-ready', refresh);
    const keydown = (event: KeyboardEvent) => {
      if (!event.getModifierState) return;
      const caps = event.getModifierState('CapsLock');
      const num = event.getModifierState('NumLock');
      const scroll = event.getModifierState('ScrollLock');
      runtime()?.grid && (runtime()!.grid!.scrollLock = scroll);
      setLocks({ caps, num, scroll });
    };
    host.addEventListener('keydown', keydown, true);
    return () => {
      observer.disconnect();
      offStatus?.();
      offReady?.();
      host.removeEventListener('keydown', keydown, true);
    };
  }, [host]);
  if (!mount) return null;
  const label = (value: string) => runtime()?.N?.(value) || value;
  const openMenuAt = (x: number, y: number) => {
    const L = runtime();
    const app = L?.app;
    if (!L?.ui || !app) return;
    const functions = [['NONE', '&None'], ['AVERAGE', '&Average'], ['COUNTA', '&Count'], ['COUNT', 'C&ount Nums'], ['MAX', '&Max'], ['MIN', 'M&in'], ['SUM', '&Sum']] as const;
    L.ui.openMenu(functions.map(([key, itemLabel]) => ({ label: itemLabel, radio: true, checked: () => app.opts?.autoCalcFn === key, run: () => { if (app.opts) app.opts.autoCalcFn = key; app.saveOpts?.(); app.updateStatus?.(); } })), { x, y });
  };
  const open = (event: ReactMouseEvent<HTMLDivElement>) => openMenuAt(event.clientX, event.clientY - 150);
  return createPortal(<>
    <div className="sb-cell mode">{snapshot.mode}</div>
    <div className="sb-cell grow" data-react-status />
    <div className="sb-cell calc" data-tip="Right-click to choose the AutoCalculate function" role="button" tabIndex={0} aria-label={label('AutoCalculate')} onClick={open} onContextMenu={(event) => { event.preventDefault(); open(event); }} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); const rect = event.currentTarget.getBoundingClientRect(); openMenuAt(rect.right, rect.bottom); } }}>{snapshot.calculation}</div>
    {([['CAPS', locks.caps], ['NUM', locks.num], ['SCRL', locks.scroll], ['END', snapshot.end]] as const).map(([name, active]) => <div key={name} className={`sb-cell ind${active ? '' : ' off'}`}>{label(name)}</div>)}
  </>, mount, 'excelx-status-bar');
}
