import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface PptxStatusRuntime {
  app?: { view: string; opts?: { statusBar?: boolean } };
  ed?: { idx: number; masterKind?: string; design(): unknown; zoomPct(): number; setZoom(value: number): void };
  model?: { designName(value?: unknown): string };
  pres?: { slides: unknown[] };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(text: string, ...args: unknown[]): string;
}
function runtime(): PptxStatusRuntime | undefined { return (window as Window & { L?: PptxStatusRuntime }).L; }

/** React-owned PPTX status values and zoom control; slide/layout calculations stay in the engine. */
export function PptxStatusBar({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [, refresh] = useState(0);
  useEffect(() => {
    const attach = () => setMount(host.querySelector<HTMLElement>('#statusbar'));
    attach();
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    const off = runtime()?.bus?.on('pptx-status-refresh', () => refresh((value) => value + 1));
    const offZoom = runtime()?.bus?.on('zoom', () => refresh((value) => value + 1));
    return () => { observer.disconnect(); off?.(); offZoom?.(); };
  }, [host]);
  const L = runtime();
  const app = L?.app;
  const ed = L?.ed;
  const pres = L?.pres;
  if (!mount || !app || !ed || !pres) return null;
  const translate = (text: string, ...args: unknown[]) => L.t?.(text, ...args) || text;
  const slideText = app.view === 'master'
    ? translate(ed.masterKind === 'title' ? 'Title Master' : 'Slide Master')
    : app.view === 'sorter'
      ? translate('Slide Sorter')
      : pres.slides.length
        ? translate('Slide {0} of {1}', ed.idx + 1, pres.slides.length)
        : translate('No slides');
  const zoom = ed.zoomPct();
  return createPortal(<>
    <div className="sb-cell fit" aria-live="polite">{slideText}</div>
    <div className="sb-cell fit">{L.model?.designName?.(ed.design()) || ''}</div>
    <div className="sb-cell grow" data-react-status="" />
    <div className="sb-cell fit">{translate('English (U.S.)')}</div>
    <div className="sb-cell fit sb-zoom"><span>{zoom}%</span><input type="range" min="10" max="400" value={zoom} aria-label={translate('Zoom')} onChange={(event) => ed.setZoom(Number(event.currentTarget.value) / 100)} /></div>
  </>, mount);
}
