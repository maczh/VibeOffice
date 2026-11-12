import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

type PrintOptions = { what: 'slides' | 'handouts' | 'notes' | 'outline'; per: number; color: 'color' | 'gray' | 'bw'; frame: boolean };
interface PrintRuntime {
  app?: { printOpts: PrintOptions; prevView?: string; setView(view: string): void; previewPDF(): Promise<void> };
  panes?: { preview?: { reactRoot?: HTMLElement; reactMounted?: boolean; render(): void } };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): PrintRuntime | undefined { return (window as Window & { L?: PrintRuntime }).L; }
function label(L: PrintRuntime, text: string): string { return L.t?.(text) || text; }

/** React-owned Print Preview controls; page composition remains supplied by the presentation engine. */
export function PptxPrintPreviewLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [, refresh] = useState(0);
  useEffect(() => {
    let slot: HTMLElement | null = null;
    const attach = () => {
      const target = host.querySelector<HTMLElement>('#react-preview-controls-root');
      if (!target || target === slot) return;
      slot?.classList.remove('vibeoffice-react-owned');
      slot = target;
      slot.classList.add('vibeoffice-react-owned');
      const pane = runtime()?.panes?.preview;
      if (pane) pane.reactMounted = true;
      setMount(slot);
    };
    const off = runtime()?.bus?.on('pptx-preview-render', () => refresh((value) => value + 1));
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer.disconnect(); off?.();
      slot?.classList.remove('vibeoffice-react-owned');
      const pane = runtime()?.panes?.preview;
      if (pane) pane.reactMounted = false;
    };
  }, [host]);
  const L = runtime();
  const app = L?.app;
  const options = app?.printOpts;
  void refresh;
  if (!mount || !L || !app || !options) return null;
  const canPrint = window.self === window.top || !!(window as Window & { VO?: { reactDirectHost?: unknown } }).VO?.reactDirectHost;
  const update = (change: Partial<PrintOptions>) => { Object.assign(options, change); L.panes?.preview?.render(); };
  return createPortal(<>
    {canPrint && <button className="btn small" type="button" onClick={() => { document.body.classList.add('printing'); window.print(); setTimeout(() => document.body.classList.remove('printing'), 500); }}>{label(L, 'Print...')}</button>}
    <label htmlFor="react-print-what">{label(L, 'Print What:')}</label>
    <select id="react-print-what" value={options.what} onChange={(event) => update({ what: event.currentTarget.value as PrintOptions['what'] })}>
      <option value="slides">{label(L, 'Slides')}</option><option value="handouts">{label(L, 'Handouts')}</option><option value="notes">{label(L, 'Notes Pages')}</option><option value="outline">{label(L, 'Outline View')}</option>
    </select>
    <label htmlFor="react-print-per">{label(L, 'Per page:')}</label>
    <select id="react-print-per" value={options.per} disabled={options.what !== 'handouts'} onChange={(event) => update({ per: Number(event.currentTarget.value) })}>
      {[1, 2, 3, 4, 6, 9].map((count) => <option key={count} value={count}>{count}</option>)}
    </select>
    <label htmlFor="react-print-color">{label(L, 'Color')}</label>
    <select id="react-print-color" value={options.color} onChange={(event) => update({ color: event.currentTarget.value as PrintOptions['color'] })}>
      <option value="color">{label(L, 'Color')}</option><option value="gray">{label(L, 'Grayscale')}</option><option value="bw">{label(L, 'Pure Black and White')}</option>
    </select>
    <label className="ck"><input type="checkbox" checked={options.frame} onChange={(event) => update({ frame: event.currentTarget.checked })} /><span>{label(L, 'Frame slides')}</span></label>
    <button className="btn small" type="button" onClick={() => { void app.previewPDF(); }}>{label(L, 'Save as PDF...')}</button>
    <button className="btn small" type="button" onClick={() => app.setView(app.prevView || 'normal')}>{label(L, 'Close')}</button>
    {!canPrint && <span className="tp-note">{label(L, 'To print, save these pages as a PDF and print the PDF.')}</span>}
  </>, mount);
}
