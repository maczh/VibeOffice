import { useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';

interface PptxRulerCatalog {
  show: boolean;
  width: number;
  height: number;
  slideWidth: number;
  slideHeight: number;
  scale: number;
  offsetX: number;
  offsetY: number;
  background: string;
  outside: string;
  tick: string;
  ink: string;
}

interface RulerRuntime {
  ed?: { reactRulerCatalog?(): PptxRulerCatalog | null };
  bus?: { on(event: string, listener: () => void): () => void };
}

function runtime(host: HTMLElement): RulerRuntime | undefined {
  return (host.ownerDocument.defaultView as (Window & { L?: RulerRuntime }) | null)?.L;
}

function drawRuler(canvas: HTMLCanvasElement, catalog: PptxRulerCatalog, horizontal: boolean, devicePixelRatio: number) {
  const dpr = devicePixelRatio || 1;
  const size = horizontal ? catalog.width : catalog.height;
  const length = horizontal ? catalog.slideWidth : catalog.slideHeight;
  const offset = horizontal ? catalog.offsetX : catalog.offsetY;
  canvas.width = Math.max(1, size * dpr);
  canvas.height = 18 * dpr;
  canvas.style.width = `${size}px`;
  canvas.style.height = '18px';
  canvas.style.transform = horizontal ? '' : 'rotate(90deg) translate(0,-18px)';
  canvas.style.transformOrigin = horizontal ? '' : '0 0';
  const context = canvas.getContext('2d');
  if (!context) return;
  context.scale(dpr, dpr);
  context.fillStyle = catalog.background;
  context.fillRect(0, 0, size, 18);
  context.fillStyle = catalog.outside;
  context.fillRect(0, 0, offset, 18);
  context.fillRect(offset + length * catalog.scale, 0, size, 18);
  context.strokeStyle = catalog.tick;
  context.fillStyle = catalog.ink;
  context.font = '9px Tahoma, Verdana, sans-serif';
  context.textAlign = 'center';
  const midpoint = length / 2;
  for (let point = 0; point <= length + 0.1; point += 9) {
    const x = Math.round(offset + point * catalog.scale) + 0.5;
    const relative = Math.abs(point - midpoint) / 72;
    const inch = Math.abs(relative - Math.round(relative)) < 0.01;
    const halfInch = Math.abs(relative * 2 - Math.round(relative * 2)) < 0.01;
    const tickLength = inch ? 0 : halfInch ? 6 : 3;
    if (inch) { context.fillText(String(Math.round(relative)), x, 12); continue; }
    context.beginPath();
    context.moveTo(x, 18);
    context.lineTo(x, 18 - tickLength);
    context.stroke();
  }
}

/** React owns the PPTX ruler canvases; the engine supplies slide and look geometry only. */
export function PptxRulers({ host }: { host: HTMLElement }) {
  const [catalog, setCatalog] = useState<PptxRulerCatalog | null>(() => runtime(host)?.ed?.reactRulerCatalog?.() || null);
  const horizontalCanvas = useRef<HTMLCanvasElement>(null);
  const verticalCanvas = useRef<HTMLCanvasElement>(null);
  useLayoutEffect(() => {
    let off: (() => void) | undefined;
    const attach = () => {
      off?.();
      off = runtime(host)?.bus?.on('pptx-rulers', () => setCatalog(runtime(host)?.ed?.reactRulerCatalog?.() || null));
      setCatalog(runtime(host)?.ed?.reactRulerCatalog?.() || null);
    };
    host.addEventListener('vibeoffice-engine-runtime-ready', attach);
    attach();
    return () => {
      host.removeEventListener('vibeoffice-engine-runtime-ready', attach);
      off?.();
    };
  }, [host]);
  useLayoutEffect(() => {
    if (!catalog?.show) return;
    const dpr = host.ownerDocument.defaultView?.devicePixelRatio || 1;
    if (horizontalCanvas.current) drawRuler(horizontalCanvas.current, catalog, true, dpr);
    if (verticalCanvas.current) drawRuler(verticalCanvas.current, catalog, false, dpr);
  }, [catalog, host]);
  const horizontal = host.querySelector<HTMLElement>('.ruler-h');
  const vertical = host.querySelector<HTMLElement>('.ruler-v');
  return <>
    {catalog?.show && horizontal ? createPortal(<canvas data-react-owned="pptx-ruler" ref={horizontalCanvas} />, horizontal) : null}
    {catalog?.show && vertical ? createPortal(<canvas data-react-owned="pptx-ruler" ref={verticalCanvas} />, vertical) : null}
  </>;
}
