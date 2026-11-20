import { useEffect, useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface PptxGridGuide { axis: 'x' | 'y'; position: number }
interface PptxGridCatalog {
  visible: boolean;
  backgroundImage: string;
  backgroundSize: string;
  backgroundPosition: string;
  guides: PptxGridGuide[];
}
interface GridRuntime {
  ed?: { reactGridCatalog?(): PptxGridCatalog | null };
  bus?: { on(event: string, listener: () => void): () => void };
}

interface SlideBackground {
  target: HTMLElement;
  image: string;
  color: string;
  repeat: string;
  position: string;
  size: string;
}

const sourceBackgroundStyles = new WeakMap<HTMLElement, Omit<SlideBackground, 'target'>>();

function slideBackground(target: HTMLElement | null): SlideBackground | null {
  if (!target) return null;
  let source = sourceBackgroundStyles.get(target);
  if (!source) {
    source = {
      image: target.style.backgroundImage,
      color: target.style.backgroundColor,
      repeat: target.style.backgroundRepeat,
      position: target.style.backgroundPosition,
      size: target.style.backgroundSize,
    };
    sourceBackgroundStyles.set(target, source);
  }
  return { target, ...source };
}

function runtime(): GridRuntime | undefined {
  return (window as Window & { L?: GridRuntime }).L;
}

/** React owns the grid texture and guide nodes; the engine retains guide coordinates and snap behavior. */
export function PptxGridLayer({ host }: { host: HTMLElement }) {
  const [catalog, setCatalog] = useState<PptxGridCatalog | null>(() => runtime()?.ed?.reactGridCatalog?.() || null);
  useEffect(() => {
    let offGrid: (() => void) | undefined;
    let offSlide: (() => void) | undefined;
    const refresh = () => {
      setCatalog(runtime()?.ed?.reactGridCatalog?.() || null);
    };
    const attach = () => {
      offGrid?.();
      offSlide?.();
      offGrid = runtime()?.bus?.on('pptx-grid', refresh);
      offSlide = runtime()?.bus?.on('slide-rendered', refresh);
      refresh();
    };
    host.addEventListener('vibeoffice-engine-runtime-ready', attach);
    attach();
    return () => {
      host.removeEventListener('vibeoffice-engine-runtime-ready', attach);
      offGrid?.();
      offSlide?.();
    };
  }, [host]);
  const background = slideBackground(host.querySelector<HTMLElement>('.ed-host > .sl.sl-edit > .sl-bg'));
  useLayoutEffect(() => {
    const target = background?.target;
    if (!target) return;
    const original = {
      image: target.style.backgroundImage,
      color: target.style.backgroundColor,
    };
    target.style.backgroundImage = 'none';
    target.style.backgroundColor = 'transparent';
    return () => {
      target.style.backgroundImage = original.image;
      target.style.backgroundColor = original.color;
    };
  }, [background?.target, background?.image, background?.color]);
  const target = host.querySelector<HTMLElement>('.ed-grid');
  return <>
    {background && createPortal(<div data-react-owned="pptx-slide-background" style={{ position: 'absolute', inset: 0, backgroundImage: background.image, backgroundColor: background.color, backgroundRepeat: background.repeat, backgroundPosition: background.position, backgroundSize: background.size }} />, background.target)}
    {catalog?.visible && target ? createPortal(
      <div className="pptx-react-grid" data-react-owned="pptx-grid" style={{ position: 'absolute', inset: 0, backgroundImage: catalog.backgroundImage, backgroundSize: catalog.backgroundSize, backgroundPosition: catalog.backgroundPosition }}>
        {catalog.guides.map(({ axis, position }) => <div key={axis} className={`guide ${axis === 'x' ? 'gv' : 'gh'}`} data-react-owned="pptx-guide" style={axis === 'x' ? { left: position } : { top: position }} />)}
      </div>,
      target,
    ) : null}
  </>;
}
