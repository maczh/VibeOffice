import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import { createPortal } from 'react-dom';
import { SvgArtwork } from '@vibeoffice/core';

interface ShapeChoice { kind: string; label: string; svg: string }
interface ShapeMenuApp {
  reactShapeMenuCatalog?(): ShapeChoice[] | null;
  reactShapeMenuAction?(kind: string): boolean;
  reactShapeMenuCancel?(slot: HTMLElement): void;
}
interface ShapeMenuRuntime { app?: ShapeMenuApp; t?(message: string): string }

function runtime(host: HTMLElement): ShapeMenuRuntime | undefined {
  return (host.ownerDocument.defaultView as (Window & { L?: ShapeMenuRuntime }) | null)?.L;
}

/** React-owned AutoShape and Change AutoShape grids; geometry and mutations remain engine-backed. */
export function PptxShapeGridMenu({ host }: { host: HTMLElement }) {
  const [slot, setSlot] = useState<HTMLElement | null>(null);
  const [activeIndex, setActiveIndex] = useState(0);
  const choiceRefs = useRef<Array<HTMLButtonElement | null>>([]);
  useEffect(() => {
    const attach = () => {
      const next = host.querySelector<HTMLElement>('#react-pptx-shape-grid-root');
      if (next !== slot) setSlot(next);
    };
    const Observer = host.ownerDocument.defaultView?.MutationObserver;
    const observer = Observer ? new Observer(attach) : null;
    observer?.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer?.disconnect();
      if (slot) runtime(host)?.app?.reactShapeMenuCancel?.(slot);
    };
  }, [host, slot]);
  useEffect(() => { setActiveIndex(0); }, [slot]);
  const L = runtime(host);
  if (!slot) return null;
  const choices = L?.app?.reactShapeMenuCatalog?.() || [];
  const t = (message: string) => L?.t?.(message) || message;
  const active = Math.min(activeIndex, Math.max(0, choices.length - 1));
  const onKeyDown = (event: KeyboardEvent<HTMLButtonElement>) => {
    const current = Math.max(0, choiceRefs.current.indexOf(host.ownerDocument.activeElement as HTMLButtonElement));
    const last = choices.length - 1;
    let next = current;
    if (event.key === 'ArrowLeft') next = current % 8 === 0 ? current : current - 1;
    else if (event.key === 'ArrowRight') next = current % 8 === 7 || current === last ? current : current + 1;
    else if (event.key === 'ArrowUp') next = Math.max(0, current - 8);
    else if (event.key === 'ArrowDown') next = Math.min(last, current + 8);
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = last;
    else if (event.key === 'Enter' || event.key === ' ') { event.stopPropagation(); return; }
    else return;
    event.preventDefault();
    event.stopPropagation();
    if (next >= 0 && next < choices.length) {
      setActiveIndex(next);
      choiceRefs.current[next]?.focus();
    }
  };
  return createPortal(<div className="shape-grid" data-menu-keyboard-owner="react" role="group" aria-label={t('AutoShapes')}>
    {choices.map((choice, index) => <button key={choice.kind} ref={(element) => { choiceRefs.current[index] = element; }} type="button" tabIndex={index === active ? 0 : -1} data-tip={t(choice.label)} aria-label={t(choice.label)} title={t(choice.label)} onFocus={() => setActiveIndex(index)} onKeyDown={onKeyDown} onClick={() => L?.app?.reactShapeMenuAction?.(choice.kind)}><SvgArtwork markup={choice.svg} owner="pptx-shape-grid" /></button>)}
  </div>, slot);
}
