import { useEffect, useRef, useState, type KeyboardEvent } from 'react';
import { createPortal } from 'react-dom';
import { SvgArtwork } from '@vibeoffice/core';
import { useExcelxRuntime } from './editor-host';

interface ShapeChoice { kind: string; label: string; svg: string }
interface DrawingRuntime {
  drawing?: {
    reactShapeMenuCatalog?(): ShapeChoice[];
    reactShapeMenuAction?(kind: string): void;
  };
  t?(message: string): string;
}

/** React-owned Excelx AutoShape grid; the workbook drawing engine retains geometry and tool state. */
export function ExcelxShapeMenu({ host }: { host: HTMLElement }) {
  const runtime = useExcelxRuntime<DrawingRuntime>();
  const [slot, setSlot] = useState<HTMLElement | null>(null);
  const [activeIndex, setActiveIndex] = useState(0);
  const buttonRefs = useRef<Array<HTMLButtonElement | null>>([]);

  useEffect(() => {
    const scan = () => {
      const next = host.querySelector<HTMLElement>('#react-excelx-shape-menu-root');
      setSlot((current) => current === next ? current : next);
    };
    const Observer = host.ownerDocument.defaultView?.MutationObserver;
    const observer = Observer ? new Observer(scan) : null;
    observer?.observe(host, { childList: true, subtree: true });
    scan();
    return () => observer?.disconnect();
  }, [host]);

  useEffect(() => { setActiveIndex(0); }, [slot]);
  if (!slot) return null;
  const choices = runtime?.drawing?.reactShapeMenuCatalog?.() || [];
  const t = (message: string) => runtime?.t?.(message) || message;
  const active = Math.min(activeIndex, Math.max(0, choices.length - 1));
  const onKeyDown = (event: KeyboardEvent<HTMLButtonElement>) => {
    const currentIndex = buttonRefs.current.indexOf(event.currentTarget);
    if (currentIndex < 0) return;
    const lastIndex = choices.length - 1;
    const next = event.key === 'ArrowLeft' ? currentIndex % 8 === 0 ? currentIndex : currentIndex - 1
      : event.key === 'ArrowRight' ? currentIndex === lastIndex || currentIndex % 8 === 7 ? currentIndex : currentIndex + 1
        : event.key === 'ArrowUp' ? Math.max(0, currentIndex - 8)
          : event.key === 'ArrowDown' ? Math.min(lastIndex, currentIndex + 8)
        : event.key === 'Home' ? 0
          : event.key === 'End' ? lastIndex : -1;
    if (next < 0) return;
    event.preventDefault();
    event.stopPropagation();
    setActiveIndex(next);
    buttonRefs.current[next]?.focus();
  };

  return createPortal(<div className="shape-grid" role="group" aria-label={t('AutoShapes')}>
    {choices.map((choice, index) => <button key={choice.kind} ref={(element) => { buttonRefs.current[index] = element; }} type="button"
      tabIndex={index === active ? 0 : -1} data-tip={t(choice.label)} aria-label={t(choice.label)} title={t(choice.label)}
      onFocus={() => setActiveIndex(index)} onKeyDown={onKeyDown} onClick={() => runtime?.drawing?.reactShapeMenuAction?.(choice.kind)}>
      <SvgArtwork markup={choice.svg} owner="excelx-shape-menu" /></button>)}
  </div>, slot);
}
