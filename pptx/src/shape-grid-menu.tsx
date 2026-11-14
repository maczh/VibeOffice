import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface ShapeChoice { kind: string; label: string; svg: string }
interface ShapeMenuApp {
  reactShapeMenuCatalog?(): ShapeChoice[] | null;
  reactShapeMenuAction?(kind: string): boolean;
  reactShapeMenuCancel?(slot: HTMLElement): void;
}
interface ShapeMenuRuntime { app?: ShapeMenuApp; t?(message: string): string }

function runtime(): ShapeMenuRuntime | undefined { return (window as Window & { L?: ShapeMenuRuntime }).L; }

/** React-owned AutoShape and Change AutoShape grids; geometry and mutations remain engine-backed. */
export function PptxShapeGridMenu({ host }: { host: HTMLElement }) {
  const [slot, setSlot] = useState<HTMLElement | null>(null);
  useEffect(() => {
    const attach = () => {
      const next = host.querySelector<HTMLElement>('#react-pptx-shape-grid-root');
      if (next !== slot) setSlot(next);
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer.disconnect();
      if (slot) runtime()?.app?.reactShapeMenuCancel?.(slot);
    };
  }, [host, slot]);
  const L = runtime();
  if (!slot) return null;
  const choices = L?.app?.reactShapeMenuCatalog?.() || [];
  const t = (message: string) => L?.t?.(message) || message;
  return createPortal(<div className="shape-grid" role="group" aria-label={t('AutoShapes')}>
    {choices.map((choice) => <button key={choice.kind} type="button" data-tip={t(choice.label)} aria-label={t(choice.label)} title={t(choice.label)} onClick={() => L?.app?.reactShapeMenuAction?.(choice.kind)} dangerouslySetInnerHTML={{ __html: choice.svg }} />)}
  </div>, slot);
}
