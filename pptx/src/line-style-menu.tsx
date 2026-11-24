import { useEffect, useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import { SvgArtwork } from '@vibeoffice/core';

interface LineChoice { id: string; label: string; svg: string; value: unknown }
interface LineMenuRequest { left: number; bottom: number; anchor?: HTMLElement; kind: string; choices: LineChoice[]; choose(value: unknown): void }
interface LineMenuRuntime { t?(message: string): string }

/** React-owned line weight/dash/arrow menus; existing PPTX commands own commits. */
export function PptxLineStyleMenu({ host }: { host: HTMLElement }) {
  const [request, setRequest] = useState<LineMenuRequest | null>(null);
  const [position, setPosition] = useState({ left: 2, top: 2 });
  const ownerDocument = host.ownerDocument;
  const ownerWindow = ownerDocument.defaultView;
  const L = (ownerWindow as (Window & { L?: LineMenuRuntime }) | null)?.L;
  useEffect(() => {
    const open = (event: Event) => {
      const detail = (event as CustomEvent<LineMenuRequest>).detail;
      if (detail && Number.isFinite(detail.left) && Number.isFinite(detail.bottom) && Array.isArray(detail.choices) && typeof detail.choose === 'function') setRequest(detail);
    };
    host.addEventListener('vibeoffice-react-line-menu', open);
    return () => host.removeEventListener('vibeoffice-react-line-menu', open);
  }, [host]);
  useLayoutEffect(() => {
    if (!request) return;
    const panel = host.querySelector<HTMLElement>('.vibeoffice-line-style-menu');
    if (!panel) return;
    const bounds = panel.getBoundingClientRect();
    setPosition({ left: Math.max(2, Math.min(request.left, ownerWindow?.innerWidth ?? bounds.width + 4)), top: Math.max(2, Math.min(request.bottom, ownerWindow?.innerHeight ?? bounds.height + 4)) });
    panel.querySelector<HTMLButtonElement>('button')?.focus({ preventScroll: true });
  }, [host, request]);
  useEffect(() => {
    if (!request) return;
    const close = () => setRequest(null);
    const outside = (event: PointerEvent) => {
      const target = event.target;
      if (target instanceof ownerDocument.defaultView!.Node && !host.querySelector('.vibeoffice-line-style-menu')?.contains(target) && !request.anchor?.contains(target)) close();
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(); request.anchor?.focus(); return; }
      if (!['ArrowUp', 'ArrowDown'].includes(event.key)) return;
      const buttons = [...host.querySelectorAll<HTMLButtonElement>('.vibeoffice-line-style-menu button')];
      const index = buttons.indexOf(ownerDocument.activeElement as HTMLButtonElement);
      if (!buttons.length) return;
      event.preventDefault(); event.stopPropagation();
      const delta = event.key === 'ArrowUp' ? -1 : 1;
      buttons[(index + delta + buttons.length) % buttons.length]?.focus();
    };
    ownerDocument.addEventListener('pointerdown', outside, true);
    ownerDocument.addEventListener('keydown', keydown, true);
    return () => { ownerDocument.removeEventListener('pointerdown', outside, true); ownerDocument.removeEventListener('keydown', keydown, true); };
  }, [host, request]);
  const engineRoot = host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root]');
  if (!request || !engineRoot) return null;
  const close = () => { setRequest(null); request.anchor?.focus(); };
  const borderGrid = request.kind === 'table-borders';
  const title = request.kind === 'weight' ? 'Line Style' : request.kind === 'dash' ? 'Dash Style' : request.kind === 'arrow' ? 'Arrow Style' : request.kind === 'table-weight' ? 'Border Width' : request.kind === 'table-dash' ? 'Border Style' : 'Borders';
  return createPortal(<div className={`menu vibeoffice-line-style-menu ${borderGrid ? 'grid-menu shape-grid' : 'style-list'}`} data-menu-keyboard-owner="react" role="menu" aria-label={L?.t?.(title) || title} tabIndex={-1} style={{ left: position.left, top: position.top, ...(borderGrid ? { gridTemplateColumns: 'repeat(4,24px)' } : {}) }}>
    {request.choices.map((choice) => <button key={choice.id} type="button" role="menuitem" data-line-choice={choice.id} aria-label={choice.label} title={choice.label} onPointerDown={(event) => event.preventDefault()} onClick={() => { close(); request.choose(choice.value); }}>
      <span aria-hidden="true"><SvgArtwork markup={choice.svg} owner="pptx-line-style-menu" /></span>{borderGrid || request.kind === 'arrow' && choice.id !== 'more' ? null : <span>{choice.label}</span>}
    </button>)}
  </div>, engineRoot);
}
