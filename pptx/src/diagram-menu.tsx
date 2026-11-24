import { useEffect, useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import { SvgArtwork } from '@vibeoffice/core';

interface DiagramChoice { id: string; label: string; preview: string; selected: boolean }
interface DiagramGroup { id: string; label: string; choices: DiagramChoice[] }
interface DiagramMenuRequest { left: number; bottom: number; anchor?: HTMLElement; kind: string; groups: DiagramGroup[]; moreLabel?: string; choose(id: string): void; more?(): void }
interface DiagramMenuRuntime { t?(message: string): string }

/** React menu shell for the engine's IntelliArt layout, colour, and style catalogs. */
export function PptxDiagramMenu({ host }: { host: HTMLElement }) {
  const [request, setRequest] = useState<DiagramMenuRequest | null>(null);
  const [position, setPosition] = useState({ left: 2, top: 2 });
  const ownerDocument = host.ownerDocument;
  const ownerWindow = ownerDocument.defaultView;
  const L = (ownerWindow as (Window & { L?: DiagramMenuRuntime }) | null)?.L;
  useEffect(() => {
    const open = (event: Event) => {
      const detail = (event as CustomEvent<DiagramMenuRequest>).detail;
      if (detail && Number.isFinite(detail.left) && Number.isFinite(detail.bottom) && Array.isArray(detail.groups) && typeof detail.choose === 'function') setRequest(detail);
    };
    host.addEventListener('vibeoffice-react-diagram-menu', open);
    return () => host.removeEventListener('vibeoffice-react-diagram-menu', open);
  }, [host]);
  useLayoutEffect(() => {
    if (!request) return;
    const panel = host.querySelector<HTMLElement>('.vibeoffice-diagram-menu');
    if (!panel) return;
    const bounds = panel.getBoundingClientRect();
    setPosition({ left: Math.max(2, Math.min(request.left, ownerWindow?.innerWidth ?? bounds.width + 4)), top: Math.max(2, Math.min(request.bottom, ownerWindow?.innerHeight ?? bounds.height + 4)) });
    panel.querySelector<HTMLButtonElement>('button')?.focus({ preventScroll: true });
  }, [host, request]);
  useEffect(() => {
    if (!request) return;
    const dismiss = () => setRequest(null);
    const outside = (event: PointerEvent) => {
      const target = event.target;
      if (target instanceof ownerDocument.defaultView!.Node && !host.querySelector('.vibeoffice-diagram-menu')?.contains(target) && !request.anchor?.contains(target)) dismiss();
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); dismiss(); request.anchor?.focus(); return; }
      if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) return;
      const buttons = [...host.querySelectorAll<HTMLButtonElement>('.vibeoffice-diagram-menu button')];
      const index = buttons.indexOf(ownerDocument.activeElement as HTMLButtonElement);
      if (!buttons.length) return;
      event.preventDefault(); event.stopPropagation();
      const delta = event.key === 'ArrowLeft' || event.key === 'ArrowUp' ? -1 : 1;
      buttons[(index + delta + buttons.length) % buttons.length]?.focus();
    };
    ownerDocument.addEventListener('pointerdown', outside, true);
    ownerDocument.addEventListener('keydown', keydown, true);
    return () => { ownerDocument.removeEventListener('pointerdown', outside, true); ownerDocument.removeEventListener('keydown', keydown, true); };
  }, [host, request]);
  const engineRoot = host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root]');
  if (!request || !engineRoot) return null;
  const close = () => { setRequest(null); request.anchor?.focus(); };
  return createPortal(<div className="menu sa-menu-grid vibeoffice-diagram-menu" data-menu-keyboard-owner="react" role="menu" tabIndex={-1} style={{ left: position.left, top: position.top }}>
    {request.groups.map((group) => <section key={group.id} aria-label={group.label}>
      {group.label ? <div className="sa-sub">{group.label}</div> : null}
      <div className="sa-pick" style={{ gridTemplateColumns: `repeat(${request.kind === 'layout' ? 5 : 4}, 1fr)` }}>
        {group.choices.map((choice) => <button key={choice.id} type="button" className={`sa-cell${choice.selected ? ' on' : ''}`} role="menuitemradio" aria-label={choice.label} aria-checked={choice.selected} data-diagram-choice={choice.id} title={choice.label} onPointerDown={(event) => event.preventDefault()} onClick={() => { close(); request.choose(choice.id); }}>
          <span aria-hidden="true"><SvgArtwork markup={choice.preview} owner="pptx-diagram-menu" /></span>
        </button>)}
      </div>
    </section>)}
    {request.moreLabel && request.more ? <button type="button" className="btn small" role="menuitem" onClick={() => { close(); request.more?.(); }}>{request.moreLabel}</button> : null}
  </div>, engineRoot);
}
