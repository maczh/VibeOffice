import { useEffect, useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface ShadowValue { dx: number; dy: number; blur: number }
interface ShadowChoice { id: string; label: string; value: ShadowValue | null }
interface ShadowMenuRequest { left: number; bottom: number; anchor?: HTMLElement; choices: ShadowChoice[]; choose(value: ShadowValue | null): void }
interface ShadowMenuRuntime { t?(message: string): string }

/** React-owned shadow preset grid; the presentation engine still owns the selected shadow transaction. */
export function PptxShadowStyleMenu({ host }: { host: HTMLElement }) {
  const [request, setRequest] = useState<ShadowMenuRequest | null>(null);
  const [position, setPosition] = useState({ left: 2, top: 2 });
  const L = (window as Window & { L?: ShadowMenuRuntime }).L;
  useEffect(() => {
    const open = (event: Event) => {
      const detail = (event as CustomEvent<ShadowMenuRequest>).detail;
      if (detail && Number.isFinite(detail.left) && Number.isFinite(detail.bottom) && Array.isArray(detail.choices) && typeof detail.choose === 'function') setRequest(detail);
    };
    host.addEventListener('vibeoffice-react-shadow-menu', open);
    return () => host.removeEventListener('vibeoffice-react-shadow-menu', open);
  }, [host]);
  useLayoutEffect(() => {
    if (!request) return;
    const panel = host.querySelector<HTMLElement>('.vibeoffice-shadow-style-menu');
    if (!panel) return;
    const bounds = panel.getBoundingClientRect();
    setPosition({
      left: Math.max(2, Math.min(request.left, window.innerWidth - bounds.width - 2)),
      top: Math.max(2, Math.min(request.bottom, window.innerHeight - bounds.height - 2)),
    });
    panel.querySelector<HTMLButtonElement>('button')?.focus({ preventScroll: true });
  }, [host, request]);
  useEffect(() => {
    if (!request) return;
    const close = () => setRequest(null);
    const outside = (event: PointerEvent) => {
      const target = event.target;
      const panel = host.querySelector('.vibeoffice-shadow-style-menu');
      if (target instanceof Node && !panel?.contains(target) && !request.anchor?.contains(target)) close();
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(); request.anchor?.focus(); return; }
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight' && event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
      const buttons = [...(host.querySelectorAll<HTMLButtonElement>('.vibeoffice-shadow-style-menu button') ?? [])];
      const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
      if (!buttons.length) return;
      event.preventDefault(); event.stopPropagation();
      const delta = event.key === 'ArrowLeft' || event.key === 'ArrowUp' ? -1 : 1;
      buttons[(index + delta + buttons.length) % buttons.length]?.focus();
    };
    document.addEventListener('pointerdown', outside, true);
    document.addEventListener('keydown', keydown, true);
    return () => {
      document.removeEventListener('pointerdown', outside, true);
      document.removeEventListener('keydown', keydown, true);
    };
  }, [host, request]);
  const engineRoot = host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root]');
  if (!request || !engineRoot) return null;
  const close = () => { setRequest(null); request.anchor?.focus(); };
  return createPortal(<div className="menu grid-menu vibeoffice-shadow-style-menu" data-menu-keyboard-owner="react" role="menu" aria-label={L?.t?.('Shadow Styles') || 'Shadow Styles'} tabIndex={-1} style={{ left: position.left, top: position.top }}>
    <div className="shape-grid" style={{ gridTemplateColumns: 'repeat(4,30px)' }}>
      {request.choices.map((choice) => <button key={choice.id} type="button" role="menuitem" data-shadow-id={choice.id} aria-label={choice.label} title={choice.label} onPointerDown={(event) => event.preventDefault()} onClick={() => { close(); request.choose(choice.value); }}>
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          {choice.value ? <rect x={6 + choice.value.dx / 1.5} y={6 + choice.value.dy / 1.5} width="12" height="12" fill="#888" filter={choice.value.blur ? 'blur(1px)' : undefined} /> : null}
          <rect x="6" y="6" width="12" height="12" fill="#fff" stroke="#000" />
          {!choice.value && <path d="M3 21L21 3" stroke="#d33" />}
        </svg>
      </button>)}
    </div>
  </div>, engineRoot);
}
