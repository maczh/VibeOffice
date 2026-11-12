import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { useVibeOffice } from './react-bridge.js';

interface TooltipState {
  text: string;
  x: number;
  y: number;
  untranslated: boolean;
}

interface EditorTooltipProps {
  host: HTMLElement;
}

/** React-owned, editor-scoped tooltip for legacy engine elements carrying data-tip. */
export function EditorTooltip({ host }: EditorTooltipProps) {
  const { translate } = useVibeOffice();
  const [tip, setTip] = useState<TooltipState | null>(null);
  const timer = useRef<number | null>(null);
  const target = useRef<Element | null>(null);
  const tipElement = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const clearTimer = () => {
      if (timer.current !== null) window.clearTimeout(timer.current);
      timer.current = null;
    };
    const hide = () => {
      clearTimer();
      target.current = null;
      setTip(null);
    };
    const over = (event: PointerEvent) => {
      const element = event.target instanceof Element ? event.target.closest('[data-tip]') : null;
      if (element === target.current) return;
      hide();
      if (!element || !host.contains(element) || event.pointerType === 'touch') return;
      const text = element.getAttribute('data-tip');
      if (!text) return;
      target.current = element;
      const { clientX: x, clientY: y } = event;
      timer.current = window.setTimeout(() => {
        timer.current = null;
        if (target.current === element) setTip({ text, x, y, untranslated: !!element.closest('[translate="no"]') });
      }, 650);
    };
    host.addEventListener('pointerover', over);
    host.addEventListener('pointerdown', hide, true);
    host.addEventListener('keydown', hide, true);
    host.addEventListener('vibeoffice-react-hide-tooltip', hide);
    return () => {
      hide();
      host.removeEventListener('pointerover', over);
      host.removeEventListener('pointerdown', hide, true);
      host.removeEventListener('keydown', hide, true);
      host.removeEventListener('vibeoffice-react-hide-tooltip', hide);
    };
  }, [host]);

  useLayoutEffect(() => {
    if (!tip || !tipElement.current) return;
    const bounds = tipElement.current.getBoundingClientRect();
    let left = tip.x + 2;
    let top = tip.y + 20;
    if (left + bounds.width > window.innerWidth - 4) left = window.innerWidth - bounds.width - 4;
    if (top + bounds.height > window.innerHeight - 4) top = tip.y - bounds.height - 6;
    tipElement.current.style.left = `${Math.max(2, left)}px`;
    tipElement.current.style.top = `${Math.max(2, top)}px`;
  }, [tip]);

  if (!tip) return null;
  const text = tip.untranslated ? tip.text : translate(tip.text);
  const engineRoot = host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root]');
  return engineRoot ? createPortal(<div ref={tipElement} className="tooltip vibeoffice-tooltip" role="tooltip" style={{ left: tip.x + 2, top: tip.y + 20 }}>{text}</div>, engineRoot) : null;
}
