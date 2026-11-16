import { useEffect, useRef } from 'react';

interface DragData {
  x0: number;
  y0: number;
  x: number;
  y: number;
  dx: number;
  dy: number;
  moved: boolean;
  data: Record<string, unknown>;
}

interface DragOptions {
  threshold?: number;
  capture?: Element | false;
  cursor?: string | false;
  move?(event: PointerEvent, data: DragData): void;
  up?(event: PointerEvent | null, data: DragData): void;
  cancel?(data: DragData, reason: string): void;
}

interface ActiveDrag {
  id?: number;
  cap: Element | null;
  data: DragData;
  move(event: PointerEvent): void;
  up(event: PointerEvent): void;
  keydown(event: KeyboardEvent): void;
  cancel(reason: string): void;
  end(ok: boolean, event: PointerEvent | null, reason?: string): void;
}

interface DirectDragHost {
  root?: HTMLElement | null;
  startDrag?: (event: PointerEvent, options: DragOptions, onEscape: () => void) => DragData | null;
  dragging?: () => boolean;
  cancelDrag?: () => void;
}

interface DragWindow extends Window {
  VO?: { reactDirectHost?: DirectDragHost };
  L?: { ui?: { brushOn?(): { stop(): void } | null } };
}

interface SwallowState {
  button: number;
  until: number;
  seen: Set<string>;
}

/** Owns direct-host drag sessions, pointer listeners, and their scoped visual state. */
export function EditorDragLayer({ host }: { host: HTMLElement }) {
  const activeRef = useRef<ActiveDrag | null>(null);
  const cleanupRef = useRef<(() => void) | null>(null);
  const cursorStyleRef = useRef<HTMLStyleElement | null>(null);
  const swallowRef = useRef<SwallowState | null>(null);

  useEffect(() => {
    const win = host.ownerDocument.defaultView as DragWindow | null;
    if (!win) return;
    const doc = host.ownerDocument;
    const setCursor = (cursor: string | null) => {
      const engineRoot = win.VO?.reactDirectHost?.root;
      if (!cursor) {
        cursorStyleRef.current?.remove();
        cursorStyleRef.current = null;
        return;
      }
      const styleHost = engineRoot?.isConnected ? engineRoot : doc.head;
      if (!cursorStyleRef.current || cursorStyleRef.current.parentElement !== styleHost) {
        cursorStyleRef.current?.remove();
        cursorStyleRef.current = doc.createElement('style');
        styleHost.appendChild(cursorStyleRef.current);
      }
      const selector = engineRoot?.isConnected
        ? '.vibeoffice-legacy-root[data-vibeoffice-engine-root], .vibeoffice-legacy-root[data-vibeoffice-engine-root] *, .vibeoffice-legacy-root[data-vibeoffice-engine-root] *::before, .vibeoffice-legacy-root[data-vibeoffice-engine-root] *::after'
        : '*, *::before, *::after';
      cursorStyleRef.current.textContent = `${selector} { cursor: ${cursor} !important; }`;
    };
    const cancelActive = (reason: string) => activeRef.current?.cancel(reason);
    const onPointerCancel = () => cancelActive('lost');
    const onBlur = () => cancelActive('blur');
    const onVisibilityChange = () => { if (doc.hidden) cancelActive('blur'); };
    const swallowButton = (button: number) => { swallowRef.current = { button, until: Date.now() + 1500, seen: new Set() }; };
    const onOtherButton = (event: Event) => {
      const mouse = event as MouseEvent;
      if (mouse.button === 0) return;
      if (activeRef.current) { swallowButton(mouse.button); cancelActive('button'); }
      const swallow = swallowRef.current;
      if (!swallow || mouse.button !== swallow.button) return;
      if (Date.now() > swallow.until) { swallowRef.current = null; return; }
      event.preventDefault(); event.stopImmediatePropagation();
      swallow.seen.add(event.type);
      if (swallow.seen.has('mouseup') && swallow.seen.has(swallow.button === 2 ? 'contextmenu' : 'auxclick')) swallowRef.current = null;
    };
    const start = (event: PointerEvent, options: DragOptions, onEscape: () => void) => {
      if (event.button != null && event.button !== 0) return null;
      activeRef.current?.end(false, null, 'new');
      const data: DragData = { x0: event.clientX, y0: event.clientY, x: event.clientX, y: event.clientY, dx: 0, dy: 0, moved: false, data: {} };
      const id = event.pointerId, threshold = options.threshold == null ? 4 : options.threshold;
      const cap = options.capture === false ? null : options.capture || event.target as Element | null;
      let cursor = options.cursor || null;
      if (!cursor && options.cursor !== false) {
        try { const value = win.getComputedStyle(event.target as Element).cursor; if (value && value !== 'auto') cursor = value; } catch { /* no cursor held */ }
      }
      const stop = () => {
        win.removeEventListener('pointermove', live.move, true);
        win.removeEventListener('pointerup', live.up, true);
        win.removeEventListener('keydown', live.keydown, true);
        cap?.removeEventListener('lostpointercapture', onLostCapture);
        if (activeRef.current === live) activeRef.current = null;
        if (cleanupRef.current === stop) cleanupRef.current = null;
      };
      const end = (ok: boolean, upEvent: PointerEvent | null, reason = '') => {
        if (activeRef.current !== live) return;
        activeRef.current = null;
        stop();
        setCursor(null);
        try { if (cap && id != null && cap.hasPointerCapture(id)) cap.releasePointerCapture(id); } catch { /* capture is optional */ }
        if (ok) options.up?.(upEvent, data); else options.cancel?.(data, reason);
      };
      const onLostCapture = () => { if (cap?.isConnected) live.cancel('lost'); };
      const live: ActiveDrag = {
        id, cap, data,
        end,
        move(moveEvent) {
          if (activeRef.current !== live || (id != null && moveEvent.pointerId !== id)) return;
          if (moveEvent.buttons & 6) { swallowButton(moveEvent.buttons & 2 ? 2 : 1); end(false, null, 'button'); return; }
          if (!(moveEvent.buttons & 1)) { end(false, null, 'lost'); return; }
          data.x = moveEvent.clientX; data.y = moveEvent.clientY; data.dx = data.x - data.x0; data.dy = data.y - data.y0;
          if (!data.moved && Math.hypot(data.dx, data.dy) < threshold) return;
          if (!data.moved && cursor) setCursor(cursor);
          data.moved = true;
          options.move?.(moveEvent, data);
        },
        up(upEvent) { if (activeRef.current === live && (id == null || upEvent.pointerId === id) && upEvent.button === 0) end(true, upEvent); },
        keydown(keyEvent) {
          if (activeRef.current !== live || /^(Shift|Control|Alt|Meta|AltGraph|CapsLock)$/.test(keyEvent.key)) return;
          keyEvent.preventDefault(); keyEvent.stopImmediatePropagation();
          if ((keyEvent.ctrlKey || keyEvent.metaKey) && /^[zy]$/i.test(keyEvent.key)) end(false, null, 'undo');
          else if (keyEvent.key === 'Escape') { end(false, null, 'esc'); onEscape(); }
        },
        cancel(reason) { end(false, null, reason); },
      };
      activeRef.current = live;
      cleanupRef.current = stop;
      try { if (cap && id != null && cap.setPointerCapture) cap.setPointerCapture(id); } catch { /* window listeners still follow */ }
      win.addEventListener('pointermove', live.move, true);
      win.addEventListener('pointerup', live.up, true);
      win.addEventListener('keydown', live.keydown, true);
      cap?.addEventListener('lostpointercapture', onLostCapture);
      return data;
    };
    const isDragging = () => !!activeRef.current;
    const cancelDrag = () => cancelActive('api');
    let attached: DirectDragHost | undefined;
    const attach = () => {
      const direct = win.VO?.reactDirectHost;
      if (!direct) return;
      attached = direct;
      direct.startDrag = start;
      direct.dragging = isDragging;
      direct.cancelDrag = cancelDrag;
      win.addEventListener('pointercancel', onPointerCancel, true);
      doc.addEventListener('pointercancel', onPointerCancel, true);
      win.addEventListener('blur', onBlur);
      doc.addEventListener('visibilitychange', onVisibilityChange);
      for (const type of ['mousedown', 'mouseup', 'auxclick', 'contextmenu']) win.addEventListener(type, onOtherButton, true);
    };
    host.addEventListener('vibeoffice-engine-runtime-ready', attach);
    attach();
    return () => {
      host.removeEventListener('vibeoffice-engine-runtime-ready', attach);
      if (attached?.startDrag === start) delete attached.startDrag;
      if (attached?.dragging === isDragging) delete attached.dragging;
      if (attached?.cancelDrag === cancelDrag) delete attached.cancelDrag;
      win.removeEventListener('pointercancel', onPointerCancel, true);
      doc.removeEventListener('pointercancel', onPointerCancel, true);
      win.removeEventListener('blur', onBlur);
      doc.removeEventListener('visibilitychange', onVisibilityChange);
      for (const type of ['mousedown', 'mouseup', 'auxclick', 'contextmenu']) win.removeEventListener(type, onOtherButton, true);
      cancelActive('api');
      cleanupRef.current?.();
      setCursor(null);
    };
  }, [host]);

  return null;
}
