import { useEffect } from 'react';

interface BrushOptions<Value = unknown> {
  el(): HTMLElement;
  pick(): Value | null | undefined;
  apply(value: Value, ...args: unknown[]): boolean;
  hint?: string;
  status?(text: string): void;
}

export interface EditorBrush<Value = unknown> {
  value: Value | null;
  sticky: boolean;
  on(): boolean;
  element(): HTMLElement;
  start(sticky?: boolean, value?: Value): boolean;
  stop(): void;
  click(event?: { detail?: number }): void;
  release(...args: unknown[]): boolean;
}

interface BrushUi {
  brush(options: BrushOptions): EditorBrush;
  brushOn?(): EditorBrush | null;
  brushReactOwned?: boolean;
  refresh?(): void;
  dialogOpen?(): boolean;
  menuOpen?(): boolean;
}

interface BrushRuntime extends Window {
  L?: { ui?: BrushUi };
  VO?: { reactDirectHost?: { root?: HTMLElement | null } };
}

/** Replace the direct-host Format Painter controller after the shared compatibility API loads. */
export function installReactBrushService(view: Window): () => void {
  const runtime = view as BrushRuntime;
  const ui = runtime.L?.ui;
  const root = runtime.VO?.reactDirectHost?.root;
  if (!ui || !root) return () => {};
  const previousBrush = ui.brush;
  const previousBrushOn = ui.brushOn;
  const previousBrushOwner = ui.brushReactOwned;
  let active: EditorBrush | null = null;
  const CustomEventConstructor = (view as Window & typeof globalThis).CustomEvent;
  const publish = (brush: EditorBrush | null) => root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-brush', { detail: { brush }, bubbles: true }));
  const brushOn = () => active;
  const brushFactory = (options: BrushOptions): EditorBrush => {
    const brush: EditorBrush = {
      value: null,
      sticky: false,
      on: () => active === brush,
      element: () => options.el(),
      start(sticky = false, value?: unknown) {
        const next = value !== undefined ? value : options.pick();
        if (next == null) return false;
        if (active && active !== brush) active.stop();
        active = brush;
        brush.value = next;
        brush.sticky = !!sticky;
        publish(brush);
        options.status?.(options.hint || '');
        ui.refresh?.();
        return true;
      },
      stop() {
        if (active !== brush) return;
        active = null;
        brush.value = null;
        brush.sticky = false;
        publish(null);
        options.status?.('');
        ui.refresh?.();
      },
      click(event) {
        if (event?.detail === 2) {
          if (brush.on()) { brush.sticky = true; options.status?.(options.hint || ''); }
          else brush.start(true);
        } else if (brush.on()) brush.stop();
        else brush.start(false);
      },
      release(...args) {
        if (!brush.on()) return false;
        const painted = !!options.apply(brush.value, ...args);
        if (painted && !brush.sticky) brush.stop();
        return painted;
      },
    };
    return brush;
  };
  ui.brush = brushFactory;
  ui.brushOn = brushOn;
  ui.brushReactOwned = true;
  return () => {
    active?.stop();
    if (ui.brush === brushFactory) ui.brush = previousBrush;
    if (ui.brushOn === brushOn) {
      if (previousBrushOn) ui.brushOn = previousBrushOn;
      else delete ui.brushOn;
    }
    if (previousBrushOwner === undefined) delete ui.brushReactOwned;
    else ui.brushReactOwned = previousBrushOwner;
  };
}

/** Owns Format Painter's direct-host visual state, Escape handling, and mount cleanup. */
export function EditorBrushLayer({ host }: { host: HTMLElement }) {
  useEffect(() => {
    const view = host.ownerDocument.defaultView as BrushRuntime | null;
    if (!view?.VO?.reactDirectHost) return;
    const ui = view.L?.ui;
    let active: EditorBrush | null = null;
    const onBrush = (event: Event) => {
      active?.element().classList.remove('painting');
      active = (event as CustomEvent<{ brush?: EditorBrush | null }>).detail?.brush ?? null;
      active?.element().classList.add('painting');
    };
    const onKeyDown = (event: Event) => {
      const key = event as KeyboardEvent;
      if (key.key !== 'Escape' || !active || ui?.dialogOpen?.() || ui?.menuOpen?.()) return;
      key.preventDefault();
      key.stopImmediatePropagation();
      active.stop();
    };
    host.addEventListener('vibeoffice-react-brush', onBrush);
    host.addEventListener('keydown', onKeyDown, true);
    return () => {
      host.removeEventListener('vibeoffice-react-brush', onBrush);
      host.removeEventListener('keydown', onKeyDown, true);
      active?.stop();
      active?.element().classList.remove('painting');
    };
  }, [host]);
  return null;
}
