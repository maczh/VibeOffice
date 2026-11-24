/** Routes the legacy drag facade into the React-owned pointer lifecycle in direct-host editors. */
export function installReactDragService(runtime: Window): () => void {
  type DragData = { x0: number; y0: number; x: number; y: number; dx: number; dy: number; moved: boolean; data: Record<string, unknown> };
  type DragOptions = { move?(event: PointerEvent, data: DragData): void; up?(event: PointerEvent | null, data: DragData): void; cancel?(data: DragData, reason: string): void };
  type DirectHost = { startDrag?(event: PointerEvent, options: DragOptions, onEscape: () => void): DragData | null; dragging?(): boolean; cancelDrag?(): void };
  type Runtime = Window & { L?: { ui?: Record<string, any> }; VO?: { reactDirectHost?: DirectHost } };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  if (!ui) return () => {};
  const previous = { drag: ui.drag, dragging: ui.dragging, dragCancel: ui.dragCancel };
  const installed = {
    drag: (event: PointerEvent, options: DragOptions) => {
      if (event.button != null && event.button !== 0) return null;
      const direct = view.VO?.reactDirectHost;
      if (typeof direct?.startDrag === 'function') return direct.startDrag(event, options, () => ui.brushOn?.()?.stop());
      return previous.drag?.(event, options) ?? null;
    },
    dragging: () => {
      const direct = view.VO?.reactDirectHost;
      return typeof direct?.dragging === 'function' ? direct.dragging() : !!previous.dragging?.();
    },
    dragCancel: () => {
      const direct = view.VO?.reactDirectHost;
      if (typeof direct?.cancelDrag === 'function') direct.cancelDrag();
      else previous.dragCancel?.();
    },
  };
  Object.assign(ui, installed);
  return () => {
    for (const name of Object.keys(installed) as Array<keyof typeof installed>) {
      if (ui[name] !== installed[name]) continue;
      if (previous[name] === undefined) delete ui[name];
      else ui[name] = previous[name];
    }
  };
}
