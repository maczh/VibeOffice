/** Routes legacy command-state refreshes to React-owned controls in direct-host editors. */
export function installReactUiRefreshService(runtime: Window): () => void {
  type Runtime = Window & {
    L?: {
      ui?: Record<string, any>;
      rafThrottle?: (callback: () => void) => () => void;
      bus?: { emit(event: string): unknown };
    };
    VO?: { reactDirectHost?: { root?: HTMLElement | null } };
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root) return () => {};

  const previous = ui.refresh;
  const refresh = view.L?.rafThrottle
    ? view.L.rafThrottle(() => view.L?.bus?.emit('ui-refresh'))
    : () => view.L?.bus?.emit('ui-refresh');
  ui.refresh = refresh;
  return () => {
    if (ui.refresh !== refresh) return;
    if (previous === undefined) delete ui.refresh;
    else ui.refresh = previous;
  };
}
