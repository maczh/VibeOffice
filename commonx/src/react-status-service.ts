/** Bridges engine status requests to the React-owned status-message portal. */
export function installReactStatusService(runtime: Window): void {
  const view = runtime as Window & {
    L?: { ui?: Record<string, unknown> };
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  if (!view.L) view.L = {};
  if (!view.L.ui) view.L.ui = {};
  view.L.ui.reactStatus = (text: unknown, duration?: number): boolean => {
    const host = view.VO?.reactDirectHost?.root;
    if (!host) return false;
    const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
    host.dispatchEvent(new CustomEventConstructor('vibeoffice-react-status', {
      detail: { text: String(text || ''), duration },
      bubbles: true,
    }));
    return true;
  };
}
