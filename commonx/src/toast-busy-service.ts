interface ToastAction { label: string; run: () => void }

/** Installs the React-host bridge for transient notices and progress overlays. */
export function installToastBusyService(runtime: Window): void {
  const view = runtime as Window & {
    L?: { ui?: Record<string, any>; t?: (text: string) => string };
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  if (!view.L) view.L = {};
  if (!view.L.ui) view.L.ui = {};
  let busy = false;
  const host = () => view.VO?.reactDirectHost?.root;
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;

  view.L.ui.toast = (text: string, duration?: number, action?: ToastAction): void => {
    const root = host();
    if (!root?.isConnected) return;
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-toast', {
      detail: {
        text: view.L?.t ? view.L.t(text) : text,
        duration: duration || 2600,
        ...(action ? { action: { label: action.label, run: action.run } } : {}),
      },
      bubbles: true,
    }));
  };

  view.L.ui.busy = (on: boolean, text?: string): void => {
    const root = host();
    if (!root?.isConnected) return;
    const next = !!on;
    if (busy !== next) {
      runtime.dispatchEvent(new CustomEventConstructor('vo-busy', { detail: { on: next, text: next ? text || 'Working...' : '' } }));
      busy = next;
    }
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-busy', {
      detail: next ? { text: view.L?.t ? view.L.t(text || 'Working...') : text || 'Working...' } : null,
      bubbles: true,
    }));
  };
}
