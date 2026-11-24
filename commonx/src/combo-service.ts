/** Installs the React-backed editable combo bridge for the legacy engine API. */
export function installReactComboService(runtime: Window): void {
  type Runtime = Window & {
    L?: { ui?: Record<string, any>; h?: (tag: string, attributes?: Record<string, unknown>) => HTMLElement };
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root || !view.L?.h) return;
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  ui.combo = (options: Record<string, any>) => {
    const control = view.L!.h!('div', {
      class: 'combo vibeoffice-react-combo',
      id: `combo-${options.id}`,
      style: `width:${options.width || 120}px`,
      'data-tip': options.tip || '',
    });
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-combo', {
      detail: { target: control, options }, bubbles: true,
    }));
    return control;
  };
}
