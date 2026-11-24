/** Routes legacy tooltip-dismiss requests to the React-owned editor tooltip. */
export function installReactTooltipService(runtime: Window): () => void {
  const view = runtime as Window & {
    L?: { ui?: Record<string, unknown> };
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root) return () => {};
  const previous = ui.hideTip;
  const hideTip = () => {
    if (!root.isConnected) return;
    const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-hide-tooltip', { bubbles: true }));
  };
  ui.hideTip = hideTip;
  return () => {
    if (ui.hideTip !== hideTip) return;
    if (previous === undefined) delete ui.hideTip;
    else ui.hideTip = previous;
  };
}
