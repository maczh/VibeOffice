/** Leaves editor-mount ownership to the React host instead of the classic shared UI toolkit. */
export function installReactUiMountService(runtime: Window): () => void {
  type Ui = { mount?: (host: HTMLElement) => (() => void) | void; unmount?: (host: HTMLElement) => void };
  type Runtime = Window & { L?: { ui?: Ui }; VO?: { reactDirectHost?: { root?: HTMLElement } } };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root) return () => {};

  const previousMount = ui.mount;
  const previousUnmount = ui.unmount;
  const mount = (_host: HTMLElement) => () => {};
  const unmount = (_host: HTMLElement) => {};
  ui.mount = mount;
  ui.unmount = unmount;

  return () => {
    if (ui.mount === mount) {
      if (previousMount === undefined) delete ui.mount;
      else ui.mount = previousMount;
    }
    if (ui.unmount === unmount) {
      if (previousUnmount === undefined) delete ui.unmount;
      else ui.unmount = previousUnmount;
    }
  };
}
