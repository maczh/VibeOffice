/** Installs the synchronous spinner facade backed by the React-owned shared spin layer. */
export function installReactSpinService(runtime: Window): void {
  type Runtime = Window & {
    L?: { ui?: Record<string, any>; h?: (tag: string, attributes?: Record<string, unknown>, ...children: unknown[]) => HTMLElement; round?: (value: number, digits?: number) => number };
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root || !view.L?.h) return;
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  let nextId = 0;
  ui.spin = (options: Record<string, any>) => {
    const id = options.id || `sp${++nextId}`;
    const format = (value: number) => options.fmt ? options.fmt(value) : (view.L?.round?.(value, options.dec == null ? 2 : options.dec) ?? value) + (options.unit || '');
    const wrap = view.L!.h!('span', { class: 'spin vibeoffice-react-spin', 'data-react-spin-id': id }) as HTMLSpanElement;
    const parse = () => {
      const input = wrap.querySelector('input');
      const value = parseFloat(String(input ? input.value : format(options.value)).replace(/[^0-9.\-]/g, ''));
      return Number.isNaN(value) ? options.value : value;
    };
    Object.defineProperties(wrap, {
      input: { configurable: true, get: () => wrap.querySelector('input') },
      get: { configurable: true, value: parse },
      set: { configurable: true, value: (value: number) => { options.value = value; const input = wrap.querySelector('input'); if (input) input.value = format(value); } },
      setDisabled: { configurable: true, value: (disabled: boolean) => {
        (wrap as HTMLSpanElement & { __reactSpinDisabled?: boolean }).__reactSpinDisabled = !!disabled;
        wrap.querySelectorAll<HTMLInputElement | HTMLButtonElement>('input,button').forEach((control) => { control.disabled = !!disabled; });
        wrap.dispatchEvent(new CustomEventConstructor('vibeoffice-react-spin-disabled'));
      } },
    });
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-spin', { detail: { target: wrap, options: Object.assign({}, options, { id }) }, bubbles: true }));
    return wrap;
  };
}
