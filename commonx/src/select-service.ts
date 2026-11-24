/** Installs the DOM-compatible select facade backed by the shared React select layer. */
export function installReactSelectService(runtime: Window): void {
  type Runtime = Window & {
    L?: { ui?: Record<string, any>; h?: (tag: string, attributes?: Record<string, unknown>) => HTMLElement };
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root || !view.L?.h) return;
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  let nextId = 0;
  ui.select = (options: unknown[], value: unknown, onChange?: (value: string) => void, settings?: Record<string, unknown>) => {
    const id = String(settings?.id || `sel${++nextId}`);
    const control = view.L!.h!('div', { class: 'vibeoffice-react-select-slot', 'data-react-select-id': id, style: 'display:contents' });
    Object.defineProperties(control, {
      value: { configurable: true, get: () => control.querySelector('select')?.value ?? String(value ?? ''), set: (next: unknown) => { const select = control.querySelector('select'); if (select) select.value = String(next); } },
      disabled: { configurable: true, get: () => control.querySelector('select')?.disabled ?? !!settings?.disabled, set: (next: unknown) => { const select = control.querySelector('select'); if (select) select.disabled = !!next; } },
      selectedIndex: { configurable: true, get: () => control.querySelector('select')?.selectedIndex ?? 0, set: (next: number) => { const select = control.querySelector('select'); if (select) select.selectedIndex = next; } },
    });
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-select', {
      detail: { target: control, options: Object.assign({}, settings || {}, { id, options, value, onChange }) }, bubbles: true,
    }));
    return control;
  };
}
