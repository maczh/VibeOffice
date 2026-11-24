/** Installs the synchronous checkbox/radio facade consumed by legacy engine features. */
export function installReactChoiceService(runtime: Window): void {
  type Runtime = Window & {
    L?: { ui?: Record<string, any>; h?: (tag: string, attributes?: Record<string, unknown>) => HTMLElement };
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root || !view.L?.h) return;
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  let checkboxId = 0;
  let radioId = 0;
  const createChoice = (type: 'checkbox' | 'radio', name: string | undefined, label: string, checked: boolean, onChange: ((value?: boolean) => void) | undefined, options?: Record<string, unknown>) => {
    const id = String(options?.id || (type === 'checkbox' ? `ck${++checkboxId}` : `rd${++radioId}`));
    const wrap = view.L!.h!('span', { class: 'ck vibeoffice-react-choice-slot', 'data-react-choice-id': id, style: 'display:contents' }) as HTMLSpanElement;
    Object.defineProperty(wrap, 'input', { configurable: true, get: () => wrap.querySelector('input') });
    const request = { id, type, label, checked: !!checked, ...(name ? { name } : {}), onChange };
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-choice', { detail: { target: wrap, options: request }, bubbles: true }));
    return wrap;
  };
  ui.check = (label: string, checked: boolean, onChange?: (checked: boolean) => void, options?: Record<string, unknown>) => createChoice('checkbox', undefined, label, checked, onChange ? (value) => { if (value != null) onChange(value); } : undefined, options);
  ui.radio = (name: string, label: string, checked: boolean, onChange?: () => void, options?: Record<string, unknown>) => createChoice('radio', name, label, checked, onChange, options);
}
