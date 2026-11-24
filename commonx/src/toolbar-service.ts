/** Installs React-backed toolbar controls while retaining the classic DOM fallback. */
export function installReactToolbarService(runtime: Window): () => void {
  type Runtime = Window & {
    L?: { ui?: Record<string, any>; h?: (tag: string, attributes?: Record<string, unknown>) => HTMLElement; icons?: { get(name: string): string } };
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root || !view.L?.h) return () => {};
  const original = new Map<string, unknown>();
  const installed = new Map<string, unknown>();
  const replace = (name: string, value: unknown) => { original.set(name, ui[name]); installed.set(name, value); ui[name] = value; };
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  const stripAmp = (text: string) => ui.stripAmp(text);
  const classicTbButton = ui.tbButton as (id: string, options?: Record<string, any>) => HTMLElement;
  const classicTbSplit = ui.tbSplit as (id: string, menuFn: (anchor: { left: number; bottom: number }, target: HTMLElement) => void, options?: Record<string, any>) => HTMLElement;
  const classicTbDrop = ui.tbDrop as (label: string, icon: string, menuFn: (anchor: { left: number; bottom: number }, target: HTMLElement) => void, tip?: string, forceDOM?: boolean) => HTMLElement;
  replace('tbButton', (id: string, options: Record<string, any> = {}) => options.forceDOM
    ? classicTbButton(id, options)
    : ({
      __vibeToolbarControl: 'button', id, opts: Object.assign({}, options),
    }));
  replace('tbSplit', (id: string, menuFn: (anchor: { left: number; bottom: number }, target: HTMLElement) => void, options: Record<string, any> = {}) => options.forceDOM
    ? classicTbSplit(id, menuFn, options)
    : ({
      __vibeToolbarControl: 'split', id, opts: Object.assign({}, options), menuFn,
    }));
  replace('tbDrop', (label: string, icon: string, menuFn: (anchor: { left: number; bottom: number }, target: HTMLElement) => void, tip?: string, forceDOM?: boolean) => {
    if (forceDOM) return classicTbDrop(label, icon, menuFn, tip, true);
    const control: Record<string, any> = { __vibeToolbarControl: 'drop', label: label || '', icon: icon || '', tip: tip || stripAmp(label || ''), menuFn, dataAs: '' };
    control.setAttribute = (name: string, value: string) => { if (name === 'data-as') control.dataAs = String(value); };
    return control;
  });
  replace('toolbar', (id: string, title: string, items: unknown[]) => {
    const toolbar = view.L!.h!('div', { class: 'toolbar vibeoffice-react-toolbar', id: `tb-${id}`, role: 'toolbar', 'aria-label': title });
    ui.reactToolbar(id, title, toolbar, items);
    return toolbar;
  });
  replace('reactToolbar', (id: string, title: string, target: HTMLElement, items: unknown[], kind?: string) => {
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-toolbar', {
      detail: { id, title, target, items, kind }, bubbles: true,
    }));
    return true;
  });
  return () => {
    for (const [name, value] of original) {
      if (ui[name] === installed.get(name)) {
        if (value === undefined) delete ui[name];
        else ui[name] = value;
      }
    }
    original.clear();
    installed.clear();
  };
}
