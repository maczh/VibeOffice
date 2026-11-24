/** Installs shared tab/field/group facades backed by React portals in the editor host. */
export function installReactStructureService(runtime: Window): void {
  type Tab = { label: string; body: HTMLElement; onShow?: () => void };
  type Runtime = Window & {
    L?: { ui?: Record<string, any>; h?: (tag: string, attributes?: Record<string, unknown>, ...children: unknown[]) => HTMLElement; t?: (text: string) => string };
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root || !view.L?.h) return;
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  const NodeConstructor = (runtime as Window & typeof globalThis).Node;
  let tabsId = 0;
  let structureId = 0;
  ui.tabs = (tabs: Tab[], active?: number) => {
    const id = `tabs${++tabsId}`;
    const head = view.L!.h!('div', { class: 'tabs-head vibeoffice-react-tabs-head', id: `${id}-head`, role: 'tablist' }) as HTMLDivElement & { __reactSelect?: (index: number) => void };
    const bodies = view.L!.h!('div', { class: 'tabs-body' });
    const initial = Math.max(0, Math.min(Number.isInteger(active) ? active! : 0, tabs.length - 1));
    const select = (index: number) => {
      if (!Number.isInteger(index) || index < 0 || index >= tabs.length) return;
      if (head.__reactSelect) head.__reactSelect(index);
      else tabs[index]?.onShow?.();
    };
    const element = view.L!.h!('div', { class: 'tabs', 'data-react-tabs-id': id }, head, bodies) as HTMLDivElement & { select?: (index: number) => void };
    element.select = select;
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-tabs', { detail: { target: head, bodyHost: bodies, tabs, active: initial, select }, bubbles: true }));
    select(initial);
    return element;
  };
  ui.field = (label: string, control: HTMLElement, options?: { cls?: string }) => {
    const controlProxy = control as HTMLElement & { input?: HTMLInputElement };
    const id = control.dataset.reactSelectId || control.dataset.reactSpinId || control.id || controlProxy.input?.id
      || control.querySelector<HTMLInputElement | HTMLSelectElement>('input,select')?.id;
    const target = view.L!.h!('div', { class: `field${options?.cls ? ` ${options.cls}` : ''}`, 'data-react-ui-id': `react-ui-field-${++structureId}` });
    target.append(control);
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-ui-structure', { detail: {
      kind: 'field', target, labelHtml: ui.ampHTML(label), labelFor: id || null, control,
    }, bubbles: true }));
    return target;
  };
  ui.group = (title: string, ...children: unknown[]) => {
    const target = view.L!.h!('fieldset', { class: 'grp', 'data-react-ui-id': `react-ui-group-${++structureId}` });
    const nodes: Node[] = [];
    const collect = (items: unknown[]) => items.forEach((child) => {
      if (child == null || child === false) return;
      if (Array.isArray(child)) { collect(child); return; }
      if (child instanceof NodeConstructor) nodes.push(child);
      else nodes.push(runtime.document.createTextNode(typeof child === 'string' ? (view.L?.t ? view.L.t(child) : child) : String(child)));
    });
    collect(children);
    target.append(...nodes);
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-ui-structure', { detail: {
      kind: 'group', target, legendHtml: ui.ampHTML(title), children: nodes,
    }, bubbles: true }));
    return target;
  };
}
