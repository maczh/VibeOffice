/* Compatibility contract for app engine calls into the shared Ribbon UI. The view and KeyTips live in React. */
// @ts-nocheck -- this adapter bridges app-specific legacy runtime interfaces at the migration boundary.

const installed = new WeakSet<Window>();

export function installReactRibbonService(runtime: Window): void {
  if (installed.has(runtime)) return;
  const window = runtime as Window & { L?: any; VO?: any; CustomEvent: typeof CustomEvent };
  const L = window.L;
  if (!L) return;
  installed.add(runtime);
  const ribbonUI = L.ribbonUI || (L.ribbonUI = {});
  const root = () => window.VO?.reactDirectHost?.root || window.__VIBE_OFFICE_DIRECT_HOST__?.root || null;
  const dispatch = (name: string, detail: Record<string, unknown>) => {
    const host = root();
    if (host?.isConnected) host.dispatchEvent(new window.CustomEvent(name, { detail, bubbles: true }));
  };
  ribbonUI.select = (id: string) => dispatch('vibeoffice-react-ribbon-select', { id });
  ribbonUI.collapse = (collapsed?: boolean) => dispatch('vibeoffice-react-ribbon-collapse', { collapsed });
  ribbonUI.built = () => !!root()?.isConnected;
  ribbonUI.init = () => {
    if (!L.ribbonSpec) return;
    const publish = () => {
      const spec = L.ribbonSpec();
      if (spec?.tabs?.length) dispatch('vibeoffice-react-ribbon', { spec });
    };
    publish();
    L.onLook?.(publish);
  };
  ribbonUI.follow = (options: any) => {
    const task = options?.task;
    const sets = options?.sets || {};
    if (!task || !L.ribbonSpec) return;
    const contexts = L.ribbonSpec().contextual || [];
    const paneId = (context: any) => sets[context.id] === true ? `tools-${context.id}` : sets[context.id];
    for (const context of contexts) {
      if (sets[context.id] === true) task.register({ id: paneId(context), title: context.set, hidden: true, render() {} });
    }
    const owned = new Set(contexts.filter((context: any) => sets[context.id]).map(paneId));
    let shown = '';
    let before: { pane: string; hidden: boolean } | null = null;
    const follow = (id: string) => {
      if ((options.on && !options.on()) || !task.el || task.el.hidden || window.innerWidth < 820) return;
      if (id) {
        if (task.current === id) return;
        if (!owned.has(task.current)) before = { pane: task.current, hidden: false };
        task.show(id);
      } else if (owned.has(task.current) && before) {
        if (before.hidden && options.close) options.close();
        else if (before.pane && task.registry[before.pane] && !owned.has(before.pane)) task.show(before.pane);
      }
    };
    L.bus.on('ui-refresh', () => {
      const context = (L.ribbonSpec().contextual || []).find((item: any) => sets[item.id] && item.when());
      const id = context ? paneId(context) : '';
      if (id !== shown) { shown = id; follow(id); }
      if (!id) before = { pane: task.current, hidden: !task.el || task.el.hidden };
    });
  };
}
