/* Compatibility contract for app engine calls into the shared Ribbon UI. The view and KeyTips live in React. */
const installed = new WeakSet<Window>();

interface RibbonContext { id: string; set: string; when(): boolean }
interface RibbonSpec { tabs?: unknown[]; contextual?: RibbonContext[] }
interface RibbonPaneTask {
  current: string;
  el?: HTMLElement | null;
  registry: Record<string, unknown>;
  register(pane: { id: string; title: string; hidden: boolean; render(): void }): void;
  show(id: string): void;
}
interface RibbonFollowOptions {
  task?: RibbonPaneTask;
  sets?: Record<string, boolean | string>;
  on?(): boolean;
  close?(): void;
}
interface RibbonUi {
  select(id: string): void;
  collapse(collapsed?: boolean): void;
  built(): boolean;
  init(): void;
  follow(options: RibbonFollowOptions): void;
}
interface RibbonServiceNamespace {
  ribbonUI?: RibbonUi;
  ribbonSpec?: () => RibbonSpec;
  onLook?(callback: () => void): unknown;
  bus?: { on(event: string, callback: () => void): unknown };
}
interface RibbonServiceWindow extends Window {
  CustomEvent: typeof CustomEvent;
  L?: RibbonServiceNamespace;
  VO?: { reactDirectHost?: { root?: HTMLElement | null } };
  __VIBE_OFFICE_DIRECT_HOST__?: { root?: HTMLElement | null };
}

export function installReactRibbonService(runtime: Window): void {
  if (installed.has(runtime)) return;
  const window = runtime as RibbonServiceWindow;
  const L = window.L;
  if (!L) return;
  installed.add(runtime);
  const ribbonUI = L.ribbonUI || (L.ribbonUI = {} as RibbonUi);
  const root = () => window.VO?.reactDirectHost?.root || window.__VIBE_OFFICE_DIRECT_HOST__?.root || null;
  const dispatch = (name: string, detail: Record<string, unknown>) => {
    const host = root();
    if (host?.isConnected) host.dispatchEvent(new window.CustomEvent(name, { detail, bubbles: true }));
  };
  ribbonUI.select = (id: string) => dispatch('vibeoffice-react-ribbon-select', { id });
  ribbonUI.collapse = (collapsed?: boolean) => dispatch('vibeoffice-react-ribbon-collapse', { collapsed });
  ribbonUI.built = () => !!root()?.isConnected;
  ribbonUI.init = () => {
    const ribbonSpec = L.ribbonSpec;
    if (!ribbonSpec) return;
    const publish = () => {
      const spec = ribbonSpec();
      if (spec?.tabs?.length) dispatch('vibeoffice-react-ribbon', { spec });
    };
    publish();
    L.onLook?.(publish);
  };
  ribbonUI.follow = (options: RibbonFollowOptions) => {
    const task = options?.task;
    const sets = options?.sets || {};
    const ribbonSpec = L.ribbonSpec;
    const bus = L.bus;
    if (!task || !ribbonSpec || !bus) return;
    const contexts = ribbonSpec().contextual || [];
    const paneId = (context: RibbonContext): string => sets[context.id] === true ? `tools-${context.id}` : String(sets[context.id] || '');
    for (const context of contexts) {
      if (sets[context.id] === true) task.register({ id: paneId(context), title: context.set, hidden: true, render() {} });
    }
    const owned = new Set(contexts.filter((context) => sets[context.id]).map(paneId));
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
    bus.on('ui-refresh', () => {
      const context = (ribbonSpec().contextual || []).find((item) => sets[item.id] && item.when());
      const id = context ? paneId(context) : '';
      if (id !== shown) { shown = id; follow(id); }
      if (!id) before = { pane: task.current, hidden: !task.el || task.el.hidden };
    });
  };
}
