/** Installs the shared menu API bridge used by engines mounted in a React editor host. */
export function installReactMenuService(runtime: Window): void {
  type MenuItem = Record<string, any>;
  type MenuOptions = { depth?: number; cls?: string; selectFirst?: boolean; fromBar?: number };
  type MenuBarApi = {
    active: number | null;
    menus: Array<{ label: string; items: unknown[] }>;
    openByKey: (key: string) => boolean;
    openIndex: (index: number, selectFirst?: boolean) => void;
  };
  type Runtime = Window & {
    L?: { ui?: Record<string, any>; icons?: { get(id: string): string }; t?: (text: string) => string; esc?: (text: string) => string };
    VO?: { reactDirectHost?: { root?: HTMLElement; menuOpen?: boolean; menuBarApi?: MenuBarApi } };
  };
  const view = runtime as Runtime;
  const L = view.L;
  const ui = L?.ui;
  const direct = view.VO?.reactDirectHost;
  const root = direct?.root;
  if (!ui || !direct || !root) return;
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  const resolveItem = (source: unknown): MenuItem => {
    const item = typeof source === 'string'
      ? source === '-' || source === '|'
        ? { sep: true }
        : ui.cmds?.[source]
          ? { cmd: source, label: ui.cmds[source].menuLabel || ui.cmds[source].label, icon: ui.cmds[source].icon, key: ui.cmds[source].key, enabled: ui.cmds[source].enabled, checked: ui.cmds[source].checked, radio: ui.cmds[source].radio }
          : { label: source, disabled: true }
      : source as MenuItem;
    if (item?.cmd && ui.cmds?.[item.cmd]) {
      const command = ui.cmds[item.cmd];
      return Object.assign({ label: command.menuLabel || command.label, icon: command.icon, key: command.key, enabled: command.enabled, checked: command.checked, radio: command.radio }, item);
    }
    return item;
  };
  const prepareItems = (items: unknown[]): MenuItem[] => items.map(resolveItem).filter((item) => item && !item.hidden).map((item) => {
    const checked = !!(item.checked && item.checked());
    const rawLabel = item.label == null ? '' : String(item.label);
    const rawHtml = (L.esc?.(rawLabel) || rawLabel).replace(/&amp;&amp;/g, '\u0000').replace(/&amp;(\S)/g, '<u>$1</u>').replace(/\u0000/g, '&amp;');
    return {
      ...item,
      disabled: !!item.disabled || !!(item.enabled && !item.enabled()),
      checked,
      labelHtml: item.label == null ? '' : item.raw ? rawHtml : ui.ampHTML(rawLabel),
      iconHtml: item.html || (item.icon ? L.icons?.get(item.icon) : checked ? item.radio ? '<span class="m-radio">●</span>' : '<span class="m-check">✓</span>' : ''),
      keyText: item.raw ? item.key || '' : ui.keyText?.(item.key) || '',
      accessKey: item.raw ? (String(item.label || '').match(/&([^&\s])/)?.[1]?.toLowerCase() || null) : ui.accessKey?.(rawLabel) || null,
    };
  });
  const isRect = (value: any): boolean => value && typeof value.getBoundingClientRect === 'function';
  const point = (at: any) => {
    const rect = isRect(at) ? at.getBoundingClientRect() : at;
    if (rect && rect.bottom != null && rect.left != null) return { x: rect.left, y: rect.bottom };
    return { x: rect?.x || 2, y: rect?.y || 2 };
  };

  ui.closeMenus = (keepDepth = 0) => {
    const keep = keepDepth || 0;
    if (root.isConnected) root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-close-menus', { detail: { keepDepth: keep }, bubbles: true }));
    if (!keep && direct.menuBarApi) direct.menuBarApi.active = null;
  };
  ui.openMenu = (source: unknown[] | (() => unknown[]), at?: any, options?: MenuOptions) => {
    const opts = options || {};
    const depth = opts.depth || 0;
    ui.closeMenus(depth);
    const items = prepareItems((typeof source === 'function' ? source() : source) as unknown[]);
    const { x, y } = point(at);
    if (opts.fromBar != null && direct.menuBarApi) direct.menuBarApi.active = opts.fromBar;
    const target = at && typeof at.dispatchEvent === 'function' ? at : root;
    target.dispatchEvent(new CustomEventConstructor('vibeoffice-react-menu', { detail: {
      items, left: x || 2, top: y || 2, side: at?.side, cls: opts.cls,
      overDialog: !!ui.dialogOpen?.(), selectFirst: opts.selectFirst, depth, fromBar: opts.fromBar,
      anchor: isRect(at) ? at : undefined,
      resolveItems: prepareItems,
      execute: (item: MenuItem) => {
        if (item.run) item.run();
        else if (item.cmd) ui.exec(item.cmd, item.arg);
        if (!ui.dialogOpen?.() && (!document.activeElement || document.activeElement === document.body || !document.activeElement.isConnected)) ui.hooks?.refocus?.();
      },
    }, bubbles: true }));
    return undefined;
  };
  ui.contextMenu = (event: MouseEvent, items: unknown[]) => { event.preventDefault(); ui.openMenu(items, { x: event.clientX, y: event.clientY }); };
  ui.menuOpen = () => !!direct.menuOpen;
  ui.menuKey = () => !!direct.menuOpen;
  ui.menuBar = (container: HTMLElement, menus: Array<{ label: string; items: unknown[] }>): MenuBarApi => {
    const bar: MenuBarApi = {
      active: null,
      menus,
      openByKey(key) {
        if (ui.barKeys?.active?.()) return ui.barKeys.byKey(key);
        const index = menus.findIndex((menu) => ui.accessKey?.(menu.label) === key.toLowerCase());
        if (index < 0) return false;
        bar.openIndex(index, true);
        return true;
      },
      openIndex(index, selectFirst) {
        if (ui.barKeys?.active?.()) { ui.barKeys.index(index); return; }
        if (!menus[index]) return;
        if (direct.menuOpen && bar.active === index) { ui.closeMenus(); return; }
        const button = container.querySelector('.menubar')?.children.item(index) || container;
        ui.openMenu(menus[index].items, button, { cls: 'from-bar', selectFirst, fromBar: index });
      },
    };
    direct.menuBarApi = bar;
    ui.menus = menus;
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-menu-bar', { detail: {
      container,
      menus: menus.map((menu) => ({ labelHtml: ui.ampHTML(menu.label), items: menu.items })),
      openIndex: (index: number, selectFirst?: boolean) => bar.openIndex(index, selectFirst),
      close: () => ui.closeMenus(),
    }, bubbles: true }));
    return bar;
  };
}
