import { useCallback, useEffect, useLayoutEffect, useRef, useState, type CSSProperties } from 'react';
import { createPortal } from 'react-dom';
import { attachLegacyDialog, type LegacyDialogMountRequest } from './legacy-dialog.js';

export interface ReactMenuItem {
  sep?: boolean;
  heading?: string;
  custom?: (close: () => void) => HTMLElement;
  label?: string;
  labelHtml?: string;
  labelStyle?: CSSProperties;
  iconHtml?: string;
  keyText?: string;
  accessKey?: string | null;
  raw?: boolean;
  disabled?: boolean;
  enabled?: () => boolean;
  checked?: boolean;
  radio?: boolean;
  bold?: boolean;
  tip?: string;
  sub?: unknown[];
  run?: () => unknown;
  cmd?: string;
  arg?: unknown;
}

export interface ReactMenuRequest {
  items: ReactMenuItem[];
  left: number;
  top: number;
  side?: number;
  cls?: string;
  overDialog?: boolean;
  selectFirst?: boolean;
  fromBar?: number;
  anchor?: Element;
  resolveItems: (items: unknown) => ReactMenuItem[];
  execute: (item: ReactMenuItem) => void;
}

interface MenuLayerProps {
  host: HTMLElement;
}

interface MenuFrame extends ReactMenuRequest {
  active: number;
}

export interface ReactMenuBarDefinition {
  labelHtml: string;
  items: unknown[];
}

export interface ReactMenuBarRequest {
  container: HTMLElement;
  menus: ReactMenuBarDefinition[];
  openIndex: (index: number, selectFirst?: boolean) => void;
  close: () => void;
}

function CustomMenuContent({ item, close }: { item: ReactMenuItem; close: () => void }) {
  const slot = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const content = item.custom?.(close);
    if (!content || !slot.current) return;
    slot.current.append(content);
    return () => content.remove();
  }, [item, close]);
  return <div ref={slot} />;
}

/** React-owned shared menu surfaces; command data and dynamic enabled state remain in the engine. */
export function MenuLayer({ host }: MenuLayerProps) {
  const [stack, setStack] = useState<MenuFrame[]>([]);
  const [menuBar, setMenuBar] = useState<ReactMenuBarRequest | null>(null);
  const [activeBar, setActiveBar] = useState<number | null>(null);
  const panelRefs = useRef<Array<HTMLDivElement | null>>([]);
  const activeHost = useCallback((active: boolean) => {
    const direct = (window as Window & { VO?: { reactDirectHost?: { menuOpen?: boolean } } }).VO?.reactDirectHost;
    if (direct) direct.menuOpen = active;
  }, []);
  const close = useCallback((keepDepth = 0) => {
    setStack((current) => current.slice(0, Math.max(0, keepDepth)));
    if (!keepDepth) setActiveBar(null);
    activeHost(keepDepth > 0);
  }, [activeHost]);

  useEffect(() => {
    const open = (event: Event) => {
      const request = (event as CustomEvent<ReactMenuRequest & { depth?: number }>).detail;
      if (!request || !Array.isArray(request.items)) return;
      const frame: MenuFrame = { ...request, active: request.selectFirst ? request.items.findIndex((item) => !item.disabled && !item.sep && !item.heading && !item.custom) : -1 };
      setStack((current) => request.depth ? [...current.slice(0, request.depth), frame] : [frame]);
      if (!request.depth) setActiveBar(request.fromBar ?? null);
      activeHost(true);
    };
    const closeEvent = (event: Event) => {
      const keepDepth = (event as CustomEvent<{ keepDepth?: number }>).detail?.keepDepth ?? 0;
      close(keepDepth);
    };
    const barEvent = (event: Event) => {
      const request = (event as CustomEvent<ReactMenuBarRequest>).detail;
      if (request?.container && Array.isArray(request.menus)) setMenuBar(request);
    };
    const dialogEvent = (event: Event) => {
      const request = (event as CustomEvent<LegacyDialogMountRequest>).detail;
      if (!request || request.host !== host || request.api) return;
      const engineRoot = host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root]');
      if (!engineRoot) return;
      try { attachLegacyDialog(engineRoot, request); }
      catch (error) { console.error('VibeOffice React dialog failed to mount:', error); }
    };
    host.addEventListener('vibeoffice-react-menu', open);
    host.addEventListener('vibeoffice-react-close-menus', closeEvent);
    host.addEventListener('vibeoffice-react-menu-bar', barEvent);
    host.addEventListener('vibeoffice-react-dialog', dialogEvent);
    return () => {
      host.removeEventListener('vibeoffice-react-menu', open);
      host.removeEventListener('vibeoffice-react-close-menus', closeEvent);
      host.removeEventListener('vibeoffice-react-menu-bar', barEvent);
      host.removeEventListener('vibeoffice-react-dialog', dialogEvent);
      setMenuBar(null);
      activeHost(false);
    };
  }, [activeHost, close, host]);

  useEffect(() => {
    if (!stack.length) return;
    const outside = (event: PointerEvent) => {
      const target = event.target;
      if (!(target instanceof Node)) return;
      if (panelRefs.current.some((panel) => panel?.contains(target)) || stack.some((menu) => menu.anchor?.contains(target))) return;
      close();
    };
    const keydown = (event: KeyboardEvent) => {
      const depth = stack.length - 1;
      const menu = stack[depth];
      if (!menu) return;
      const selectable = menu.items.map((item, index) => ({ item, index })).filter(({ item }) => !item.sep && !item.heading && !item.custom && !item.disabled);
      const move = (direction: number) => {
        if (!selectable.length) return;
        const index = selectable.findIndex(({ index }) => index === menu.active);
        const nextIndex = index < 0 ? (direction > 0 ? 0 : selectable.length - 1) : (index + direction + selectable.length) % selectable.length;
        const next = selectable[nextIndex];
        if (next) setStack((current) => current.map((frame, level) => level === depth ? { ...frame, active: next.index } : frame));
      };
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(depth); }
      else if (event.key === 'ArrowDown') { event.preventDefault(); event.stopPropagation(); move(1); }
      else if (event.key === 'ArrowUp') { event.preventDefault(); event.stopPropagation(); move(-1); }
      else if (event.key === 'ArrowRight') {
        const item = menu.items[menu.active];
        if (item?.sub) { event.preventDefault(); event.stopPropagation(); openSub(depth, menu.active); }
      } else if (event.key === 'ArrowLeft' && depth > 0) { event.preventDefault(); event.stopPropagation(); close(depth); }
      else if ((event.key === 'ArrowLeft' || event.key === 'ArrowRight') && depth === 0 && menu.fromBar != null && menuBar) {
        event.preventDefault(); event.stopPropagation();
        menuBar.openIndex((menu.fromBar + (event.key === 'ArrowRight' ? 1 : -1) + menuBar.menus.length) % menuBar.menus.length, true);
      }
      else if (event.key === 'Enter' || event.key === ' ') {
        const item = menu.items[menu.active];
        if (item) { event.preventDefault(); event.stopPropagation(); activate(depth, menu.active); }
      } else if (event.key.length === 1) {
        const needle = event.key.toLowerCase();
        const hit = selectable.find(({ item }) => item.accessKey === needle);
        if (hit) { event.preventDefault(); event.stopPropagation(); activate(depth, hit.index); }
        else { event.preventDefault(); event.stopPropagation(); }
      } else { event.preventDefault(); event.stopPropagation(); }
    };
    document.addEventListener('pointerdown', outside, true);
    document.addEventListener('keydown', keydown, true);
    return () => {
      document.removeEventListener('pointerdown', outside, true);
      document.removeEventListener('keydown', keydown, true);
    };
  }, [stack, close, menuBar]);

  const openSub = (depth: number, index: number) => {
    const menu = stack[depth], row = panelRefs.current[depth]?.querySelectorAll<HTMLElement>('.m-item')[index], item = menu?.items[index];
    if (!menu || !row || !item?.sub) return;
    const rect = row.getBoundingClientRect();
    const submenuItems = menu.resolveItems(item.sub);
    setStack((current) => [...current.slice(0, depth + 1), {
      ...menu, items: submenuItems, left: rect.right - 3, top: rect.top - 3, side: rect.left, anchor: row, active: -1,
    }]);
  };
  const activate = (depth: number, index: number) => {
    const menu = stack[depth], item = menu?.items[index];
    if (!menu || !item || item.disabled || (item.enabled && !item.enabled()) || item.sep || item.heading || item.custom) return;
    if (item.sub) { openSub(depth, index); return; }
    close();
    menu.execute(item);
  };

  useLayoutEffect(() => {
    stack.forEach((menu, depth) => {
      const panel = panelRefs.current[depth];
      if (!panel) return;
      const rect = panel.getBoundingClientRect();
      const left = Math.max(2, menu.left + rect.width > window.innerWidth - 2 ? menu.side != null ? menu.side - rect.width + 3 : window.innerWidth - rect.width - 2 : menu.left);
      const top = Math.max(2, menu.top + rect.height > window.innerHeight - 2 ? window.innerHeight - rect.height - 2 : menu.top);
      panel.style.left = `${left}px`;
      panel.style.top = `${top}px`;
      if (menu.active >= 0) panel.querySelectorAll<HTMLElement>('.m-item')[menu.active]?.classList.add('act');
    });
    if (stack.length) panelRefs.current.at(-1)?.focus({ preventScroll: true });
  }, [stack]);

  const engineRoot = host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root]');
  return <>{menuBar && createPortal(<div className="menubar" role="menubar">{menuBar.menus.map((menu, index) => <div
    key={`menubar:${index}`}
    className={`mb-item${activeBar === index ? ' open' : ''}`}
    role="menuitem"
    tabIndex={-1}
    dangerouslySetInnerHTML={{ __html: menu.labelHtml }}
    onPointerDown={(event) => { event.preventDefault(); activeBar === index ? menuBar.close() : menuBar.openIndex(index); }}
    onPointerEnter={() => { if (activeBar != null && activeBar !== index) menuBar.openIndex(index, true); }}
  />)}</div>, menuBar.container)}{engineRoot && createPortal(stack.map((menu, depth) => <div
    key={`menu:${depth}`}
    ref={(element) => { panelRefs.current[depth] = element; }}
    className={`menu vibeoffice-menu-layer${menu.cls ? ` ${menu.cls}` : ''}${menu.overDialog ? ' over-dlg' : ''}`}
    role="menu"
    tabIndex={-1}
    style={{ left: menu.left, top: menu.top }}
    onPointerLeave={() => {}}
  >{menu.items.map((item, index) => {
    if (item.sep) return <div key={`sep:${index}`} className="m-sep" />;
    if (item.heading) return <div key={`heading:${index}`} className="m-head">{item.heading}</div>;
    if (item.custom) return <CustomMenuContent key={`custom:${index}`} item={item} close={() => close()} />;
    const disabled = !!item.disabled || !!(item.enabled && !item.enabled());
    const checked = !!item.checked;
    return <div key={`${item.cmd || item.label || index}:${index}`} className={`m-item${disabled ? ' dis' : ''}${checked ? ' chk' : ''}${item.sub ? ' has-sub' : ''}${item.bold ? ' bold' : ''}${menu.active === index ? ' act' : ''}`} role="menuitem" aria-disabled={disabled ? 'true' : undefined} title={item.tip || undefined} onPointerEnter={() => {
      setStack((current) => current.map((frame, level) => level === depth ? { ...frame, active: index } : frame));
      if (item.sub && !disabled) openSub(depth, index);
      else close(depth + 1);
    }} onClick={(event) => { event.stopPropagation(); activate(depth, index); }}>
      <span className={`m-ico${checked && item.iconHtml ? ' on' : ''}`} dangerouslySetInnerHTML={{ __html: item.iconHtml || (checked ? item.radio ? '<span class="m-radio">●</span>' : '<span class="m-check">✓</span>' : '') }} />
      <span className="m-lbl" translate={item.raw ? 'no' : undefined} style={item.labelStyle} dangerouslySetInnerHTML={{ __html: item.labelHtml || '' }} />
      <span className="m-key" translate={item.raw ? 'no' : undefined}>{item.keyText}</span>
      <span className="m-sub">{item.sub ? '▶' : ''}</span>
    </div>;
  })}</div>), engineRoot)}</>;
}
