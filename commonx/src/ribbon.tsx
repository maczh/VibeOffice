import { createElement, useCallback, useEffect, useMemo, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

interface LegacyCommand {
  label?: string;
  tip?: string;
  tbLabel?: string;
  key?: string;
  icon?: string;
  enabled?: () => boolean;
  checked?: () => boolean;
}

interface LegacyRibbonItem {
  cmd?: string;
  arg?: unknown;
  size?: 'large' | 'medium' | 'small';
  label?: string;
  tip?: string;
  icon?: string;
  split?: string;
  drop?: string;
  menu?: (...args: unknown[]) => unknown;
  make?: () => HTMLElement;
  row?: Array<LegacyRibbonItem | string>;
}

interface LegacyRibbonGroup {
  label: string;
  icon?: string;
  launcher?: string;
  items: Array<LegacyRibbonItem | string>;
}

interface LegacyRibbonTab { id: string; label: string; groups: LegacyRibbonGroup[] }
interface LegacyRibbonContext { id: string; set: string; when: () => boolean; open?: boolean; tabs: LegacyRibbonTab[] }
interface LegacyRibbonSpec {
  qat?: Array<LegacyRibbonItem | string>;
  file?: () => unknown[];
  icons?: Record<string, string>;
  tabs: LegacyRibbonTab[];
  contextual?: LegacyRibbonContext[];
}

interface LegacyRibbonUi {
  cmds: Record<string, LegacyCommand>;
  exec(id: string, arg?: unknown, event?: Event): unknown;
  openMenu(items: unknown[], rect: DOMRect | { left: number; bottom: number }): unknown;
  stripAmp(text: string): string;
  ampHTML(text: string): string;
}

interface LegacyRibbonRuntime {
  ui: LegacyRibbonUi;
  icons: { get(name: string, size?: number): string };
  t(text: string): string;
  store: { get(key: string, fallback: unknown): unknown; set(key: string, value: unknown): void };
  bus: { on(event: string, callback: () => void): () => void };
  ribbonUI: { select(id: string): void; collapse(on?: boolean): void };
  ribbonSpec?: () => LegacyRibbonSpec;
}

function runtime(): LegacyRibbonRuntime | undefined { return (window as Window & { L?: LegacyRibbonRuntime }).L; }
function enabled(command: LegacyCommand | undefined): boolean { try { return !command?.enabled || !!command.enabled(); } catch { return false; } }
function checked(command: LegacyCommand | undefined): boolean { try { return !!command?.checked?.(); } catch { return false; } }

function LegacyControl({ item, size, spec }: { item: LegacyRibbonItem; size: LegacyRibbonItem['size']; spec: LegacyRibbonSpec }) {
  const host = useMemo(() => document.createElement('span'), []);
  useEffect(() => {
    const office = runtime();
    if (!office) return;
    const value = typeof item === 'string' ? { cmd: item } : item;
    const id = value.cmd;
    const icon = value.icon || (id ? spec.icons?.[id] : undefined);
    const label = size === 'small' ? '' : value.label ?? (id ? office.ui.stripAmp(office.ui.cmds[id]?.label || id).replace(/(\.\.\.|…)$/, '') : '');
    let element: HTMLElement | null = null;
    if (value.make) element = value.make();
    if (!element) return;
    element.classList.add(`rb-${size || 'medium'}`);
    if (size === 'large' && icon) {
      const svg = element.querySelector<SVGElement>('svg.ico');
      if (svg) svg.outerHTML = office.icons.get(icon, 32);
    }
    host.replaceChildren(element);
    return () => { element?.remove(); host.replaceChildren(); };
  }, [host, item, size, spec]);
  return <span className="vibeoffice-ribbon-control" style={{ display: 'contents' }} ref={(element) => { if (element && host.parentElement !== element) element.append(host); }} />;
}

/** Split and drop actions are React buttons; the legacy callback only supplies the existing menu model. */
function RibbonMenuControl({ item, size, spec }: { item: LegacyRibbonItem; size: LegacyRibbonItem['size']; spec: LegacyRibbonSpec }) {
  const office = runtime();
  if (!office) return null;
  const split = !!item.split;
  const id = item.split;
  const command = id ? office.ui.cmds[id] : undefined;
  const actualSize = item.size || size || 'medium';
  const label = actualSize === 'small' ? '' : item.label ?? (id ? office.ui.stripAmp(command?.label || id).replace(/(\.\.\.|…)$/, '') : item.drop || '');
  const icon = item.icon || (id ? spec.icons?.[id] || command?.icon : undefined);
  const tip = split
    ? office.ui.stripAmp(command?.tip || command?.label || id || '') + (command?.key ? ` (${command.key})` : '')
    : office.ui.stripAmp(item.tip || item.drop || label);
  const iconHtml = icon ? office.icons.get(icon, actualSize === 'large' ? 32 : undefined) : '';
  const labelHtml = label ? `<span class="tb-lbl">${office.ui.ampHTML(label)}</span>` : '';
  const openMenu = (anchor: HTMLElement) => {
    if (!item.menu) return;
    const rect = anchor.getBoundingClientRect();
    item.menu({ left: rect.left, bottom: rect.bottom }, anchor);
  };

  if (split && id) return <div className={`tb-split rb-${actualSize}`}>
    <button type="button" className={`tb-btn${label ? ' with-label' : ''}`} data-cmd={id} data-tip={tip} aria-label={office.ui.stripAmp(command?.label || label || id)} aria-pressed={command?.checked ? checked(command) : undefined} disabled={!enabled(command)} dangerouslySetInnerHTML={{ __html: iconHtml + labelHtml }} onPointerDown={(event) => event.preventDefault()} onClick={(event) => office.ui.exec(id, item.arg, event.nativeEvent)} />
    <button type="button" className="tb-dd" aria-label={`More ${office.ui.stripAmp(command?.label || id)} options`} aria-haspopup="menu" data-tip={`More ${office.ui.stripAmp(command?.label || id)} options`} onPointerDown={(event) => event.preventDefault()} onClick={(event) => openMenu(event.currentTarget.parentElement as HTMLElement)}><span className="dd-arrow" /></button>
  </div>;

  const dropLabel = actualSize === 'small' ? '' : item.label ?? item.drop ?? '';
  const dropTip = office.ui.stripAmp(item.tip || item.drop || dropLabel);
  return <button type="button" className={`tb-btn tb-drop rb-${actualSize}${dropLabel ? ' with-label' : ''}`} data-tip={dropTip} aria-label={dropTip} aria-haspopup="menu" dangerouslySetInnerHTML={{ __html: iconHtml + (dropLabel ? `<span class="tb-lbl">${office.ui.ampHTML(dropLabel)}</span>` : '') + '<span class="dd-arrow"></span>' }} onPointerDown={(event) => event.preventDefault()} onClick={(event) => openMenu(event.currentTarget)} />;
}

function RibbonItem({ value, size = 'medium', spec, refresh }: { value: LegacyRibbonItem | string; size?: LegacyRibbonItem['size']; spec: LegacyRibbonSpec; refresh: number }): ReactNode {
  const office = runtime();
  if (!office) return null;
  const item: LegacyRibbonItem = typeof value === 'string' ? { cmd: value } : value;
  if (item.row) return <div className="rb-row">{item.row.map((child, index) => child === '|' ? <span key={`gap:${index}`} className="rb-gap" /> : <RibbonItem key={`item:${index}`} value={child} size="small" spec={spec} refresh={refresh} />)}</div>;
  if (item.make) return <LegacyControl item={item} size={item.size || size} spec={spec} />;
  if (item.split || item.drop) return <RibbonMenuControl item={item} size={item.size || size} spec={spec} />;
  const id = item.cmd;
  if (!id) return null;
  const command = office.ui.cmds[id];
  const icon = item.icon || spec.icons?.[id] || command?.icon;
  const actualSize = item.size || size;
  const label = actualSize === 'small' ? '' : item.label ?? office.ui.stripAmp(command?.label || id).replace(/(\.\.\.|…)$/, '');
  const title = office.ui.stripAmp(command?.tip || command?.label || id) + (command?.key ? ` (${command.key})` : '');
  return <button
    type="button"
    className={`tb-btn rb-${actualSize}${label ? ' with-label' : ''}${checked(command) ? ' on' : ''}${enabled(command) ? '' : ' dis'}`}
    data-cmd={id}
    data-tip={title}
    aria-label={office.ui.stripAmp(command?.label || label || id)}
    aria-pressed={command?.checked ? checked(command) : undefined}
    disabled={!enabled(command)}
    onPointerDown={(event) => event.preventDefault()}
    onClick={(event) => office.ui.exec(id, item.arg, event.nativeEvent)}
    key={`${id}:${refresh}`}
  >{icon && <span className="tb-icon" aria-hidden="true" dangerouslySetInnerHTML={{ __html: office.icons.get(icon, actualSize === 'large' ? 32 : undefined) }} />}{label && <span className="tb-lbl">{office.t(label)}</span>}</button>;
}

function RibbonGroup({ group, spec, refresh }: { group: LegacyRibbonGroup; spec: LegacyRibbonSpec; refresh: number }) {
  const office = runtime();
  if (!office) return null;
  const columns: Array<{ kind: 'rows' | 'stack'; items: Array<LegacyRibbonItem | string> }> = [];
  const large: Array<LegacyRibbonItem | string> = [];
  let column: { kind: 'rows' | 'stack'; items: Array<LegacyRibbonItem | string> } | null = null;
  for (const value of group.items) {
    if (value === '|') { column = null; continue; }
    const item: LegacyRibbonItem = typeof value === 'string' ? { cmd: value } : value;
    if (item.size === 'large') { large.push(value); column = null; continue; }
    const kind = item.row ? 'rows' : 'stack';
    if (!column || column.kind !== kind || column.items.length >= (kind === 'rows' ? 2 : 3)) {
      column = { kind, items: [] };
      columns.push(column);
    }
    column.items.push(value);
  }
  const first = group.items.flatMap((value) => typeof value === 'object' && value.row ? value.row : [value]).find((value) => value !== '|' && !(typeof value === 'object' && value.make));
  const firstItem = typeof first === 'string' ? { cmd: first } : first;
  const firstId = firstItem?.cmd || firstItem?.split;
  const icon = group.icon || (firstId ? spec.icons?.[firstId] || office.ui.cmds[firstId]?.icon : undefined);
  return <div className="rb-group" role="group" aria-label={group.label}>
    <button type="button" className="tb-btn rb-large rb-fold tb-drop with-label" aria-label={group.label} data-tip={group.label}>
      {icon && <span className="tb-icon" aria-hidden="true" dangerouslySetInnerHTML={{ __html: office.icons.get(icon, 32) }} />}<span className="tb-lbl">{office.t(group.label)}</span><span className="dd-arrow" />
    </button>
    <div className="rb-body">
      {large.map((item, index) => <RibbonItem key={`large:${index}`} value={item} size="large" spec={spec} refresh={refresh} />)}
      {columns.map((column, colIndex) => <div key={`column:${colIndex}`} className="rb-col" data-kind={column.kind}>{column.items.map((item, index) => <RibbonItem key={`${colIndex}:${index}`} value={item} spec={spec} refresh={refresh} />)}</div>)}
    </div>
    <div className="rb-label">{office.t(group.label)}{group.launcher && <button type="button" className="rb-launch" aria-label={office.ui.stripAmp(office.ui.cmds[group.launcher]?.label || group.launcher)} onPointerDown={(event) => event.preventDefault()} onClick={(event) => office.ui.exec(group.launcher!, undefined, event.nativeEvent)}>↗</button>}</div>
  </div>;
}

function RibbonTab({ tab, selected, select, spec, refresh }: { tab: LegacyRibbonTab; selected: boolean; select: () => void; spec: LegacyRibbonSpec; refresh: number }) {
  const office = runtime();
  if (!office) return null;
  return <button type="button" role="tab" aria-selected={selected} className={`rb-tab${selected ? ' on' : ''}`} dangerouslySetInnerHTML={{ __html: office.ui.ampHTML(tab.label) }} onPointerDown={(event) => event.preventDefault()} onDoubleClick={() => runtime()?.ribbonUI.collapse()} onClick={select} />;
}

/** React-owned Ribbon structure, tabs, command buttons and live enabled/checked state. */
export function EditorRibbon({ host }: { host: HTMLElement }) {
  const [spec, setSpec] = useState<LegacyRibbonSpec | null>(null);
  const [refresh, setRefresh] = useState(0);
  const [selected, setSelected] = useState('home');
  const [collapsed, setCollapsed] = useState(false);
  const [peek, setPeek] = useState(false);
  const engineRoot = host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root]');
  const target = engineRoot?.querySelector<HTMLElement>('.cmdbars');

  const select = useCallback((id: string) => setSelected(id), []);
  const toggleCollapse = useCallback((value?: boolean) => {
    const next = value == null ? !collapsed : !!value;
    setCollapsed(next);
    setPeek(false);
    runtime()?.store.set('ribbonCollapsed', next);
  }, [collapsed]);

  useEffect(() => {
    const onSpec = (event: Event) => {
      const detail = (event as CustomEvent<{ spec?: LegacyRibbonSpec }>).detail;
      if (detail?.spec?.tabs?.length) setSpec(detail.spec);
    };
    const onSelect = (event: Event) => {
      const id = (event as CustomEvent<{ id?: string }>).detail?.id;
      if (id) select(id);
    };
    const onCollapse = (event: Event) => toggleCollapse((event as CustomEvent<{ collapsed?: boolean }>).detail?.collapsed);
    const onLook = () => { const next = runtime()?.ribbonSpec?.(); if (next) setSpec(next); };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'F1' && event.ctrlKey && !event.altKey && !event.shiftKey) { event.preventDefault(); event.stopPropagation(); toggleCollapse(); }
      else if (event.key === 'Escape' && peek) setPeek(false);
    };
    const onPointerDown = (event: PointerEvent) => { if (peek && target && !target.contains(event.target as Node)) setPeek(false); };
    host.addEventListener('vibeoffice-react-ribbon', onSpec);
    host.addEventListener('vibeoffice-react-ribbon-select', onSelect);
    host.addEventListener('vibeoffice-react-ribbon-collapse', onCollapse);
    window.addEventListener('vo-look', onLook);
    document.addEventListener('keydown', onKeyDown, true);
    document.addEventListener('pointerdown', onPointerDown, true);
    const offRefresh = runtime()?.bus.on('ui-refresh', () => setRefresh((current) => current + 1));
    setCollapsed(!!runtime()?.store.get('ribbonCollapsed', false));
    return () => { host.removeEventListener('vibeoffice-react-ribbon', onSpec); host.removeEventListener('vibeoffice-react-ribbon-select', onSelect); host.removeEventListener('vibeoffice-react-ribbon-collapse', onCollapse); window.removeEventListener('vo-look', onLook); document.removeEventListener('keydown', onKeyDown, true); document.removeEventListener('pointerdown', onPointerDown, true); offRefresh?.(); };
  }, [host, peek, select, target, toggleCollapse]);

  const activeContexts = useMemo(() => (spec?.contextual || []).filter((context) => { try { return context.when(); } catch { return false; } }), [refresh, spec]);
  if (!target || !spec) return null;
  const office = runtime();
  if (!office) return null;
  const tabList = [...spec.tabs, ...activeContexts.flatMap((context) => context.tabs)];
  const activeTab = tabList.find((tab) => tab.id === selected) || spec.tabs[0];
  if (!activeTab) return null;
  const openFile = (event: React.MouseEvent<HTMLButtonElement>) => {
    const rect = event.currentTarget.getBoundingClientRect();
    office.ui.openMenu(spec.file?.() || [], { left: rect.left, bottom: rect.bottom });
  };
  return createPortal(<div className={`ribbon vibeoffice-react-ribbon${collapsed ? ' collapsed' : ''}${peek ? ' peek' : ''}`}>
    <div className="rb-bar">
      <div className="rb-qat" role="toolbar" aria-label="Quick Access Toolbar">{(spec.qat || []).map((item, index) => <RibbonItem key={`qat:${index}`} value={item} size="small" spec={spec} refresh={refresh} />)}</div>
      <div className="rb-tabs" role="tablist">
        <button type="button" role="tab" className="rb-tab file" onPointerDown={(event) => event.preventDefault()} onClick={openFile}>{office.t('File')}</button>
        {spec.tabs.map((tab) => <RibbonTab key={tab.id} tab={tab} selected={activeTab.id === tab.id} select={() => { select(tab.id); if (collapsed) setPeek(true); }} spec={spec} refresh={refresh} />)}
        {activeContexts.map((context) => <div key={context.id} className="rb-ctx"><span className="rb-ctx-name">{office.t(context.set)}</span><div className="rb-ctx-tabs">{context.tabs.map((tab) => <RibbonTab key={tab.id} tab={tab} selected={activeTab.id === tab.id} select={() => select(tab.id)} spec={spec} refresh={refresh} />)}</div></div>)}
      </div>
      <button type="button" className="rb-pin" aria-label={office.t('Collapse the Ribbon (Ctrl+F1)')} data-tip="Collapse the Ribbon (Ctrl+F1)" onClick={() => toggleCollapse()}><svg width="10" height="6" viewBox="0 0 10 6" aria-hidden="true"><path d={collapsed ? 'M1 1l4 4 4-4' : 'M1 5l4-4 4 4'} fill="none" stroke="currentColor" strokeWidth="1.4" /></svg></button>
    </div>
    <div className="rb-panelwrap">{collapsed && !peek ? null : <div className="rb-panel" onClick={(event) => { if (peek && (event.target as HTMLElement).closest('.tb-btn:not(.tb-drop)')) setTimeout(() => setPeek(false), 0); }}><div className="rb-groups" role="tabpanel">{activeTab.groups.map((group, index) => <RibbonGroup key={`${activeTab.id}:${index}`} group={group} spec={spec} refresh={refresh} />)}</div></div>}</div>
  </div>, target);
}
