import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';

interface ToolbarCommand {
  label?: string;
  tbLabel?: string;
  tip?: string;
  key?: string;
  icon?: string;
  enabled?: () => boolean;
  checked?: () => boolean;
}

interface ToolbarMenuAnchor { left: number; bottom: number }
interface ToolbarButtonControl { __vibeToolbarControl: 'button'; id: string; opts: { label?: string; icon?: string; arg?: unknown } }
interface ToolbarSplitControl { __vibeToolbarControl: 'split'; id: string; opts: { label?: string; icon?: string; arg?: unknown }; menuFn: (anchor: ToolbarMenuAnchor, target: HTMLElement) => void }
interface ToolbarDropControl { __vibeToolbarControl: 'drop'; label: string; icon: string; tip: string; dataAs?: string; setAttribute?(name: string, value: string): void; menuFn: (anchor: ToolbarMenuAnchor, target: HTMLElement) => void }
type ToolbarControl = ToolbarButtonControl | ToolbarSplitControl | ToolbarDropControl;

interface ToolbarRuntime {
  app?: { reactToolbarVisibility?(): Record<string, boolean> };
  ui: {
    cmds: Record<string, ToolbarCommand>;
    exec(id: string, arg?: unknown, event?: Event): unknown;
    stripAmp(text: string): string;
    ampHTML(text: string): string;
  };
  icons: { get(name: string): string };
  bus: { on(event: string, callback: () => void): () => void };
}

interface ToolbarRequest {
  id: string;
  title: string;
  target: HTMLDivElement;
  items: Array<string | Node | ToolbarControl>;
  kind?: 'toolbar' | 'viewbar';
}

function runtime(): ToolbarRuntime | undefined { return (window as Window & { L?: ToolbarRuntime }).L; }
function commandEnabled(command: ToolbarCommand): boolean { try { return !command.enabled || !!command.enabled(); } catch { return false; } }
function commandChecked(command: ToolbarCommand): boolean { try { return !!command.checked?.(); } catch { return false; } }

function CommandButton({ id, refresh, options, reactControl }: { id: string; refresh: number; options?: ToolbarSplitControl['opts']; reactControl?: boolean }) {
  const office = runtime();
  if (!office) return null;
  const command = office.ui.cmds[id] || {};
  const label = options?.label ?? command.tbLabel;
  const isEnabled = commandEnabled(command);
  const isChecked = commandChecked(command);
  const tip = office.ui.stripAmp(command.tip || command.label || id) + (command.key ? ` (${command.key})` : '');
  const content = (options?.icon || command.icon ? office.icons.get(options?.icon || command.icon || '') : '') + (label ? `<span class="tb-lbl">${office.ui.ampHTML(label)}</span>` : '');
  return <button
    key={`${id}:${refresh}`}
    type="button"
    className={`tb-btn${label ? ' with-label' : ''}${isChecked ? ' on' : ''}${isEnabled ? '' : ' dis'}`}
    data-cmd={id}
    data-react-command={id}
    data-react-toolbar-control={reactControl ? 'button' : undefined}
    data-tip={tip}
    aria-label={office.ui.stripAmp(command.label || id)}
    aria-pressed={command.checked ? isChecked : undefined}
    disabled={!isEnabled}
    dangerouslySetInnerHTML={{ __html: content }}
    onPointerDown={(event) => event.preventDefault()}
    onClick={(event) => office.ui.exec(id, options?.arg, event.nativeEvent)}
  />;
}

function ReactToolbarControl({ control, refresh }: { control: ToolbarControl; refresh: number }) {
  const trigger = useRef<HTMLButtonElement>(null);
  if (control.__vibeToolbarControl === 'button') return <CommandButton id={control.id} refresh={refresh} options={control.opts} reactControl />;
  if (control.__vibeToolbarControl === 'split') {
    const command = runtime()?.ui.cmds[control.id] || {};
    const open = () => {
      const rect = trigger.current?.parentElement?.getBoundingClientRect();
      if (rect) control.menuFn({ left: rect.left, bottom: rect.bottom }, trigger.current!.parentElement!);
    };
    return <div className="tb-split" data-cmd={control.id} data-react-toolbar-control="split"><CommandButton id={control.id} refresh={refresh} options={control.opts} /><button ref={trigger} className="tb-dd" type="button" aria-label="More options" data-tip={command.label || ''} onPointerDown={(event) => event.preventDefault()} onClick={open}><span className="dd-arrow" /></button></div>;
  }
  const office = runtime();
  const open = () => {
    const rect = trigger.current?.getBoundingClientRect();
    if (rect) control.menuFn({ left: rect.left, bottom: rect.bottom }, trigger.current!);
  };
  const accessibleLabel = office?.ui.stripAmp(control.tip || control.label) || control.tip || control.label;
  const content = (control.icon && office ? office.icons.get(control.icon) : '') + (control.label && office ? `<span class="tb-lbl">${office.ui.ampHTML(control.label)}</span>` : '') + '<span class="dd-arrow"></span>';
  return <button ref={trigger} className={`tb-btn tb-drop${control.label ? ' with-label' : ''}`} type="button" data-tip={control.tip} aria-label={accessibleLabel} data-react-toolbar-control="drop" data-as={control.dataAs || undefined} dangerouslySetInnerHTML={{ __html: content }} onPointerDown={(event) => event.preventDefault()} onClick={open} />;
}

function isToolbarControl(item: Node | ToolbarControl): item is ToolbarControl {
  return !(item instanceof Node) && (item.__vibeToolbarControl === 'button' || item.__vibeToolbarControl === 'split' || item.__vibeToolbarControl === 'drop');
}

function LegacyToolbarNode({ node }: { node: Node }) {
  const slot = useRef<HTMLSpanElement>(null);
  useLayoutEffect(() => {
    slot.current?.append(node);
    return () => { if (node.parentNode) node.parentNode.removeChild(node); };
  }, [node]);
  return <span ref={slot} style={{ display: 'contents' }} />;
}

/** React owns the shared toolbar row and simple command buttons; app-specific DOM controls remain mounted in slots. */
export function EditorToolbarLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<ToolbarRequest[]>([]);
  const [refresh, setRefresh] = useState(0);
  const originalHidden = useRef(new Map<HTMLElement, boolean>());

  useEffect(() => {
    const onToolbar = (event: Event) => {
      const request = (event as CustomEvent<ToolbarRequest>).detail;
      if (!request || !(request.target instanceof HTMLElement) || !Array.isArray(request.items)) return;
      setRequests((current) => [...current.filter((item) => item.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-toolbar', onToolbar);
    const offRefresh = runtime()?.bus.on('ui-refresh', () => setRefresh((current) => current + 1));
    return () => { host.removeEventListener('vibeoffice-react-toolbar', onToolbar); offRefresh?.(); setRequests([]); };
  }, [host]);

  useLayoutEffect(() => {
    const visibility = runtime()?.app?.reactToolbarVisibility?.();
    if (!visibility) return;
    for (const request of requests) {
      const { target } = request;
      if (!originalHidden.current.has(target)) originalHidden.current.set(target, target.hidden);
      target.hidden = visibility[request.id] === false;
    }
    for (const row of host.querySelectorAll<HTMLElement>('.tb-row')) {
      if (!originalHidden.current.has(row)) originalHidden.current.set(row, row.hidden);
      row.hidden = Array.from(row.children).filter((child): child is HTMLElement => child instanceof HTMLElement && child.classList.contains('toolbar')).every((toolbar) => toolbar.hidden);
    }
    const drawingRow = host.querySelector<HTMLElement>('#drawrow');
    const drawing = requests.find((request) => request.id === 'drawing');
    if (drawingRow && drawing) {
      if (!originalHidden.current.has(drawingRow)) originalHidden.current.set(drawingRow, drawingRow.hidden);
      drawingRow.hidden = drawing.target.hidden;
    }
  }, [host, requests, refresh]);

  useEffect(() => () => {
    for (const [element, hidden] of originalHidden.current) element.hidden = hidden;
    originalHidden.current.clear();
  }, [host]);

  return <>{requests.map((request) => {
    const children = request.items.map((item, index) => {
      if (item === '|') return <span key={`sep:${index}`} className="tb-sep" />;
      if (typeof item === 'string') return <CommandButton key={`cmd:${item}:${index}`} id={item} refresh={refresh} />;
      if (isToolbarControl(item)) return <ReactToolbarControl key={`${item.__vibeToolbarControl}:${index}`} control={item} refresh={refresh} />;
      if (item instanceof Node) return <LegacyToolbarNode key={`node:${index}`} node={item} />;
      return null;
    });
    return createPortal(request.kind === 'viewbar' ? <>{children}</> : <><span className="tb-grip" aria-hidden="true" />{children}<span className="tb-end" /></>, request.target, request.id);
  })}</>;
}
