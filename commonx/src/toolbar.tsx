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

interface ToolbarRuntime {
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
  items: Array<string | Node>;
}

function runtime(): ToolbarRuntime | undefined { return (window as Window & { L?: ToolbarRuntime }).L; }
function commandEnabled(command: ToolbarCommand): boolean { try { return !command.enabled || !!command.enabled(); } catch { return false; } }
function commandChecked(command: ToolbarCommand): boolean { try { return !!command.checked?.(); } catch { return false; } }

function CommandButton({ id, refresh }: { id: string; refresh: number }) {
  const office = runtime();
  if (!office) return null;
  const command = office.ui.cmds[id] || {};
  const label = command.tbLabel;
  const isEnabled = commandEnabled(command);
  const isChecked = commandChecked(command);
  const tip = office.ui.stripAmp(command.tip || command.label || id) + (command.key ? ` (${command.key})` : '');
  const content = (command.icon ? office.icons.get(command.icon) : '') + (label ? `<span class="tb-lbl">${office.ui.ampHTML(label)}</span>` : '');
  return <button
    key={`${id}:${refresh}`}
    type="button"
    className={`tb-btn${label ? ' with-label' : ''}${isChecked ? ' on' : ''}${isEnabled ? '' : ' dis'}`}
    data-cmd={id}
    data-react-command={id}
    data-tip={tip}
    aria-label={office.ui.stripAmp(command.label || id)}
    aria-pressed={command.checked ? isChecked : undefined}
    disabled={!isEnabled}
    dangerouslySetInnerHTML={{ __html: content }}
    onPointerDown={(event) => event.preventDefault()}
    onClick={(event) => office.ui.exec(id, undefined, event.nativeEvent)}
  />;
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

  return <>{requests.map((request) => {
    const children = request.items.map((item, index) => {
      if (item === '|') return <span key={`sep:${index}`} className="tb-sep" />;
      if (typeof item === 'string') return <CommandButton key={`cmd:${item}:${index}`} id={item} refresh={refresh} />;
      if (item instanceof Node) return <LegacyToolbarNode key={`node:${index}`} node={item} />;
      return null;
    });
    return createPortal(<><span className="tb-grip" aria-hidden="true" />{children}<span className="tb-end" /></>, request.target, request.id);
  })}</>;
}
