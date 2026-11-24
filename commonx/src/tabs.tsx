import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal, flushSync } from 'react-dom';
import { AccessLabel } from './access-label.js';

interface TabItem { label: string; body: HTMLElement; onShow?: () => void }
interface TabsTarget extends HTMLDivElement { __reactSelect?: (index: number) => void }
interface TabsRequest { target: TabsTarget; bodyHost: HTMLElement; tabs: TabItem[]; active: number; select(index: number): void }
interface TabsRuntime { ui: { ampHTML(text: string): string } }
function runtime(): TabsRuntime | undefined { return (window as Window & { L?: TabsRuntime }).L; }

function TabPanel({ request, tab, index, active }: { request: TabsRequest; tab: TabItem; index: number; active: boolean }) {
  const panelRef = useRef<HTMLDivElement>(null);
  const body = tab.body;
  const originalHidden = useRef<string | null>(null);
  useLayoutEffect(() => {
    const panel = panelRef.current;
    if (!panel) return;
    originalHidden.current = body.getAttribute('hidden');
    panel.append(body);
    return () => {
      request.bodyHost.append(body);
      if (originalHidden.current === null) body.removeAttribute('hidden');
      else body.setAttribute('hidden', originalHidden.current);
      originalHidden.current = null;
    };
  }, [body, request.bodyHost]);
  useLayoutEffect(() => { body.hidden = !active; }, [active, body]);
  return <div ref={panelRef} className="tab-pane" id={`${request.target.id.replace(/-head$/, '')}-panel-${index}`} role="tabpanel" aria-labelledby={`${request.target.id}-tab-${index}`} hidden={!active} />;
}

function TabsControl({ request }: { request: TabsRequest }) {
  const [active, setActive] = useState(request.active);
  useLayoutEffect(() => {
    const select = (index: number) => {
      flushSync(() => setActive(index));
      request.tabs[index]?.onShow?.();
    };
    request.target.__reactSelect = select;
    return () => { if (request.target.__reactSelect === select) delete request.target.__reactSelect; };
  }, [request]);
  const html = (label: string) => runtime()?.ui.ampHTML(label) || label;
  const panelId = (index: number) => `${request.target.id.replace(/-head$/, '')}-panel-${index}`;
  const headers = createPortal(<>{request.tabs.map((tab, index) => <button key={`${index}:${tab.label}`} type="button" className={`tab${index === active ? ' on' : ''}`} role="tab" id={`${request.target.id}-tab-${index}`} aria-controls={panelId(index)} aria-selected={index === active} tabIndex={index === active ? 0 : -1} onClick={() => request.select(index)} onKeyDown={(event) => {
    if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight' && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? request.tabs.length - 1 : (index + (event.key === 'ArrowRight' ? 1 : request.tabs.length - 1)) % request.tabs.length;
    request.select(next);
    request.target.querySelector<HTMLButtonElement>(`#${CSS.escape(`${request.target.id}-tab-${next}`)}`)?.focus();
  }}><AccessLabel html={html(tab.label)} /></button>)}</>, request.target);
  const panels = createPortal(<>{request.tabs.map((tab, index) => <TabPanel key={`${index}:${tab.label}`} request={request} tab={tab} index={index} active={index === active} />)}</>, request.bodyHost);
  return <>{headers}{panels}</>;
}

/** React-owned tab headers preserve the legacy DOM panel and select/onShow contract. */
export function EditorTabsLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<TabsRequest[]>([]);
  useLayoutEffect(() => {
    const onTabs = (event: Event) => {
      const request = (event as CustomEvent<TabsRequest>).detail;
      if (!request || !(request.target instanceof HTMLDivElement) || !(request.bodyHost instanceof HTMLElement) || !Array.isArray(request.tabs) || typeof request.select !== 'function') return;
      setRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-tabs', onTabs);
    const observer = new MutationObserver(() => setRequests((current) => current.filter((entry) => entry.target.isConnected)));
    observer.observe(host, { childList: true, subtree: true });
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-tabs', onTabs); };
  }, [host]);
  return <>{requests.map((request) => <TabsControl key={request.target.id} request={request} />)}</>;
}
