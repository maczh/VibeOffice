import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface TabItem { label: string; body: HTMLElement; onShow?: () => void }
interface TabsTarget extends HTMLDivElement { __reactSetActive?: (index: number) => void }
interface TabsRequest { target: TabsTarget; tabs: TabItem[]; active: number; select(index: number): void }
interface TabsRuntime { ui: { ampHTML(text: string): string } }
function runtime(): TabsRuntime | undefined { return (window as Window & { L?: TabsRuntime }).L; }

function TabsControl({ request }: { request: TabsRequest }) {
  const [active, setActive] = useState(request.active);
  useEffect(() => {
    request.target.__reactSetActive = setActive;
    return () => { delete request.target.__reactSetActive; };
  }, [request.target]);
  const html = (label: string) => runtime()?.ui.ampHTML(label) || label;
  return createPortal(<>{request.tabs.map((tab, index) => <button key={`${index}:${tab.label}`} type="button" className={`tab${index === active ? ' on' : ''}`} role="tab" id={`${request.target.id}-tab-${index}`} aria-controls={`${request.target.id}-panel-${index}`} aria-selected={index === active} tabIndex={index === active ? 0 : -1} dangerouslySetInnerHTML={{ __html: html(tab.label) }} onClick={() => request.select(index)} onKeyDown={(event) => {
    if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight' && event.key !== 'Home' && event.key !== 'End') return;
    event.preventDefault();
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? request.tabs.length - 1 : (index + (event.key === 'ArrowRight' ? 1 : request.tabs.length - 1)) % request.tabs.length;
    request.select(next);
    request.target.querySelector<HTMLButtonElement>(`#${CSS.escape(`${request.target.id}-tab-${next}`)}`)?.focus();
  }} />)}</>, request.target);
}

/** React-owned tab headers preserve the legacy DOM panel and select/onShow contract. */
export function EditorTabsLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<TabsRequest[]>([]);
  useEffect(() => {
    const onTabs = (event: Event) => {
      const request = (event as CustomEvent<TabsRequest>).detail;
      if (!request || !(request.target instanceof HTMLDivElement) || !Array.isArray(request.tabs) || typeof request.select !== 'function') return;
      setRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-tabs', onTabs);
    const observer = new MutationObserver(() => setRequests((current) => current.filter((entry) => entry.target.isConnected)));
    observer.observe(host, { childList: true, subtree: true });
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-tabs', onTabs); };
  }, [host]);
  return <>{requests.map((request) => <TabsControl key={request.target.id} request={request} />)}</>;
}
