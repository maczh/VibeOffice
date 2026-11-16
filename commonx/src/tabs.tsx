import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal, flushSync } from 'react-dom';

interface TabItem { label: string; body: HTMLElement; onShow?: () => void }
interface TabsTarget extends HTMLDivElement { __reactSelect?: (index: number) => void }
interface TabsRequest { target: TabsTarget; tabs: TabItem[]; active: number; select(index: number): void }
interface TabsRuntime { ui: { ampHTML(text: string): string } }
function runtime(): TabsRuntime | undefined { return (window as Window & { L?: TabsRuntime }).L; }

function TabsControl({ request }: { request: TabsRequest }) {
  const [active, setActive] = useState(request.active);
  const originalPanels = useRef<Array<{ id: string | null; role: string | null; labelledby: string | null; tabPane: boolean; hidden: string | null }> | null>(null);
  useLayoutEffect(() => {
    const select = (index: number) => {
      flushSync(() => setActive(index));
      request.tabs[index]?.onShow?.();
    };
    request.target.__reactSelect = select;
    return () => { if (request.target.__reactSelect === select) delete request.target.__reactSelect; };
  }, [request]);
  useLayoutEffect(() => {
    originalPanels.current = request.tabs.map((tab) => ({
      id: tab.body.getAttribute('id'),
      role: tab.body.getAttribute('role'),
      labelledby: tab.body.getAttribute('aria-labelledby'),
      tabPane: tab.body.classList.contains('tab-pane'),
      hidden: tab.body.getAttribute('hidden'),
    }));
    request.tabs.forEach((tab, index) => {
      tab.body.classList.add('tab-pane');
      const tabId = `${request.target.id}-tab-${index}`;
      tab.body.id = `${request.target.id.replace(/-head$/, '')}-panel-${index}`;
      tab.body.setAttribute('role', 'tabpanel');
      tab.body.setAttribute('aria-labelledby', tabId);
    });
    return () => {
      request.tabs.forEach((tab, index) => {
        const original = originalPanels.current?.[index];
        if (!original) return;
        const restoreAttribute = (name: string, value: string | null) => value === null ? tab.body.removeAttribute(name) : tab.body.setAttribute(name, value);
        restoreAttribute('id', original.id);
        restoreAttribute('role', original.role);
        restoreAttribute('aria-labelledby', original.labelledby);
        restoreAttribute('hidden', original.hidden);
        if (!original.tabPane) tab.body.classList.remove('tab-pane');
      });
      originalPanels.current = null;
    };
  }, [request.target, request.tabs]);
  useLayoutEffect(() => {
    request.tabs.forEach((tab, index) => { tab.body.hidden = index !== active; });
  }, [active, request.tabs]);
  const html = (label: string) => runtime()?.ui.ampHTML(label) || label;
  const panelId = (index: number) => `${request.target.id.replace(/-head$/, '')}-panel-${index}`;
  return createPortal(<>{request.tabs.map((tab, index) => <button key={`${index}:${tab.label}`} type="button" className={`tab${index === active ? ' on' : ''}`} role="tab" id={`${request.target.id}-tab-${index}`} aria-controls={panelId(index)} aria-selected={index === active} tabIndex={index === active ? 0 : -1} dangerouslySetInnerHTML={{ __html: html(tab.label) }} onClick={() => request.select(index)} onKeyDown={(event) => {
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
