import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { AccessLabel } from './access-label.js';

interface FieldRequest { kind: 'field'; target: HTMLElement; labelHtml: string; labelFor: string | null; control: HTMLElement }
interface GroupRequest { kind: 'group'; target: HTMLElement; legendHtml: string; children: HTMLElement[] }
type StructuredRequest = FieldRequest | GroupRequest;

function FieldControl({ request }: { request: FieldRequest }) {
  const controlSlot = useRef<HTMLSpanElement>(null);
  useLayoutEffect(() => {
    const slot = controlSlot.current;
    if (!slot) return;
    slot.append(request.control);
    return () => { request.target.append(request.control); };
  }, [request]);
  return createPortal(<><label htmlFor={request.labelFor ?? undefined}><span className="vibeoffice-react-ui-content"><AccessLabel html={request.labelHtml} /></span></label><span ref={controlSlot} style={{ display: 'contents' }} /></>, request.target);
}

function GroupControl({ request }: { request: GroupRequest }) {
  const childSlot = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const slot = childSlot.current;
    if (!slot) return;
    request.children.forEach((child) => slot.append(child));
    return () => { request.children.forEach((child) => request.target.append(child)); };
  }, [request]);
  return createPortal(<><legend><span className="vibeoffice-react-ui-content"><AccessLabel html={request.legendHtml} /></span></legend><div ref={childSlot} className="vibeoffice-react-group-content" /></>, request.target);
}

/** React-owned shared field/group structure and translated content around compatible engine DOM controls. */
export function EditorUiContentLayer({ host }: { host: HTMLElement }) {
  const [structuredRequests, setStructuredRequests] = useState<StructuredRequest[]>([]);
  useEffect(() => {
    const onStructured = (event: Event) => {
      const request = (event as CustomEvent<StructuredRequest>).detail;
      if (!request || !(request.target instanceof HTMLElement) || (request.kind !== 'field' && request.kind !== 'group')) return;
      if (request.kind === 'field' && (!(request.control instanceof HTMLElement) || typeof request.labelHtml !== 'string')) return;
      if (request.kind === 'group' && (!Array.isArray(request.children) || typeof request.legendHtml !== 'string')) return;
      setStructuredRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-ui-structure', onStructured);
    const observer = new MutationObserver(() => {
      setStructuredRequests((current) => current.filter((entry) => entry.target.isConnected));
    });
    observer.observe(host, { childList: true, subtree: true });
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-ui-structure', onStructured); };
  }, [host]);
  return <>{structuredRequests.map((request) => request.kind === 'field'
    ? <FieldControl key={`field:${request.target.dataset.reactUiId}`} request={request} />
    : <GroupControl key={`group:${request.target.dataset.reactUiId}`} request={request} />)}</>;
}
