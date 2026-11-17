import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';

interface ContentOptions { html: string; id: string }
interface ContentRequest { target: HTMLElement; options: ContentOptions }
interface FieldRequest { kind: 'field'; target: HTMLElement; labelHtml: string; labelFor: string | null; control: HTMLElement }
interface GroupRequest { kind: 'group'; target: HTMLElement; legendHtml: string; children: HTMLElement[] }
type StructuredRequest = FieldRequest | GroupRequest;

function ContentControl({ request }: { request: ContentRequest }) {
  return createPortal(<span className="vibeoffice-react-ui-content" dangerouslySetInnerHTML={{ __html: request.options.html }} />, request.target, request.options.id);
}

function FieldControl({ request }: { request: FieldRequest }) {
  const controlSlot = useRef<HTMLSpanElement>(null);
  useLayoutEffect(() => {
    const slot = controlSlot.current;
    if (!slot) return;
    slot.append(request.control);
    return () => { request.target.append(request.control); };
  }, [request]);
  return createPortal(<><label htmlFor={request.labelFor ?? undefined}><span className="vibeoffice-react-ui-content" dangerouslySetInnerHTML={{ __html: request.labelHtml }} /></label><span ref={controlSlot} style={{ display: 'contents' }} /></>, request.target);
}

function GroupControl({ request }: { request: GroupRequest }) {
  const childSlot = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const slot = childSlot.current;
    if (!slot) return;
    request.children.forEach((child) => slot.append(child));
    return () => { request.children.forEach((child) => request.target.append(child)); };
  }, [request]);
  return createPortal(<><legend><span className="vibeoffice-react-ui-content" dangerouslySetInnerHTML={{ __html: request.legendHtml }} /></legend><div ref={childSlot} className="vibeoffice-react-group-content" /></>, request.target);
}

/** React-owned shared field/group structure and translated content around compatible engine DOM controls. */
export function EditorUiContentLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<ContentRequest[]>([]);
  const [structuredRequests, setStructuredRequests] = useState<StructuredRequest[]>([]);
  useEffect(() => {
    const onContent = (event: Event) => {
      const request = (event as CustomEvent<ContentRequest>).detail;
      if (!request || !(request.target instanceof HTMLElement) || !request.options || typeof request.options.html !== 'string') return;
      setRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    const onStructured = (event: Event) => {
      const request = (event as CustomEvent<StructuredRequest>).detail;
      if (!request || !(request.target instanceof HTMLElement) || (request.kind !== 'field' && request.kind !== 'group')) return;
      if (request.kind === 'field' && (!(request.control instanceof HTMLElement) || typeof request.labelHtml !== 'string')) return;
      if (request.kind === 'group' && (!Array.isArray(request.children) || typeof request.legendHtml !== 'string')) return;
      setStructuredRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-ui-content', onContent);
    host.addEventListener('vibeoffice-react-ui-structure', onStructured);
    const observer = new MutationObserver(() => {
      setRequests((current) => current.filter((entry) => entry.target.isConnected));
      setStructuredRequests((current) => current.filter((entry) => entry.target.isConnected));
    });
    observer.observe(host, { childList: true, subtree: true });
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-ui-content', onContent); host.removeEventListener('vibeoffice-react-ui-structure', onStructured); };
  }, [host]);
  return <>{requests.map((request) => <ContentControl key={request.options.id} request={request} />)}{structuredRequests.map((request) => request.kind === 'field'
    ? <FieldControl key={`field:${request.target.dataset.reactUiId}`} request={request} />
    : <GroupControl key={`group:${request.target.dataset.reactUiId}`} request={request} />)}</>;
}
