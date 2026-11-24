import { useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import { AccessLabel } from './access-label.js';
import { SvgArtwork } from './svg-artwork.js';

type ButtonOptions = Record<string, unknown> & { id: string; label: string; html?: string };
interface ButtonRequest { target: HTMLButtonElement; options: ButtonOptions }
type RegisteredButton = HTMLButtonElement & { __vibeofficeReactButtonRequest?: ButtonRequest };
interface ButtonRuntime { ui: { ampHTML(text: string): string } }
function runtime(): ButtonRuntime | undefined { return (window as Window & { L?: ButtonRuntime }).L; }
function ButtonControl({ request }: { request: ButtonRequest }) {
  const { target, options } = request;
  const html = options.html ?? runtime()?.ui.ampHTML(options.label) ?? options.label;
  const markup = String(html);
  const content = /^\s*<svg(?:\s|>)/i.test(markup)
    ? <SvgArtwork markup={markup} owner="button-icon" />
    : <AccessLabel html={markup} />;
  return createPortal(<span className="vibeoffice-react-button-label">{content}</span>, target, options.id);
}

/** React-owned shared button rendering with a DOM-compatible proxy for legacy callers. */
export function EditorButtonLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<ButtonRequest[]>([]);
  useLayoutEffect(() => {
    const remember = (request: ButtonRequest, refresh = false) => setRequests((current) => {
      const index = current.findIndex((entry) => entry.target === request.target);
      if (index < 0) return [...current, request];
      if (!refresh && current[index] === request) return current;
      const next = current.slice();
      next[index] = request;
      return next;
    });
    const onButton = (event: Event) => {
      const request = (event as CustomEvent<ButtonRequest>).detail;
      if (!request || !(request.target instanceof HTMLButtonElement) || !request.options || typeof request.options.id !== 'string') return;
      remember(request, true);
    };
    const recover = () => {
      for (const target of host.querySelectorAll<RegisteredButton>('[data-vibeoffice-react-button]')) {
        const request = target.__vibeofficeReactButtonRequest;
        if (request && request.target === target && request.options && typeof request.options.id === 'string') remember(request);
      }
    };
    host.addEventListener('vibeoffice-react-button', onButton);
    const observer = new MutationObserver(() => {
      recover();
      setRequests((current) => current.filter((entry) => entry.target.isConnected));
    });
    observer.observe(host, { childList: true, subtree: true });
    recover();
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-button', onButton); };
  }, [host]);
  return <>{requests.map((request) => <ButtonControl key={request.target.id} request={request} />)}</>;
}
