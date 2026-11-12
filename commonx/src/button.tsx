import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

type ButtonOptions = Record<string, unknown> & { id: string; label: string; html?: string };
interface ButtonRequest { target: HTMLButtonElement; options: ButtonOptions }
interface ButtonRuntime { ui: { ampHTML(text: string): string } }
function runtime(): ButtonRuntime | undefined { return (window as Window & { L?: ButtonRuntime }).L; }
function ButtonControl({ request }: { request: ButtonRequest }) {
  const { target, options } = request;
  const html = options.html ?? runtime()?.ui.ampHTML(options.label) ?? options.label;
  return createPortal(<span className="vibeoffice-react-button-label" dangerouslySetInnerHTML={{ __html: String(html) }} />, target, options.id);
}

/** React-owned shared button rendering with a DOM-compatible proxy for legacy callers. */
export function EditorButtonLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<ButtonRequest[]>([]);
  useEffect(() => {
    const onButton = (event: Event) => {
      const request = (event as CustomEvent<ButtonRequest>).detail;
      if (!request || !(request.target instanceof HTMLButtonElement) || !request.options || typeof request.options.id !== 'string') return;
      setRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-button', onButton);
    const observer = new MutationObserver(() => setRequests((current) => current.filter((entry) => entry.target.isConnected)));
    observer.observe(host, { childList: true, subtree: true });
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-button', onButton); };
  }, [host]);
  return <>{requests.map((request) => <ButtonControl key={request.target.id} request={request} />)}</>;
}
