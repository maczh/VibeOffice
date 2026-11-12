import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface ContentOptions { html: string; id: string }
interface ContentRequest { target: HTMLElement; options: ContentOptions }

function ContentControl({ request }: { request: ContentRequest }) {
  return createPortal(<span className="vibeoffice-react-ui-content" dangerouslySetInnerHTML={{ __html: request.options.html }} />, request.target, request.options.id);
}

/** React-owned translated text for shared DOM-compatible field labels and group legends. */
export function EditorUiContentLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<ContentRequest[]>([]);
  useEffect(() => {
    const onContent = (event: Event) => {
      const request = (event as CustomEvent<ContentRequest>).detail;
      if (!request || !(request.target instanceof HTMLElement) || !request.options || typeof request.options.html !== 'string') return;
      setRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-ui-content', onContent);
    const observer = new MutationObserver(() => setRequests((current) => current.filter((entry) => entry.target.isConnected)));
    observer.observe(host, { childList: true, subtree: true });
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-ui-content', onContent); };
  }, [host]);
  return <>{requests.map((request) => <ContentControl key={request.options.id} request={request} />)}</>;
}
