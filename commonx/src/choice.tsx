import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface ChoiceOptions {
  id: string;
  type: 'checkbox' | 'radio';
  name?: string;
  label: string;
  checked: boolean;
  onChange?: (checked?: boolean) => void;
}
interface ChoiceRequest { target: HTMLSpanElement; options: ChoiceOptions }
interface ChoiceRuntime { ui: { ampHTML(text: string): string } }
function runtime(): ChoiceRuntime | undefined { return (window as Window & { L?: ChoiceRuntime }).L; }

function ChoiceControl({ request }: { request: ChoiceRequest }) {
  const { target, options } = request;
  const office = runtime();
  const [html, setHtml] = useState(() => office?.ui.ampHTML(options.label) || options.label);
  useEffect(() => { setHtml(office?.ui.ampHTML(options.label) || options.label); }, [office, options.label]);
  const changed = (input: HTMLInputElement) => {
    if (options.type === 'checkbox') options.onChange?.(input.checked);
    else if (input.checked) options.onChange?.();
  };
  return createPortal(<label className="ck" htmlFor={options.id}>
    <input id={options.id} type={options.type} name={options.name} defaultChecked={options.checked} onChange={(event) => changed(event.currentTarget)} />
    <span dangerouslySetInnerHTML={{ __html: html }} />
  </label>, target, target.dataset.reactChoiceId || options.id);
}

/** React-owned checkbox/radio controls preserving the legacy imperative input reference. */
export function EditorChoiceLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<ChoiceRequest[]>([]);
  useEffect(() => {
    const onChoice = (event: Event) => {
      const request = (event as CustomEvent<ChoiceRequest>).detail;
      if (!request || !(request.target instanceof HTMLSpanElement) || !request.options || !['checkbox', 'radio'].includes(request.options.type)) return;
      setRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-choice', onChoice);
    const observer = new MutationObserver(() => setRequests((current) => current.filter((entry) => entry.target.isConnected)));
    observer.observe(host, { childList: true, subtree: true });
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-choice', onChoice); };
  }, [host]);
  return <>{requests.map((request) => createPortal(<ChoiceControl request={request} />, request.target, request.target.dataset.reactChoiceId || request.options.id))}</>;
}
