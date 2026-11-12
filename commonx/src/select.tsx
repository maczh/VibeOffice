import { useEffect, useState, type CSSProperties } from 'react';
import { createPortal } from 'react-dom';

type SelectOption = [unknown, unknown] | unknown;
interface SelectOptions {
  id: string;
  options: SelectOption[];
  value: unknown;
  disabled?: boolean;
  translate?: string;
  style?: string | CSSProperties;
  class?: string;
  title?: string;
  onChange?: (value: string) => void;
}
interface SelectRequest { target: HTMLDivElement; options: SelectOptions }
function cssStyle(value?: string | CSSProperties): CSSProperties | undefined {
  if (!value) return undefined;
  if (typeof value !== 'string') return value;
  const declaration = document.createElement('span').style;
  declaration.cssText = value;
  return Object.fromEntries(Array.from({ length: declaration.length }, (_, index) => declaration.item(index)).map((property) => [property.replace(/-([a-z])/g, (_match, letter: string) => letter.toUpperCase()), declaration.getPropertyValue(property)])) as CSSProperties;
}

function SelectControl({ request }: { request: SelectRequest }) {
  const { target, options } = request;
  const initial = String(options.value ?? '');
  return createPortal(<select id={options.id} className={options.class} style={cssStyle(options.style)} title={options.title} translate={options.translate === 'no' ? 'no' : undefined} disabled={!!options.disabled} defaultValue={initial} onChange={(event) => options.onChange?.(event.currentTarget.value)}>
    {options.options.map((option, index) => {
      const pair = Array.isArray(option);
      const value = pair ? option[0] : option;
      const label = pair ? option[1] : option;
      return <option key={`${String(value)}:${index}`} value={String(value)} translate={!pair || options.translate === 'no' ? 'no' : undefined}>{String(label)}</option>;
    })}
  </select>, target, target.id || options.id);
}

/** React-owned native selects for shared legacy feature APIs in direct-host mode. */
export function EditorSelectLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<SelectRequest[]>([]);
  useEffect(() => {
    const onSelect = (event: Event) => {
      const request = (event as CustomEvent<SelectRequest>).detail;
      if (!request || !(request.target instanceof HTMLDivElement) || !request.options || !Array.isArray(request.options.options)) return;
      setRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-select', onSelect);
    const observer = new MutationObserver(() => setRequests((current) => current.filter((entry) => entry.target.isConnected)));
    observer.observe(host, { childList: true, subtree: true });
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-select', onSelect); };
  }, [host]);
  return <>{requests.map((request) => createPortal(<SelectControl request={request} />, request.target, request.target.id || request.options.id))}</>;
}
