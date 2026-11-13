import { useEffect, useState, type CSSProperties } from 'react';
import { createPortal } from 'react-dom';

type ListItem = string | number | { value?: string | number; label?: string; style?: string; group?: boolean };
interface ListBoxOptions { translate?: string; onInput?(value: string): void; onChange?(value: string | number): void; onDbl?(value: string | number): void; pick?(value: string | number, double: boolean): void }
interface ListBoxRequest { target: HTMLDivElement; inputTarget: HTMLDivElement | null; inputId: string; inputValue: string; items: ListItem[]; value: unknown; options: ListBoxOptions }
interface ListBoxRuntime { t(text: string): string }
function runtime(): ListBoxRuntime | undefined { return (window as Window & { L?: ListBoxRuntime }).L; }
function itemValue(item: ListItem): string | number | undefined { return typeof item === 'object' ? item.value : item; }
function itemLabel(item: ListItem, translate?: string): string {
  const label = typeof item === 'object' ? item.label ?? String(item.value) : String(item);
  return translate === 'no' ? label : runtime()?.t(label) || label;
}
function itemStyle(text?: string): CSSProperties | undefined {
  if (!text) return undefined;
  const style = document.createElement('div').style;
  style.cssText = text;
  return Object.fromEntries(Array.from({ length: style.length }, (_, index) => style.item(index)).map((property) => [property.replace(/-([a-z])/g, (_match, letter: string) => letter.toUpperCase()), style.getPropertyValue(property)])) as CSSProperties;
}
function ListBoxControl({ request }: { request: ListBoxRequest }) {
  const [value, setValue] = useState(request.value);
  const [inputValue, setInputValue] = useState(request.inputValue);
  const translate = request.options.translate === 'no' ? 'no' : request.options.translate === 'yes' ? 'yes' : undefined;
  useEffect(() => { setValue(request.value); }, [request]);
  useEffect(() => { setInputValue(request.inputValue); }, [request.inputValue]);
  const choose = (item: ListItem, double = false) => {
    const next = itemValue(item);
    setValue(next);
    setInputValue(itemLabel(item, request.options.translate));
    if (next !== undefined) {
      if (request.options.pick) request.options.pick(next, double);
      else { request.options.onChange?.(next); if (double) request.options.onDbl?.(next); }
    }
  };
  return <>{request.inputTarget ? createPortal(<input id={request.inputId} className="listbox-in" type="text" autoComplete="off" aria-autocomplete="list" aria-controls={request.target.id || undefined} translate={translate} value={inputValue} onChange={(event) => { const next = event.currentTarget.value; setInputValue(next); request.options.onInput?.(next); }} />, request.inputTarget, `${request.inputId}-input`) : null}{createPortal(<>{request.items.map((item, index) => {
    const grouped = typeof item === 'object' && !!item.group;
    const selected = !grouped && value != null && String(itemValue(item)) === String(value);
    return <div key={`${String(itemValue(item))}:${index}`} className={grouped ? 'li-group' : `li${selected ? ' on' : ''}`} role={grouped ? undefined : 'option'} aria-selected={grouped ? undefined : selected} style={itemStyle(typeof item === 'object' ? item.style : undefined)} onClick={grouped ? undefined : () => choose(item)} onDoubleClick={grouped ? undefined : () => choose(item, true)}>{itemLabel(item, request.options.translate)}</div>;
  })}</>, request.target, request.target.id)}</>;
}

/** Wordx feature listbox rows rendered by React; the existing helper retains its synchronous input/value API. */
export function WordxListBoxLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<ListBoxRequest[]>([]);
  useEffect(() => {
    const onListBox = (event: Event) => {
      const request = (event as CustomEvent<ListBoxRequest>).detail;
      if (!request || !(request.target instanceof HTMLDivElement) || (request.inputTarget != null && !(request.inputTarget instanceof HTMLDivElement)) || !Array.isArray(request.items) || !request.options) return;
      setRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-wordx-list-box', onListBox);
    const observer = new MutationObserver(() => setRequests((current) => current.filter((entry) => entry.target.isConnected)));
    observer.observe(host, { childList: true, subtree: true });
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-wordx-list-box', onListBox); };
  }, [host]);
  return <>{requests.map((request) => <ListBoxControl key={request.target.id} request={request} />)}</>;
}
