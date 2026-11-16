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
  const options = request.items.filter((item) => !(typeof item === 'object' && item.group) && itemValue(item) !== undefined);
  const activeIndex = Math.max(0, options.findIndex((item) => String(itemValue(item)) === String(value)));
  const activeId = request.target.id ? `${request.target.id}-option-${activeIndex}` : undefined;
  const move = (key: string, event: { preventDefault(): void }) => {
    if (!options.length) return;
    let nextIndex = activeIndex;
    if (key === 'ArrowDown') nextIndex = Math.min(options.length - 1, activeIndex + 1);
    else if (key === 'ArrowUp') nextIndex = Math.max(0, activeIndex - 1);
    else if (key === 'Home') nextIndex = 0;
    else if (key === 'End') nextIndex = options.length - 1;
    else if (key === 'Enter') { const active = options[activeIndex]; if (active) choose(active); event.preventDefault(); return; }
    else return;
    event.preventDefault();
    const next = options[nextIndex];
    if (!next) return;
    choose(next);
    const nextId = request.target.id ? `${request.target.id}-option-${nextIndex}` : undefined;
    if (nextId) {
      const option = request.target.querySelector<HTMLElement>(`#${CSS.escape(nextId)}`);
      if (option) {
        const list = request.target;
        const listBounds = list.getBoundingClientRect();
        const optionBounds = option.getBoundingClientRect();
        const visibleTop = listBounds.top + list.clientTop;
        const visibleBottom = visibleTop + list.clientHeight;
        if (optionBounds.top < visibleTop) list.scrollTop -= visibleTop - optionBounds.top;
        else if (optionBounds.bottom > visibleBottom) list.scrollTop += optionBounds.bottom - visibleBottom;
      }
    }
  };
  useEffect(() => {
    const target = request.target;
    const hadTabIndex = target.hasAttribute('tabindex');
    const previousTabIndex = target.getAttribute('tabindex');
    const hadActiveDescendant = target.hasAttribute('aria-activedescendant');
    const previousActiveDescendant = target.getAttribute('aria-activedescendant');
    if (request.inputTarget) return;
    target.setAttribute('tabindex', '0');
    target.setAttribute('aria-activedescendant', activeId || '');
    const onKeyDown = (event: KeyboardEvent) => move(event.key, event);
    target.addEventListener('keydown', onKeyDown);
    return () => {
      target.removeEventListener('keydown', onKeyDown);
      if (hadTabIndex) target.setAttribute('tabindex', previousTabIndex || ''); else target.removeAttribute('tabindex');
      if (hadActiveDescendant) target.setAttribute('aria-activedescendant', previousActiveDescendant || ''); else target.removeAttribute('aria-activedescendant');
    };
  }, [request.target, request.inputTarget, activeIndex, options, activeId]);
  return <>{request.inputTarget ? createPortal(<input id={request.inputId} className="listbox-in" type="text" autoComplete="off" aria-autocomplete="list" aria-controls={request.target.id || undefined} aria-activedescendant={activeId} translate={translate} value={inputValue} onKeyDown={(event) => move(event.key, event)} onChange={(event) => { const next = event.currentTarget.value; setInputValue(next); request.options.onInput?.(next); }} />, request.inputTarget, `${request.inputId}-input`) : null}{createPortal(<>{request.items.map((item, index) => {
    const grouped = typeof item === 'object' && !!item.group;
    const selected = !grouped && value != null && String(itemValue(item)) === String(value);
    const optionIndex = grouped ? -1 : options.findIndex((option) => String(itemValue(option)) === String(itemValue(item)));
    return <div id={optionIndex < 0 || !request.target.id ? undefined : `${request.target.id}-option-${optionIndex}`} key={`${String(itemValue(item))}:${index}`} className={grouped ? 'li-group' : `li${selected ? ' on' : ''}`} role={grouped ? undefined : 'option'} aria-selected={grouped ? undefined : selected} style={itemStyle(typeof item === 'object' ? item.style : undefined)} onClick={grouped ? undefined : () => choose(item)} onDoubleClick={grouped ? undefined : () => choose(item, true)}>{itemLabel(item, request.options.translate)}</div>;
  })}</>, request.target, request.target.id || undefined)}</>;
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
