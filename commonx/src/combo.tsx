import { useEffect, useRef, useState, type CSSProperties } from 'react';
import { createPortal } from 'react-dom';

type ComboOption = string | number | { heading?: string; sep?: boolean; value?: unknown; label?: string; run?: () => unknown };
type ComboNote = string | { text?: string; tip?: string } | null | undefined;
interface ComboOptions {
  id: string;
  width?: number;
  tip?: string;
  options(): ComboOption[];
  value(): unknown;
  onChange(value: string): void;
  enabled?(): boolean;
  note?(value: unknown): ComboNote;
  preview?(value: unknown): string;
  listCls?: string;
}
interface ComboRequest { target: HTMLDivElement; options: ComboOptions }
interface ComboRuntime {
  ui: {
    openMenu(items: unknown[], at: { left: number; bottom: number }, options?: { cls?: string }): void;
    hooks: { commitEdit(): void; refocus(): void };
  };
  bus: { on(event: string, callback: () => void): () => void };
}
function runtime(): ComboRuntime | undefined { return (window as Window & { L?: ComboRuntime }).L; }
function currentValue(options: ComboOptions): string { const value = options.value(); return value == null ? '' : String(value); }
function cssStyle(text?: string): CSSProperties | undefined {
  if (!text) return undefined;
  const declaration = document.createElement('span').style;
  declaration.cssText = text;
  return Object.fromEntries(Array.from({ length: declaration.length }, (_, index) => declaration.item(index)).map((property) => [property.replace(/-([a-z])/g, (_match, letter: string) => letter.toUpperCase()), declaration.getPropertyValue(property)])) as CSSProperties;
}

function ComboControl({ request, refresh }: { request: ComboRequest; refresh: number }) {
  const office = runtime();
  const { options, target } = request;
  const inputRef = useRef<HTMLInputElement>(null);
  const buttonRef = useRef<HTMLButtonElement>(null);
  const [focused, setFocused] = useState(false);
  const [draft, setDraft] = useState(currentValue(options));
  useEffect(() => { if (!focused) setDraft(currentValue(options)); }, [focused, options, refresh]);
  if (!office) return null;
  const isEnabled = () => { try { return !options.enabled || !!options.enabled(); } catch { return false; } };
  const openList = () => {
    office.ui.hooks.commitEdit();
    const rect = target.getBoundingClientRect();
    const items = options.options().map((option) => {
      if (typeof option === 'object' && option && 'heading' in option && option.heading) return { heading: option.heading };
      if (typeof option === 'object' && option && 'sep' in option && option.sep) return { sep: true };
      if (typeof option === 'object' && option && 'run' in option && option.run) {
        return { label: option.label, run: async () => { await option.run?.(); buttonRef.current?.click(); } };
      }
      const value = typeof option === 'object' && option ? option.value : option;
      const label = typeof option === 'object' && option ? option.label : undefined;
      const note = options.note?.(value);
      const noteText = typeof note === 'object' && note ? note.text : note;
      const noteTip = typeof note === 'object' && note ? note.tip : undefined;
      return {
        label: String(label || value).replace(/&/g, '&&'), raw: true,
        ...(options.preview ? { labelStyle: cssStyle(options.preview(value)) } : {}),
        value,
        run: () => { const text = String(value); setDraft(text); options.onChange(text); office.ui.hooks.refocus(); },
        checked: () => String(value) === currentValue(options), key: noteText || '', tip: noteTip,
      };
    });
    office.ui.openMenu(items, { left: rect.left, bottom: rect.bottom }, { cls: `combo-list${options.listCls ? ` ${options.listCls}` : ''}` });
  };
  const commit = () => { options.onChange(draft); setFocused(false); inputRef.current?.blur(); office.ui.hooks.refocus(); };
  return createPortal(<>
    <input ref={inputRef} type="text" className="combo-in" id={options.id} aria-label={options.tip || options.id} autoComplete="off" spellCheck={false} disabled={!isEnabled()} value={draft} onChange={(event) => setDraft(event.currentTarget.value)} onFocus={(event) => { setFocused(true); event.currentTarget.select(); office.ui.hooks.commitEdit(); }} onBlur={() => setFocused(false)} onKeyDown={(event) => {
      if (event.key === 'Enter') { event.preventDefault(); commit(); }
      if (event.key === 'Escape') { event.preventDefault(); setDraft(currentValue(options)); setFocused(false); inputRef.current?.blur(); office.ui.hooks.refocus(); }
      event.stopPropagation();
    }} />
    <button ref={buttonRef} className="combo-btn" type="button" tabIndex={-1} aria-label="Open list" onPointerDown={(event) => event.preventDefault()} onClick={openList}><span className="dd-arrow" /></button>
  </>, target);
}

/** React-owned editable combo controls; engine models retain option/value and edit callbacks. */
export function EditorComboLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<ComboRequest[]>([]);
  const [refresh, setRefresh] = useState(0);
  useEffect(() => {
    const onCombo = (event: Event) => {
      const request = (event as CustomEvent<ComboRequest>).detail;
      if (!request || !(request.target instanceof HTMLDivElement) || !request.options || typeof request.options.value !== 'function') return;
      setRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-combo', onCombo);
    const observer = new MutationObserver(() => setRequests((current) => current.filter((entry) => entry.target.isConnected)));
    observer.observe(host, { childList: true, subtree: true });
    const offRefresh = runtime()?.bus.on('ui-refresh', () => setRefresh((current) => current + 1));
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-combo', onCombo); offRefresh?.(); };
  }, [host]);
  return <>{requests.map((request) => createPortal(<ComboControl request={request} refresh={refresh} />, request.target, request.target.id))}</>;
}
