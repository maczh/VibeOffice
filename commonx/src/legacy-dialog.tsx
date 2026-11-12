import { createElement, useEffect, useLayoutEffect, useRef, useState, type PointerEvent as ReactPointerEvent } from 'react';
import { flushSync } from 'react-dom';
import { createRoot } from 'react-dom/client';

export interface LegacyDialogButton {
  label: string;
  labelHtml: string;
  primary?: boolean;
  cancel?: boolean;
  onClick?: (close: (index: number | null) => void, button: HTMLButtonElement) => unknown;
}

export interface LegacyDialogOptions {
  title: string;
  body?: HTMLElement;
  width?: number;
  buttons?: LegacyDialogButton[];
  onClose?: (index: number | null) => void;
  help?: () => void;
  focus?: string | HTMLElement;
  onOpen?: (dialog: HTMLElement) => void;
  modeless?: boolean;
}

export interface LegacyDialogApi {
  el: HTMLDivElement;
  body: HTMLDivElement;
  close: (index: number | null) => void;
  done: Promise<number | null>;
  buttons: HTMLButtonElement[];
}

export interface LegacyDialogMountRequest {
  host: HTMLElement;
  token: object;
  options: LegacyDialogOptions;
  done: Promise<number | null>;
  resolve: (index: number | null) => void;
  stack: object[];
  restoreFocus: (dialog: HTMLElement) => void;
  api?: LegacyDialogApi;
  close: (index: number | null) => void;
  unmount?: () => void;
}

function ampAccessKey(label: string): string | null {
  const match = /&([^&\s])/.exec(label);
  return match?.[1]?.toLowerCase() ?? null;
}

function LegacyDialog({ request }: { request: LegacyDialogMountRequest }) {
  const { options } = request;
  const dialogRef = useRef<HTMLDivElement>(null);
  const bodyRef = useRef<HTMLDivElement>(null);
  const titleRef = useRef<HTMLDivElement>(null);
  const buttonRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const drag = useRef<{ x: number; y: number; left: number; top: number } | null>(null);
  const [position, setPosition] = useState({ left: 4, top: 4 });
  const buttons = options.buttons?.length ? options.buttons : [{ label: 'OK', labelHtml: 'OK', primary: true }];

  useLayoutEffect(() => {
    const body = bodyRef.current;
    const dialog = dialogRef.current;
    if (body && options.body) body.append(options.body);
    if (!dialog || !body) return;
    dialog.style.width = `${Math.min(options.width || 380, window.innerWidth - 16)}px`;
    const rect = dialog.getBoundingClientRect();
    setPosition({ left: Math.max(4, (window.innerWidth - rect.width) / 2), top: Math.max(4, (window.innerHeight - rect.height) / 2.4) });
    const api: LegacyDialogApi = {
      el: dialog,
      body,
      close: request.close,
      done: request.done,
      buttons: buttonRefs.current.filter((button): button is HTMLButtonElement => !!button),
    };
    request.api = api;
    const focus = options.focus
      ? typeof options.focus === 'string' ? body.querySelector<HTMLElement>(options.focus) : options.focus
      : body.querySelector<HTMLElement>('input:not([type=checkbox]):not([type=radio]),select,textarea')
        || buttonRefs.current.find((button, index) => buttons[index]?.primary)
        || buttonRefs.current[0];
    if (focus instanceof HTMLElement) { focus.focus(); if (focus instanceof HTMLInputElement) focus.select(); }
    options.onOpen?.(dialog);
  }, []);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (options.modeless || request.stack.at(-1) !== request.token) return;
      if (event.key === 'Escape') {
        event.preventDefault(); event.stopPropagation();
        const index = buttons.findIndex((button) => button.cancel || /^cancel$|^close$/i.test(button.label.replace(/&([^&\s])/g, '$1')));
        if (index >= 0) buttonRefs.current[index]?.click(); else request.close(null);
      } else if (event.key === 'Enter' && !event.shiftKey && !(event.target instanceof HTMLElement && (event.target.tagName === 'TEXTAREA' || event.target.tagName === 'BUTTON' || event.target.isContentEditable || event.target.tagName === 'SELECT'))) {
        const index = buttons.findIndex((button) => button.primary);
        if (index >= 0) { event.preventDefault(); event.stopPropagation(); buttonRefs.current[index]?.click(); }
      } else if (event.key === 'Tab') {
        const dialog = dialogRef.current;
        if (!dialog) return;
        const focusable = [...dialog.querySelectorAll<HTMLElement>('input,select,textarea,button,[tabindex="0"]')].filter((element) => !(element as HTMLButtonElement).disabled && element.offsetParent !== null);
        if (!focusable.length) return;
        const index = focusable.indexOf(document.activeElement as HTMLElement);
        if (event.shiftKey && index <= 0) { event.preventDefault(); focusable[focusable.length - 1]?.focus(); }
        else if (!event.shiftKey && index === focusable.length - 1) { event.preventDefault(); focusable[0]?.focus(); }
      }
    };
    document.addEventListener('keydown', onKeyDown, true);
    return () => document.removeEventListener('keydown', onKeyDown, true);
  }, [buttons, options.modeless, request]);

  const onPointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    if ((event.target as HTMLElement).closest('button')) return;
    const dialog = dialogRef.current;
    if (!dialog) return;
    drag.current = { x: event.clientX, y: event.clientY, left: dialog.offsetLeft, top: dialog.offsetTop };
    titleRef.current?.setPointerCapture(event.pointerId);
  };
  const onPointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const start = drag.current, dialog = dialogRef.current;
    if (!start || !dialog) return;
    const rect = dialog.getBoundingClientRect();
    setPosition({
      left: Math.max(-rect.width + 60, Math.min(window.innerWidth - 60, start.left + event.clientX - start.x)),
      top: Math.max(0, Math.min(window.innerHeight - 30, start.top + event.clientY - start.y)),
    });
  };

  return <div className={`dlg-overlay${options.modeless ? ' modeless' : ''}`}>
    <div ref={dialogRef} className="dlg" role="dialog" aria-modal={!options.modeless} aria-label={options.title} style={{ left: position.left, top: position.top }}>
      <div ref={titleRef} className="dlg-title" onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={() => { drag.current = null; }} onLostPointerCapture={() => { drag.current = null; }}>
        <span className="dlg-ttl">{options.title}</span>
        {options.help && <button className="dlg-x dlg-help" type="button" aria-label="Help" onClick={() => options.help?.()}>?</button>}
        <button className="dlg-x" type="button" aria-label="Close" onClick={() => request.close(null)}><svg width="9" height="9" viewBox="0 0 9 9"><path d="M1 1l7 7M8 1L1 8" stroke="currentColor" strokeWidth="1.5" /></svg></button>
      </div>
      <div ref={bodyRef} className="dlg-body" />
      <div className="dlg-foot">{buttons.map((button, index) => <button
        key={`${button.label}:${index}`}
        ref={(element) => { buttonRefs.current[index] = element; }}
        className={`btn${button.primary ? ' primary' : ''}`}
        type="button"
        dangerouslySetInnerHTML={{ __html: button.labelHtml }}
        onClick={(event) => {
          const result = button.onClick?.(request.close, event.currentTarget);
          if (result !== false) request.close(index);
        }}
      />)}</div>
    </div>
  </div>;
}

/** Attach the synchronous legacy dialog API to React-owned dialog chrome. */
export function attachLegacyDialog(host: HTMLElement, request: LegacyDialogMountRequest): LegacyDialogApi {
  const rootElement = host.ownerDocument.createElement('div');
  rootElement.className = 'vibeoffice-legacy-dialog-root';
  host.append(rootElement);
  const root = createRoot(rootElement);
  const remove = () => {
    flushSync(() => root.unmount());
    rootElement.remove();
  };
  request.unmount = remove;
  flushSync(() => root.render(createElement(LegacyDialog, { request })));
  if (!request.api) { remove(); throw new Error('React dialog did not initialize synchronously'); }
  return request.api;
}
