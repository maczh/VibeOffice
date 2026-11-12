import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface SpinOptions {
  id: string;
  value: number;
  min?: number;
  max?: number;
  step?: number;
  dec?: number;
  unit?: string;
  fmt?(value: number): string;
  onChange?(value: number): void;
}
interface SpinRequest { target: HTMLSpanElement; options: SpinOptions }
interface SpinTarget extends HTMLSpanElement { __reactSpinDisabled?: boolean }
interface SpinRuntime { round(value: number, digits?: number): number; clamp(value: number, min: number, max: number): number }
function runtime(): SpinRuntime | undefined { return (window as Window & { L?: SpinRuntime }).L; }
function format(options: SpinOptions, value: number): string {
  if (options.fmt) return options.fmt(value);
  const digits = options.dec ?? 2;
  return `${Math.round((value + Number.EPSILON) * 10 ** digits) / 10 ** digits}${options.unit || ''}`;
}
function parse(options: SpinOptions, input: HTMLInputElement | null): number {
  if (!input) return options.value;
  const value = parseFloat(String(input.value).replace(/[^0-9.\-]/g, ''));
  return Number.isNaN(value) ? options.value : value;
}

function SpinControl({ request }: { request: SpinRequest }) {
  const { options, target } = request;
  const [disabled, setDisabled] = useState(!!(target as SpinTarget).__reactSpinDisabled);
  const step = options.step || 1;
  const change = (direction: number) => {
    const input = target.querySelector<HTMLInputElement>('input');
    const office = runtime();
    const next = office?.clamp(office.round(parse(options, input) + direction * step, 4), options.min ?? -1e9, options.max ?? 1e9)
      ?? Math.min(options.max ?? 1e9, Math.max(options.min ?? -1e9, parse(options, input) + direction * step));
    if (input) input.value = format(options, next);
    options.value = next;
    options.onChange?.(next);
  };
  useEffect(() => {
    const updateDisabled = () => setDisabled(!!(target as SpinTarget).__reactSpinDisabled);
    target.addEventListener('vibeoffice-react-spin-disabled', updateDisabled);
    return () => target.removeEventListener('vibeoffice-react-spin-disabled', updateDisabled);
  }, [target]);
  return createPortal(<>
    <input type="text" className="spin-in" id={options.id} inputMode="decimal" autoComplete="off" defaultValue={format(options, options.value)} disabled={disabled} onBlur={(event) => { const value = parse(options, event.currentTarget); if (value !== options.value) { options.value = value; options.onChange?.(value); } }} onKeyDown={(event) => {
      if (event.key === 'ArrowUp' || event.key === 'ArrowDown') { event.preventDefault(); change(event.key === 'ArrowUp' ? 1 : -1); }
    }} />
    <span className="spin-bs"><button type="button" className="spin-b up" tabIndex={-1} aria-label="Increase" disabled={disabled} onClick={() => change(1)}><span /></button><button type="button" className="spin-b dn" tabIndex={-1} aria-label="Decrease" disabled={disabled} onClick={() => change(-1)}><span /></button></span>
  </>, target, target.id || options.id);
}

/** React-owned numeric spinner retaining the legacy imperative get/set control contract. */
export function EditorSpinLayer({ host }: { host: HTMLElement }) {
  const [requests, setRequests] = useState<SpinRequest[]>([]);
  useEffect(() => {
    const onSpin = (event: Event) => {
      const request = (event as CustomEvent<SpinRequest>).detail;
      if (!request || !(request.target instanceof HTMLSpanElement) || !request.options || typeof request.options.value !== 'number') return;
      setRequests((current) => [...current.filter((entry) => entry.target !== request.target), request]);
    };
    host.addEventListener('vibeoffice-react-spin', onSpin);
    const observer = new MutationObserver(() => setRequests((current) => current.filter((entry) => entry.target.isConnected)));
    observer.observe(host, { childList: true, subtree: true });
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-spin', onSpin); };
  }, [host]);
  return <>{requests.map((request) => createPortal(<SpinControl request={request} />, request.target, request.target.dataset.reactSpinId || request.options.id))}</>;
}
