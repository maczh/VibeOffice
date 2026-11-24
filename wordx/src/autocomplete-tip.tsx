import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface AutocompleteTipState { text: string; x: number; y: number }
interface AutocompleteRuntime { bus?: { on(event: string, listener: (detail?: unknown) => void): () => void } }
function runtime(host: HTMLElement): AutocompleteRuntime | undefined {
  const view = host.ownerDocument.defaultView;
  return (view as (Window & { L?: AutocompleteRuntime }) | null)?.L;
}

/** React-owned AutoComplete hint; candidate selection and insertion remain in the Wordx document engine. */
export function WordxAutocompleteTip({ host }: { host: HTMLElement }) {
  const [tip, setTip] = useState<AutocompleteTipState | null>(null);
  useEffect(() => {
    let offShow: (() => void) | undefined;
    let offHide: (() => void) | undefined;
    let attached: AutocompleteRuntime | undefined;
    const attach = () => {
      const L = runtime(host);
      if (!L?.bus || L === attached) return;
      offShow?.(); offHide?.();
      attached = L;
      offShow = L.bus.on('wordx-autocomplete-show', (detail) => {
        if (!detail || typeof detail !== 'object') return;
        const request = detail as Partial<AutocompleteTipState>;
        if (typeof request.text !== 'string' || !Number.isFinite(request.x) || !Number.isFinite(request.y)) return;
        setTip({ text: request.text, x: request.x!, y: request.y! });
      });
      offHide = L.bus.on('wordx-autocomplete-hide', () => setTip(null));
    };
    host.addEventListener('vibeoffice-engine-runtime-ready', attach);
    attach();
    return () => { host.removeEventListener('vibeoffice-engine-runtime-ready', attach); offShow?.(); offHide?.(); };
  }, [host]);

  const root = host.matches('.vibeoffice-legacy-root[data-vibeoffice-engine-root]') ? host : host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root]');
  if (!tip || !root) return null;
  return createPortal(<div className="tooltip ac-tip" role="status" aria-live="polite" data-react-autocomplete="wordx" style={{ left: tip.x, top: tip.y }}>{tip.text}</div>, root);
}
