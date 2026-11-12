import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';

interface StatusRequest { text: string; duration?: number }

/** React-owned status text for the legacy editor status-bar layout. */
export function EditorStatusMessage({ host }: { host: HTMLElement }) {
  const [target, setTarget] = useState<HTMLElement | null>(null);
  const [text, setText] = useState('');
  const timer = useRef<number | undefined>(undefined);

  useEffect(() => {
    const findTarget = () => host.querySelector<HTMLElement>('[data-react-status]');
    const updateTarget = () => setTarget(findTarget());
    updateTarget();
    const observer = new MutationObserver(updateTarget);
    observer.observe(host, { childList: true, subtree: true });
    const onStatus = (event: Event) => {
      const request = (event as CustomEvent<StatusRequest>).detail;
      if (!request || typeof request.text !== 'string') return;
      window.clearTimeout(timer.current);
      setText(request.text);
      if (request.duration && request.duration > 0) {
        timer.current = window.setTimeout(() => setText((current) => current === request.text ? '' : current), request.duration);
      }
    };
    host.addEventListener('vibeoffice-react-status', onStatus);
    return () => { observer.disconnect(); host.removeEventListener('vibeoffice-react-status', onStatus); window.clearTimeout(timer.current); };
  }, [host]);

  return target ? createPortal(<span className="vibeoffice-react-status" data-react-status-content>{text}</span>, target) : null;
}
