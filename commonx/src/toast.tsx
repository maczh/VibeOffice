import { useEffect } from 'react';
import { useVibeOffice } from './react-bridge.js';

export interface ToastRequest {
  text: string;
  duration: number;
  action?: { label: string; run: () => void };
}

interface ToastProps {
  request: ToastRequest;
  onDismiss: () => void;
}

export function Toast({ request, onDismiss }: ToastProps) {
  const { translate } = useVibeOffice();
  useEffect(() => {
    const timer = window.setTimeout(onDismiss, request.duration);
    return () => window.clearTimeout(timer);
  }, [onDismiss, request.duration, request.text]);

  const runAction = () => {
    request.action?.run();
    onDismiss();
  };

  return <div className="toast vibeoffice-toast" role="status" aria-live="polite">
    <span>{request.text}</span>
    {request.action && <button className="toast-act" type="button" onClick={runAction}>{translate(request.action.label)}</button>}
  </div>;
}
