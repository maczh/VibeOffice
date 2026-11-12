import { useVibeOffice } from './react-bridge.js';

export interface BusyRequest {
  text: string;
}

interface BusyOverlayProps {
  request: BusyRequest;
}

/** React-owned modal progress veil for long-running editor operations. */
export function BusyOverlay({ request }: BusyOverlayProps) {
  const { translate } = useVibeOffice();
  return <div className="busy vibeoffice-busy" role="status" aria-live="polite" aria-busy="true">
    <div className="busy-box"><span className="busy-spin" aria-hidden="true" /><span className="busy-t">{translate(request.text)}</span></div>
  </div>;
}
