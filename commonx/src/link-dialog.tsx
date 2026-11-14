import { useEffect, useRef, useState } from 'react';
import { OfficeDialog } from './dialog.js';
import { useVibeOffice } from './react-bridge.js';

export interface LinkRequest {
  url: string;
  mail: boolean;
  respond: () => void;
}

interface LinkDialogProps {
  request: LinkRequest;
  onClose: () => void;
}

/** Security interstitial for document links: show and copy the address, never navigate to it. */
export function LinkDialog({ request, onClose }: LinkDialogProps) {
  const { translate } = useVibeOffice();
  const [copied, setCopied] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  const finish = () => {
    request.respond();
    onClose();
  };
  useEffect(() => {
    inputRef.current?.focus();
    inputRef.current?.select();
  }, []);

  const copy = async () => {
    try { await navigator.clipboard.writeText(request.url); }
    catch {
      const input = inputRef.current;
      if (!input) return;
      input.select();
      document.execCommand('copy');
    }
    setCopied(true);
  };

  return <OfficeDialog title={translate('Hyperlink')} className="vibeoffice-link-dialog" style={{ width: 'min(92vw, 440px)' }} onRequestClose={finish}>
    <div className="vibeoffice-link-dialog__body">
      <p>{translate('This link goes to:')}</p>
      <input ref={inputRef} id="vibeoffice-link-address" aria-label={translate('Address')} type="text" readOnly spellCheck={false} value={request.url} />
      <p>{request.mail
        ? translate('To write to this address, copy it and paste it into your e-mail program.')
        : translate("Links in documents can lead to harmful sites. If you trust this one, copy its address and paste it into your browser's address bar.")}</p>
      <p className="vibeoffice-link-dialog__copied" role="status">{copied ? translate('Copied.') : ''}</p>
    </div>
    <div className="vibeoffice-link-dialog__actions">
      <button className="primary" type="button" onClick={() => { void copy(); }}>{translate('Copy Address')}</button>
      <button type="button" onClick={finish}>{translate('Close')}</button>
    </div>
  </OfficeDialog>;
}
