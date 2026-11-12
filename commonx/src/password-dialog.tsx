import { useState } from 'react';
import { OfficeDialog } from './dialog.js';
import { useVibeOffice } from './react-bridge.js';

export interface PasswordRequest {
  text: string;
  wrong: boolean;
  respond: (password: string | null) => void;
}

interface PasswordDialogProps {
  request: PasswordRequest;
  onClose: () => void;
}

/** Password prompt used when opening protected Office files. */
export function PasswordDialog({ request, onClose }: PasswordDialogProps) {
  const { translate } = useVibeOffice();
  const [password, setPassword] = useState('');
  const finish = (result: string | null) => {
    request.respond(result);
    onClose();
  };

  return <OfficeDialog title={translate('Password')} className="vibeoffice-password-dialog" style={{ width: 'min(92vw, 340px)' }} onRequestClose={() => finish(null)}>
    <form onSubmit={(event) => { event.preventDefault(); finish(password); }}>
      <p>{request.wrong ? translate('The password is incorrect. Try again.') : request.text}</p>
      <label htmlFor="vibeoffice-password-input">{translate('Enter password to open file')}</label>
      <input id="vibeoffice-password-input" type="password" autoComplete="off" autoFocus value={password} onChange={(event) => setPassword(event.currentTarget.value)} />
      <div className="vibeoffice-password-dialog__actions">
        <button className="primary" type="submit">{translate('OK')}</button>
        <button type="button" onClick={() => finish(null)}>{translate('Cancel')}</button>
      </div>
    </form>
  </OfficeDialog>;
}
