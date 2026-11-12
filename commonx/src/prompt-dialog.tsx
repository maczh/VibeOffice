import { useState } from 'react';
import { OfficeDialog } from './dialog.js';
import { useVibeOffice } from './react-bridge.js';

export interface PromptRequest {
  title: string;
  label: string;
  value: string;
  respond: (value: string | null) => void;
}

interface PromptDialogProps {
  request: PromptRequest;
  onClose: () => void;
}

/** Shared React text prompt used by legacy editor features. */
export function PromptDialog({ request, onClose }: PromptDialogProps) {
  const { translate } = useVibeOffice();
  const [value, setValue] = useState(request.value);
  const finish = (result: string | null) => {
    request.respond(result);
    onClose();
  };

  return <OfficeDialog title={translate(request.title)} className="vibeoffice-prompt-dialog" onRequestClose={() => finish(null)}>
    <form onSubmit={(event) => { event.preventDefault(); finish(value); }}>
      <label htmlFor="vibeoffice-prompt-input">{translate(request.label)}</label>
      <input id="vibeoffice-prompt-input" type="text" autoFocus value={value} onChange={(event) => setValue(event.currentTarget.value)} />
      <div className="vibeoffice-prompt-dialog__actions">
        <button className="primary" type="submit">{translate('OK')}</button>
        <button type="button" onClick={() => finish(null)}>{translate('Cancel')}</button>
      </div>
    </form>
  </OfficeDialog>;
}
