import { OfficeDialog } from './dialog.js';
import { useVibeOffice } from './react-bridge.js';

export interface CompatibilityGroup {
  what: string;
  places: string[];
}

export interface CompatibilityRequest {
  groups: CompatibilityGroup[];
  respond: (proceed: boolean) => void;
}

interface CompatibilityDialogProps {
  request: CompatibilityRequest;
  onClose: () => void;
}

/** Save-time explanation of content that cannot be preserved in the chosen format. */
export function CompatibilityDialog({ request, onClose }: CompatibilityDialogProps) {
  const { translate } = useVibeOffice();
  const finish = (proceed: boolean) => {
    request.respond(proceed);
    onClose();
  };
  const cancel = () => finish(false);

  return <OfficeDialog title={translate('Compatibility Checker')} className="vibeoffice-compatibility-dialog" style={{ width: 'min(92vw, 540px)' }} onRequestClose={cancel}>
    <div className="vibeoffice-compatibility-dialog__body">
      <p>{translate("Some of this file can't be saved exactly as it was:")}</p>
      <ul>{request.groups.map(({ what, places }, index) => <li key={`${what}:${index}`}>
        <div>{what}</div>
        {places.length > 0 && <div className="hint">{places.length > 4 ? translate('{0} and {1} more', places.slice(0, 4).join(', '), places.length - 4) : places.join(', ')}</div>}
      </li>)}</ul>
      <p>{translate('Click Continue to save anyway, or Cancel to go back to your document without saving.')}</p>
    </div>
    <div className="vibeoffice-compatibility-dialog__actions">
      <button className="primary" type="button" accessKey="c" onClick={() => finish(true)}>{translate('Continue')}</button>
      <button type="button" onClick={cancel}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
