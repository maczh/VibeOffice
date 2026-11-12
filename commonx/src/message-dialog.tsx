import { OfficeDialog } from './dialog.js';
import { useVibeOffice } from './react-bridge.js';

export interface MessageRequest {
  title: string;
  text: string;
  icon: 'info' | 'warn' | 'question' | 'error';
  iconSvg: string;
  buttons: string[];
  defaultButton: number;
  escapeIndex: number;
  width: number;
  respond: (result: number | null) => void;
}

interface MessageDialogProps {
  request: MessageRequest;
  onClose: () => void;
}

function visibleButtonLabel(value: string): { label: string; accessKey?: string } {
  const translated = value;
  const access = /&([^&\s])/.exec(translated)?.[1];
  const label = translated.replace(/&&/g, '\u0000').replace(/&(\S)/g, '$1').replace(/\u0000/g, '&');
  return { label, ...(access ? { accessKey: access } : {}) };
}

export function MessageDialog({ request, onClose }: MessageDialogProps) {
  const { translate } = useVibeOffice();
  const finish = (result: number | null) => {
    request.respond(result);
    onClose();
  };
  const cancel = () => finish(request.escapeIndex >= 0 ? request.escapeIndex : null);
  const lines = request.text.split('\n');

  return <OfficeDialog title={translate(request.title)} className={`vibeoffice-message-dialog vibeoffice-message-dialog--${request.icon}`} style={{ width: `min(92vw, ${request.width}px)` }} onRequestClose={cancel}>
    <div className="vibeoffice-message-dialog__body">
      <span aria-hidden="true" dangerouslySetInnerHTML={{ __html: request.iconSvg }} />
      <div role="document">{lines.map((line, index) => <span key={`${index}:${line}`}>{index > 0 && <br />}{line}</span>)}</div>
    </div>
    <div className="vibeoffice-message-dialog__actions">
      {request.buttons.map((rawLabel, index) => {
        const { label, accessKey } = visibleButtonLabel(translate(rawLabel));
        return <button key={label} type="button" accessKey={accessKey} className={index === request.defaultButton ? 'primary' : undefined} onClick={() => finish(index)}>{label}</button>;
      })}
    </div>
  </OfficeDialog>;
}
