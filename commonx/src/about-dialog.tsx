import { OfficeDialog } from './dialog.js';
import { useVibeOffice } from './react-bridge.js';

export interface AboutRequest {
  name: string;
  text: readonly string[];
  version: string;
  licensesUrl: string;
  icon?: string;
  respond?: (result: number) => void;
}

interface AboutDialogProps {
  request: AboutRequest;
  onClose: () => void;
}

const TRADEMARK_NOTE = 'Open source under the Apache License 2.0, provided as is, without warranty. VibeOffice.work is an independent project. Product names in its documentation are trademarks of their owners; VibeOffice.work is not affiliated with them.';
const FONT_NOTE = 'Font names are used solely to identify the fonts specified in documents and to maintain interoperability with Office Open XML files and the applications that create them; such names may be trademarks of their respective owners. VibeOffice renders text with the fonts installed on your computer. Where a specified font is not installed, VibeOffice substitutes an open-source font supplied with it or another font available on your computer, and the document may therefore appear differently than in other applications. VibeOffice does not include or distribute any commercial font software.';

export function AboutDialog({ request, onClose }: AboutDialogProps) {
  const { translate } = useVibeOffice();
  const finish = () => {
    request.respond?.(0);
    onClose();
  };
  const marker = '__LICENSE_LINK__';
  const licenseSentence = translate('Made possible by open source software: {0}.', marker);
  const [licenseBefore, licenseAfter = ''] = licenseSentence.split(marker);

  return <OfficeDialog title={translate('About {0}', request.name)} className="vibeoffice-about-dialog" onRequestClose={finish}>
    <div className="col about">
      <div className="about-head">
        {request.icon && <span aria-hidden="true" dangerouslySetInnerHTML={{ __html: request.icon }} />}
        <div><div className="about-name">{request.name}</div><div>{translate('Version {0}', request.version)}</div></div>
      </div>
      {request.text.map((paragraph) => <p key={paragraph}>{translate(paragraph)}</p>)}
      <p>{licenseBefore}<a className="about-oss" href={request.licensesUrl} target="_blank" rel="noopener">{translate('Open source licenses')}</a>{licenseAfter}</p>
      <p>{translate(TRADEMARK_NOTE)}</p>
      <p>{translate(FONT_NOTE)}</p>
    </div>
    <div className="vibeoffice-about-dialog__actions"><button type="button" onClick={finish}>{translate('OK')}</button></div>
  </OfficeDialog>;
}
