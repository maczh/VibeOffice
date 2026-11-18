import { useVibeOffice } from '@vibeoffice/core';

/** React-owned static DOM contract consumed by the existing Wordx document engine. */
export function WordxEngineShell() {
  const { translate } = useVibeOffice();
  return <div className="vibeoffice-legacy-root" data-vibeoffice-engine-root data-vibeoffice-react-shell="wordx">
    <div id="app" className="app" aria-label={translate('Wordx word processor')}>
      <div className="cmdbars">
        <div className="menurow" id="menurow" />
        <div className="toolbars" id="toolbars" />
      </div>
      <div className="main" id="main">
        <div className="docmap" id="docmap" hidden />
        <div className="splitter v" id="split-map" hidden role="separator" aria-orientation="vertical" />
        <div className="center" id="center">
          <div className="work" id="work">
            <div className="r-corner" id="rcorner" />
            <div className="ruler-h" id="hruler" />
            <div className="ruler-v" id="vruler" />
            <div className="scroller" id="scroller" tabIndex={-1}>
              <div className="sizer" data-react-owned="wordx-page-sizer">
                <div id="pages" className="pages" data-react-owned="wordx-pages" contentEditable="true" spellCheck={false} role="textbox" aria-multiline="true" aria-label={translate('Document')} />
              </div>
            </div>
          </div>
          <div className="viewbar" id="viewbar" />
          <div className="revpane" id="revpane" hidden />
        </div>
        <div className="taskpane" id="taskpane" />
      </div>
      <div className="drawrow" id="drawrow" />
      <div className="statusbar" id="statusbar" />
    </div>
  </div>;
}
