import { useVibeOffice } from '@vibeoffice/core';

/** Static workbook engine nodes rendered by React before the existing spreadsheet services mount. */
export function ExcelxEngineShell() {
  const { translate } = useVibeOffice();
  return <div className="vibeoffice-legacy-root" data-vibeoffice-engine-root="" data-vibeoffice-react-shell="excelx">
    <div id="app" className="app" aria-label={translate('Excelx spreadsheet')}>
      <div className="cmdbars"><div className="menurow" id="menurow" /><div className="toolbars" id="toolbars" /></div>
      <div className="fbar" id="fbar">
        <div className="namebox" id="namebox-wrap">
          <input id="namebox" aria-label={translate('Name Box')} autoComplete="off" spellCheck={false} />
          <div id="react-namebox" />
        </div>
        <div className="fb-btns fb-react-btns" id="react-fbar-buttons" />
        <div className="fb-edit" id="fb-edit">
          <textarea className="fb-input" id="fbinput" rows={1} aria-label={translate('Formula Bar')} spellCheck={false} autoComplete="off" />
          <div id="react-fbar" />
        </div>
      </div>
      <div className="main" id="main">
        <div className="center" id="center">
          <div className="gridwrap cur-cell" id="gridwrap" tabIndex={0} aria-label={translate('Worksheet grid')} role="grid">
            <canvas className="gcv" aria-hidden="true" data-react-owned="excelx-grid-canvas" />
            <div className="glayer" id="glayer" /><div className="sb vscroll" id="vscroll" />
            <textarea className="celledit" id="celledit" data-react-owned="excelx-cell-editor" aria-label={translate('Cell editor')} spellCheck={false} autoComplete="off" autoCapitalize="off" wrap="off" rows={1} style={{ opacity: 0, pointerEvents: 'none' }} />
          </div>
          <div className="tabrow" id="tabrow">
            <div className="tabnav" id="tabnav" /><div className="tabs" id="tabs" role="tablist" aria-label={translate('Worksheets')} />
            <div className="tabsplit" id="tabsplit" role="separator" aria-orientation="vertical" />
            <div className="sb hscroll" id="hscroll" /><div className="sb-corner" />
          </div>
        </div>
        <div className="taskpane" id="taskpane" hidden />
      </div>
      <div className="drawrow" id="drawrow" hidden /><div className="statusbar" id="statusbar" />
    </div>
  </div>;
}
