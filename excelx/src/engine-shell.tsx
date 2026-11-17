/** Static workbook engine nodes rendered by React before the existing spreadsheet services mount. */
export function ExcelxEngineShell() {
  return <div className="vibeoffice-legacy-root" data-vibeoffice-engine-root="" data-vibeoffice-react-shell="excelx">
    <div id="app" className="app" aria-label="Excelx spreadsheet">
      <div className="cmdbars"><div className="menurow" id="menurow" /><div className="toolbars" id="toolbars" /></div>
      <div className="fbar" id="fbar">
        <div className="namebox" id="namebox-wrap">
          <input id="namebox" aria-label="Name Box" autoComplete="off" spellCheck={false} />
          <div id="react-namebox" />
        </div>
        <div className="fb-btns fb-react-btns" id="react-fbar-buttons" />
        <div className="fb-edit" id="fb-edit">
          <textarea className="fb-input" id="fbinput" rows={1} aria-label="Formula Bar" spellCheck={false} autoComplete="off" />
          <div id="react-fbar" />
        </div>
      </div>
      <div className="main" id="main">
        <div className="center" id="center">
          <div className="gridwrap cur-cell" id="gridwrap" tabIndex={0} aria-label="Worksheet grid" role="grid">
            <div className="glayer" id="glayer" /><div className="sb vscroll" id="vscroll" />
          </div>
          <div className="tabrow" id="tabrow">
            <div className="tabnav" id="tabnav" /><div className="tabs" id="tabs" role="tablist" aria-label="Sheets" />
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
