/** Static presentation engine nodes rendered by React before the existing slide services mount. */
export function PptxEngineShell() {
  return <div className="vibeoffice-legacy-root" data-vibeoffice-engine-root="" data-vibeoffice-react-shell="pptx">
    <div id="app" className="app" aria-label="PPTX presentation editor">
      <div className="cmdbars"><div className="menurow" id="menurow" /><div className="toolbars" id="toolbars" /></div>
      <div className="main" id="main">
        <div className="leftpane" id="leftpane" />
        <div className="splitter v" id="split-left" role="separator" aria-orientation="vertical" />
        <div className="center" id="center">
          <div className="editor-wrap" id="editorwrap" />
          <div className="splitter h" id="split-notes" role="separator" aria-orientation="horizontal" />
          <div className="notes" id="notes" />
        </div>
        <div className="taskpane" id="taskpane" />
      </div>
      <div className="viewbar" id="viewbar" /><div className="drawrow" id="drawrow" /><div className="statusbar" id="statusbar" />
    </div>
  </div>;
}
