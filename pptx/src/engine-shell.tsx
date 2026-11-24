import { useEffect, useRef, useState } from 'react';
import { useVibeOffice } from '@vibeoffice/core';

interface CanvasRuntime {
  pres?: unknown;
  ed?: { slide?(): unknown };
  bus?: { on(event: string, listener: () => void): () => void };
}

function presentationRuntime(shell: HTMLElement | null): CanvasRuntime | undefined {
  const view = shell?.ownerDocument.defaultView as (Window & { L?: CanvasRuntime }) | null;
  return view?.L;
}

function presentationHasSlide(runtime?: CanvasRuntime): boolean {
  if (!runtime?.pres || !runtime.ed?.slide) return false;
  try { return Boolean(runtime.ed.slide()); } catch { return false; }
}

/** Static presentation engine nodes rendered by React before the existing slide services mount. */
export function PptxEngineShell() {
  const { translate } = useVibeOffice();
  const shell = useRef<HTMLDivElement>(null);
  const [hasSlide, setHasSlide] = useState(false);
  useEffect(() => {
    const engineHost = shell.current?.closest<HTMLElement>('.vibe-editor__engine');
    if (!engineHost) return;
    let off: (() => void) | undefined;
    const attach = () => {
      off?.();
      const runtime = presentationRuntime(shell.current);
      off = runtime?.bus?.on('slide-rendered', () => setHasSlide(presentationHasSlide(runtime)));
      setHasSlide(presentationHasSlide(runtime));
    };
    engineHost.addEventListener('vibeoffice-engine-runtime-ready', attach);
    attach();
    return () => { engineHost.removeEventListener('vibeoffice-engine-runtime-ready', attach); off?.(); };
  }, []);
  return <div ref={shell} className="vibeoffice-legacy-root" data-vibeoffice-engine-root="" data-vibeoffice-react-shell="pptx">
    <div id="app" className="app" aria-label={translate('PPTX presentation editor')}>
      <div className="cmdbars"><div className="menurow" id="menurow" /><div className="toolbars" id="toolbars" /></div>
      <div className="main" id="main">
        <div className="leftpane" id="leftpane" />
        <div className="splitter v" id="split-left" role="separator" aria-orientation="vertical" />
        <div className="center" id="center">
          <div className="editor-wrap" id="editorwrap">
            <div className="ruler ruler-h" hidden />
            <div className="ruler ruler-v" hidden />
            <div className="ed-scroll" data-react-owned="pptx-canvas-surface" tabIndex={-1} aria-label={translate('Slide editing area')}>
              <div className="ed-canvas" data-react-canvas-state="pptx" hidden={!hasSlide}>
                <div className="ed-host" />
                <div className="ed-grid" hidden />
                <div className="ed-overlay" />
              </div>
              <div className="ed-empty" data-react-canvas-state="pptx" hidden={hasSlide}>{translate('Click to add first slide')}</div>
            </div>
          </div>
          <div className="splitter h" id="split-notes" role="separator" aria-orientation="horizontal" />
          <div className="notes" id="notes" />
        </div>
        <div className="taskpane" id="taskpane" />
      </div>
      <div className="viewbar" id="viewbar" /><div className="drawrow" id="drawrow" /><div className="statusbar" id="statusbar" />
    </div>
  </div>;
}
