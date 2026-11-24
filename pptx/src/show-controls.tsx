import { useCallback, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { PptxPreviewArtwork } from './slides-pane.js';

interface ShowStartOptions { from?: number; rehearse?: boolean; preview?: number }
interface ShowElements { root: HTMLElement; stage: HTMLElement; inkCanvas: HTMLCanvasElement; popbar: HTMLElement }
interface ShowRuntime {
  show?: { reactMounted?: boolean; reactEvents?: { pointerdown(event: PointerEvent): void; pointermove(event: PointerEvent): void; pointerup(event: PointerEvent): void; pointercancel(event: PointerEvent): void; contextmenu(event: MouseEvent): void; wheel(event: WheelEvent): void; blur(): void; resize(): void; keydown(event: KeyboardEvent): void }; startReact?(options: ShowStartOptions, elements: ShowElements): void; prev(): void; next(): void; openPointerMenu(anchor: HTMLElement): void; openShowMenu(anchor: HTMLElement): void; gotoCatalog?(): string | null; overlayCatalog?(): { blank: string | null; ended: boolean } | null; rehearseCatalog?(): { current: string; total: string; paused: boolean } | null; toggleRehearsePause?(): void; repeatRehearse?(): void };
  icons?: { get(name: string): string };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}
function runtime(host: HTMLElement): ShowRuntime | undefined {
  return (host.ownerDocument.defaultView as (Window & { L?: ShowRuntime }) | null)?.L;
}

/** React-owned navigation controls for the engine-owned full-screen slide show. */
export function PptxShowControlsLayer({ host }: { host: HTMLElement }) {
  const L = runtime(host);
  const show = L?.show;
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [rehearseMount, setRehearseMount] = useState<HTMLElement | null>(null);
  const [gotoMount, setGotoMount] = useState<HTMLElement | null>(null);
  const [overlayMount, setOverlayMount] = useState<HTMLElement | null>(null);
  const [showRoot, setShowRoot] = useState<HTMLElement | null>(null);
  const [showRequest, setShowRequest] = useState<{ id: number; options: ShowStartOptions } | null>(null);
  const showRootRef = useRef<HTMLElement | null>(null);
  const startRequestId = useRef(0);
  const [revision, setRevision] = useState(0);
  const bindShowRoot = useCallback((element: HTMLElement | null) => {
    showRootRef.current = element;
    setShowRoot(element);
    const ownedSlot = (selector: string) => {
      const slot = element?.querySelector<HTMLElement>(selector) || null;
      slot?.classList.add('vibeoffice-react-owned');
      return slot;
    };
    setMount(ownedSlot('.popbar'));
    setOverlayMount(ownedSlot('#react-show-overlay-root'));
    setGotoMount(ownedSlot('#react-goto-root'));
    setRehearseMount(ownedSlot('#react-rehearse-root'));
  }, []);
  useEffect(() => {
    let slot: HTMLElement | null = null;
    let timerSlot: HTMLElement | null = null;
    let gotoSlot: HTMLElement | null = null;
    let overlaySlot: HTMLElement | null = null;
    let offRehearse: (() => void) | undefined;
    let offGoto: (() => void) | undefined;
    let offOverlay: (() => void) | undefined;
    const findShowSlot = (selector: string) => showRootRef.current?.querySelector<HTMLElement>(selector) || host.querySelector<HTMLElement>(selector);
    const onShowStart = (event: Event) => {
      const options = (event as CustomEvent<ShowStartOptions>).detail || {};
      setShowRequest({ id: ++startRequestId.current, options });
    };
    const onShowEnd = (event: Event) => {
      if (showRootRef.current !== (event as CustomEvent<HTMLElement>).detail) return;
      showRootRef.current = null;
      setShowRoot(null);
      setShowRequest(null);
    };
    host.addEventListener('vibeoffice-react-show-start', onShowStart);
    host.addEventListener('vibeoffice-react-show-end', onShowEnd);
    const attach = () => {
      const engine = runtime(host);
      if (engine?.show) engine.show.reactMounted = true;
      if (!offRehearse && engine?.bus) offRehearse = engine.bus.on('pptx-show-rehearse', () => setRevision((value) => value + 1));
      if (!offGoto && engine?.bus) offGoto = engine.bus.on('pptx-show-goto', () => {
        setGotoMount(findShowSlot('#react-goto-root'));
        setRevision((value) => value + 1);
      });
      if (!offOverlay && engine?.bus) offOverlay = engine.bus.on('pptx-show-overlay', () => {
        setOverlayMount(findShowSlot('#react-show-overlay-root'));
        setRevision((value) => value + 1);
      });
      const target = findShowSlot('.popbar');
      if (target !== slot) {
        slot?.classList.remove('vibeoffice-react-owned');
        slot = target;
        slot?.classList.add('vibeoffice-react-owned');
        setMount(slot);
      }
      const timer = findShowSlot('#react-rehearse-root');
      if (timer !== timerSlot) {
        timerSlot?.classList.remove('vibeoffice-react-owned');
        timerSlot = timer;
        timerSlot?.classList.add('vibeoffice-react-owned');
        setRehearseMount(timerSlot);
      }
      const goto = findShowSlot('#react-goto-root');
      if (goto !== gotoSlot) {
        gotoSlot?.classList.remove('vibeoffice-react-owned');
        gotoSlot = goto;
        gotoSlot?.classList.add('vibeoffice-react-owned');
        setGotoMount(gotoSlot);
      }
      const overlay = findShowSlot('#react-show-overlay-root');
      if (overlay !== overlaySlot) {
        overlaySlot?.classList.remove('vibeoffice-react-owned');
        overlaySlot = overlay;
        overlaySlot?.classList.add('vibeoffice-react-owned');
        setOverlayMount(overlaySlot);
      }
    };
    const Observer = host.ownerDocument.defaultView?.MutationObserver || MutationObserver;
    const observer = new Observer(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer.disconnect();
      offRehearse?.();
      offGoto?.();
      offOverlay?.();
      host.removeEventListener('vibeoffice-react-show-start', onShowStart);
      host.removeEventListener('vibeoffice-react-show-end', onShowEnd);
      slot?.classList.remove('vibeoffice-react-owned');
      timerSlot?.classList.remove('vibeoffice-react-owned');
      gotoSlot?.classList.remove('vibeoffice-react-owned');
      overlaySlot?.classList.remove('vibeoffice-react-owned');
      const show = runtime(host)?.show;
      if (show) show.reactMounted = false;
    };
  }, [host]);
  useEffect(() => {
    const events = show?.reactEvents;
    if (!showRoot || !events) return;
    const ownerDocument = host.ownerDocument;
    const ownerWindow = ownerDocument.defaultView;
    if (!ownerWindow) return;
    showRoot.addEventListener('pointerdown', events.pointerdown);
    showRoot.addEventListener('pointermove', events.pointermove);
    showRoot.addEventListener('pointerup', events.pointerup);
    showRoot.addEventListener('pointercancel', events.pointercancel);
    showRoot.addEventListener('contextmenu', events.contextmenu);
    showRoot.addEventListener('wheel', events.wheel, { passive: true });
    ownerDocument.addEventListener('keydown', events.keydown, true);
    ownerWindow.addEventListener('blur', events.blur);
    ownerWindow.addEventListener('resize', events.resize);
    return () => {
      showRoot.removeEventListener('pointerdown', events.pointerdown);
      showRoot.removeEventListener('pointermove', events.pointermove);
      showRoot.removeEventListener('pointerup', events.pointerup);
      showRoot.removeEventListener('pointercancel', events.pointercancel);
      showRoot.removeEventListener('contextmenu', events.contextmenu);
      showRoot.removeEventListener('wheel', events.wheel);
      ownerDocument.removeEventListener('keydown', events.keydown, true);
      ownerWindow.removeEventListener('blur', events.blur);
      ownerWindow.removeEventListener('resize', events.resize);
    };
  }, [show, showRoot]);
  useEffect(() => {
    if (!showRequest || !showRoot || !show?.startReact) return;
    const stage = showRoot.querySelector<HTMLElement>('.stage');
    const inkCanvas = showRoot.querySelector<HTMLCanvasElement>('canvas.ink');
    const popbar = showRoot.querySelector<HTMLElement>('.popbar');
    if (!stage || !inkCanvas || !popbar) return;
    show.startReact(showRequest.options, { root: showRoot, stage, inkCanvas, popbar });
  }, [show, showRoot, showRequest]);
  const rehearsal = show?.rehearseCatalog?.();
  const gotoBuffer = show?.gotoCatalog?.();
  const overlay = show?.overlayCatalog?.();
  void revision;
  if (!show) return null;
  const control = (name: string, icon: string, action: (button: HTMLButtonElement) => void) => <button type="button" aria-label={name} onClick={(event) => { event.stopPropagation(); action(event.currentTarget); }}><PptxPreviewArtwork markup={L.icons?.get(icon) || ''} /></button>;
  return <>
    {showRequest && <div className="show" data-react-show-shell="true" tabIndex={0} role="application" aria-label={L.t?.('Slide Show') || 'Slide Show'} ref={bindShowRoot}>
      <div className="stage" />
      <canvas className="ink" />
      <div className="popbar" />
      <div id="react-show-overlay-root" />
      <div className="gotobox" id="react-goto-root" />
      <div className="rehearse" id="react-rehearse-root" role="toolbar" aria-label={L.t?.('Rehearsal') || 'Rehearsal'} />
    </div>}
    {mount && createPortal(<>
      {control('Previous', 'showPrev', () => show.prev())}
      {control('Pointer options', 'showPen', (button) => show.openPointerMenu(button))}
      {control('Slide show menu', 'showMenu', (button) => show.openShowMenu(button))}
      {control('Next', 'showNext', () => show.next())}
    </>, mount)}
    {rehearseMount && rehearsal && createPortal(<>
      <button className="tb-btn" type="button" aria-label="Next" onClick={(event) => { event.stopPropagation(); show.next(); }}><PptxPreviewArtwork markup={L.icons?.get('forward') || ''} /></button>
      <button className="tb-btn" type="button" aria-label={rehearsal.paused ? 'Resume' : 'Pause'} onClick={(event) => { event.stopPropagation(); show.toggleRehearsePause?.(); }}>{rehearsal.paused ? '▶' : '❚❚'}</button>
      <span className="rt">{rehearsal.current}</span>
      <button className="tb-btn" type="button" aria-label="Repeat" onClick={(event) => { event.stopPropagation(); show.repeatRehearse?.(); }}><PptxPreviewArtwork markup={L.icons?.get('undo') || ''} /></button>
      <span className="tot">{rehearsal.total}</span>
    </>, rehearseMount)}
    {gotoMount && gotoBuffer != null && createPortal(L.t?.('Go to slide {0}', gotoBuffer) || `Go to slide ${gotoBuffer}`, gotoMount)}
    {overlayMount && overlay?.blank && createPortal(<div className="blank" style={{ background: overlay.blank }} aria-label={overlay.blank === '#000' ? 'Black screen' : 'White screen'} />, overlayMount)}
    {overlayMount && overlay?.ended && createPortal(<div className="endscreen">End of slide show, click to exit.</div>, overlayMount)}
  </>;
}
