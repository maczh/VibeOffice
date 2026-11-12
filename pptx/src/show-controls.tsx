import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface ShowRuntime {
  show?: { reactMounted?: boolean; prev(): void; next(): void; openPointerMenu(anchor: HTMLElement): void; openShowMenu(anchor: HTMLElement): void; gotoCatalog?(): string | null; overlayCatalog?(): { blank: string | null; ended: boolean } | null; rehearseCatalog?(): { current: string; total: string; paused: boolean } | null; toggleRehearsePause?(): void; repeatRehearse?(): void };
  icons?: { get(name: string): string };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}
function runtime(): ShowRuntime | undefined { return (window as Window & { L?: ShowRuntime }).L; }

/** React-owned navigation controls for the engine-owned full-screen slide show. */
export function PptxShowControlsLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [rehearseMount, setRehearseMount] = useState<HTMLElement | null>(null);
  const [gotoMount, setGotoMount] = useState<HTMLElement | null>(null);
  const [overlayMount, setOverlayMount] = useState<HTMLElement | null>(null);
  const [revision, setRevision] = useState(0);
  useEffect(() => {
    let slot: HTMLElement | null = null;
    let timerSlot: HTMLElement | null = null;
    let gotoSlot: HTMLElement | null = null;
    let overlaySlot: HTMLElement | null = null;
    let offRehearse: (() => void) | undefined;
    let offGoto: (() => void) | undefined;
    let offOverlay: (() => void) | undefined;
    const attach = () => {
      const engine = runtime();
      if (engine?.show) engine.show.reactMounted = true;
      if (!offRehearse && engine?.bus) offRehearse = engine.bus.on('pptx-show-rehearse', () => setRevision((value) => value + 1));
      if (!offGoto && engine?.bus) offGoto = engine.bus.on('pptx-show-goto', () => {
        setGotoMount(host.querySelector<HTMLElement>('#react-goto-root'));
        setRevision((value) => value + 1);
      });
      if (!offOverlay && engine?.bus) offOverlay = engine.bus.on('pptx-show-overlay', () => {
        setOverlayMount(host.querySelector<HTMLElement>('#react-show-overlay-root'));
        setRevision((value) => value + 1);
      });
      const target = host.querySelector<HTMLElement>('.show .popbar');
      if (target !== slot) {
        slot?.classList.remove('vibeoffice-react-owned');
        slot = target;
        slot?.classList.add('vibeoffice-react-owned');
        setMount(slot);
      }
      const timer = host.querySelector<HTMLElement>('#react-rehearse-root');
      if (timer !== timerSlot) {
        timerSlot?.classList.remove('vibeoffice-react-owned');
        timerSlot = timer;
        timerSlot?.classList.add('vibeoffice-react-owned');
        setRehearseMount(timerSlot);
      }
      const goto = host.querySelector<HTMLElement>('#react-goto-root');
      if (goto !== gotoSlot) {
        gotoSlot?.classList.remove('vibeoffice-react-owned');
        gotoSlot = goto;
        gotoSlot?.classList.add('vibeoffice-react-owned');
        setGotoMount(gotoSlot);
      }
      const overlay = host.querySelector<HTMLElement>('#react-show-overlay-root');
      if (overlay !== overlaySlot) {
        overlaySlot?.classList.remove('vibeoffice-react-owned');
        overlaySlot = overlay;
        overlaySlot?.classList.add('vibeoffice-react-owned');
        setOverlayMount(overlaySlot);
      }
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer.disconnect();
      offRehearse?.();
      offGoto?.();
      offOverlay?.();
      slot?.classList.remove('vibeoffice-react-owned');
      timerSlot?.classList.remove('vibeoffice-react-owned');
      gotoSlot?.classList.remove('vibeoffice-react-owned');
      overlaySlot?.classList.remove('vibeoffice-react-owned');
      const show = runtime()?.show;
      if (show) show.reactMounted = false;
    };
  }, [host]);
  const L = runtime();
  const show = L?.show;
  const rehearsal = show?.rehearseCatalog?.();
  const gotoBuffer = show?.gotoCatalog?.();
  const overlay = show?.overlayCatalog?.();
  void revision;
  if (!show) return null;
  const control = (name: string, icon: string, action: (button: HTMLButtonElement) => void) => <button type="button" aria-label={name} dangerouslySetInnerHTML={{ __html: L.icons?.get(icon) || '' }} onClick={(event) => { event.stopPropagation(); action(event.currentTarget); }} />;
  return <>
    {mount && createPortal(<>
      {control('Previous', 'showPrev', () => show.prev())}
      {control('Pointer options', 'showPen', (button) => show.openPointerMenu(button))}
      {control('Slide show menu', 'showMenu', (button) => show.openShowMenu(button))}
      {control('Next', 'showNext', () => show.next())}
    </>, mount)}
    {rehearseMount && rehearsal && createPortal(<>
      <button className="tb-btn" type="button" aria-label="Next" dangerouslySetInnerHTML={{ __html: L.icons?.get('forward') || '' }} onClick={(event) => { event.stopPropagation(); show.next(); }} />
      <button className="tb-btn" type="button" aria-label={rehearsal.paused ? 'Resume' : 'Pause'} onClick={(event) => { event.stopPropagation(); show.toggleRehearsePause?.(); }}>{rehearsal.paused ? '▶' : '❚❚'}</button>
      <span className="rt">{rehearsal.current}</span>
      <button className="tb-btn" type="button" aria-label="Repeat" dangerouslySetInnerHTML={{ __html: L.icons?.get('undo') || '' }} onClick={(event) => { event.stopPropagation(); show.repeatRehearse?.(); }} />
      <span className="tot">{rehearsal.total}</span>
    </>, rehearseMount)}
    {gotoMount && gotoBuffer != null && createPortal(L.t?.('Go to slide {0}', gotoBuffer) || `Go to slide ${gotoBuffer}`, gotoMount)}
    {overlayMount && overlay?.blank && createPortal(<div className="blank" style={{ background: overlay.blank }} aria-label={overlay.blank === '#000' ? 'Black screen' : 'White screen'} />, overlayMount)}
    {overlayMount && overlay?.ended && createPortal(<div className="endscreen">End of slide show, click to exit.</div>, overlayMount)}
  </>;
}
