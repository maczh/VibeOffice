import { useEffect, useRef, type HTMLAttributes, type ReactNode } from 'react';
import { useVibeOffice } from './react-bridge.js';
import './editor.css';

export type EditorKind = 'wordx' | 'excelx' | 'pptx';

export interface EditorFrameProps extends Omit<HTMLAttributes<HTMLElement>, 'children'> {
  appName: string;
  editor: EditorKind;
  children?: ReactNode;
}

/** A mounted document engine owns its DOM subtree and returns a disposer for listeners/resources. */
export interface EditorMountContext {
  host: HTMLElement;
  editor: EditorKind;
  signal: AbortSignal;
}

export type EditorMount = (context: EditorMountContext) => void | (() => void) | Promise<void | (() => void)>;

export interface LegacyEngineFrameOptions {
  url: string | URL;
  title: string;
  launch?: 'new' | 'sample';
}

/** Mount an existing same-origin app page in React-host mode; the page will not consult legacy Recent Files storage. */
export function createLegacyEngineFrame({ url, title, launch = 'new' }: LegacyEngineFrameOptions): EditorMount {
  const source = new URL(url, typeof window === 'undefined' ? 'http://localhost/' : window.location.href);
  if (!['http:', 'https:'].includes(source.protocol)) throw new TypeError('The editor runtime URL must use HTTP or HTTPS');
  source.searchParams.delete('open');
  source.searchParams.delete('recent');
  source.searchParams.delete('draft');
  source.searchParams.set('reactHost', '');
  source.searchParams.set(launch, '');

  return ({ host }) => {
    const frame = host.ownerDocument.createElement('iframe');
    frame.title = title;
    frame.src = source.href;
    frame.setAttribute('loading', 'eager');
    frame.style.cssText = 'display:block;width:100%;height:100%;min-height:600px;border:0';
    host.replaceChildren(frame);
    return () => frame.remove();
  };
}

export interface EngineEditorFrameProps extends EditorFrameProps {
  mount?: EditorMount;
  onMountError?: (error: unknown) => void;
}

/** Settle async mounts without leaking resources or producing unhandled rejections after unmount. */
export function settleEditorMount(
  mounted: void | (() => void) | Promise<void | (() => void)>,
  isDisposed: () => boolean,
  setCleanup: (cleanup: void | (() => void)) => void,
  onError: (error: unknown) => void,
): void {
  Promise.resolve(mounted).then((cleanup) => {
    if (isDisposed()) { if (typeof cleanup === 'function') cleanup(); }
    else setCleanup(cleanup);
  }).catch((error: unknown) => { if (!isDisposed()) onError(error); });
}

/** Bridge for legacy engines that manage their own DOM, without letting them replace React's outer host. */
export function EngineEditorFrame({ mount, onMountError, children, ...props }: EngineEditorFrameProps) {
  const hostRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const host = hostRef.current;
    if (!host || !mount) return;
    let disposed = false;
    let dispose: void | (() => void);
    const controller = new AbortController();
    const engineWindow = host.ownerDocument.defaultView as (Window & { L?: { createMountScope?: (element: HTMLElement) => { dispose(): void }; mountDomRoot?: (element: HTMLElement) => () => void; ui?: { mount?: (element: HTMLElement) => void | (() => void); unmount?: (element: HTMLElement) => void } } }) | null;
    const mountScope = engineWindow?.L?.createMountScope?.(host);
    const uiCleanup = mountScope ? null : engineWindow?.L?.ui?.mount?.(host);
    const restoreDomRoot = mountScope ? null : engineWindow?.L?.mountDomRoot?.(host);
    let mounted: void | (() => void) | Promise<void | (() => void)>;
    try { mounted = mount({ host, editor: props.editor, signal: controller.signal }); }
    catch (error) { controller.abort(); mountScope?.dispose(); restoreDomRoot?.(); uiCleanup?.(); throw error; }
    restoreDomRoot?.();
    settleEditorMount(mounted, () => disposed, (cleanup) => { dispose = cleanup; }, (error) => {
      if (onMountError) onMountError(error);
      else console.error('VibeOffice editor mount failed:', error);
    });
    return () => {
      disposed = true;
      controller.abort();
      if (dispose) dispose();
      mountScope?.dispose();
      if (typeof uiCleanup === 'function') uiCleanup();
      else engineWindow?.L?.ui?.unmount?.(host);
      host.replaceChildren();
    };
  }, [mount, onMountError, props.editor]);

  return <EditorFrame {...props}><EditorChrome appName={props.appName} /><div className="vibe-editor__engine" ref={hostRef}>{children}</div></EditorFrame>;
}

function EditorChrome({ appName }: { appName: string }) {
  const office = useVibeOffice();
  const sendToEngine = (element: HTMLElement, message: Record<string, unknown>) => {
    const frame = element.closest('.vibe-editor')?.querySelector('iframe');
    if (frame) {
      const origin = new URL(frame.src, window.location.href).origin;
      if (origin !== window.location.origin) throw new TypeError('The editor runtime must be served from the host origin');
      const engine = frame.contentWindow as (Window & {
        VO?: { look?: { set(id: string): void }; lang?: { set(id: string): void } };
        L?: { app?: { newDocument?: () => unknown; newWorkbook?: () => unknown; newPresentation?: () => unknown }; ui?: { cmds?: Record<string, unknown>; exec?(id: string): unknown } };
      }) | null;
      if (message.type === 'vibeoffice-react-command' && message.command === 'new' && engine?.L?.app) {
        const appId = element.closest('[data-vibe-editor]')?.getAttribute('data-vibe-editor');
        const createBlank = appId === 'wordx' ? engine.L.app.newDocument : appId === 'excelx' ? engine.L.app.newWorkbook : appId === 'pptx' ? engine.L.app.newPresentation : undefined;
        if (typeof createBlank === 'function') createBlank.call(engine.L.app);
        else frame.contentWindow?.postMessage(message, origin);
      } else if (message.type === 'vibeoffice-react-command' && typeof message.command === 'string' && engine?.L?.ui?.cmds?.[message.command] && engine.L.ui.exec) {
        engine.L.ui.exec(message.command);
      } else if (message.type === 'vibeoffice-react-chrome' && typeof message.look === 'string' && engine?.VO?.look) {
        engine.VO.look.set(message.look);
      } else if (message.type === 'vibeoffice-react-chrome' && typeof message.language === 'string' && engine?.VO?.lang) {
        engine.VO.lang.set(message.language);
      } else frame.contentWindow?.postMessage(message, origin);
    } else {
      const host = element.closest('.vibe-editor')?.querySelector('.vibe-editor__engine');
      host?.dispatchEvent(new CustomEvent(message.type === 'vibeoffice-react-command' ? 'vibeoffice-react-command' : 'vibeoffice-react-chrome', { detail: message, bubbles: false }));
    }
  };
  const command = (name: string, element: HTMLElement) => sendToEngine(element, { type: 'vibeoffice-react-command', command: name });
  const setLook = (look: typeof office.look, element: HTMLElement) => { sendToEngine(element, { type: 'vibeoffice-react-chrome', look }); office.setLook(look); };
  const setLanguage = (language: typeof office.language, element: HTMLElement) => { sendToEngine(element, { type: 'vibeoffice-react-chrome', language }); office.setLanguage(language); };
  return <nav className="vibe-editor__chrome" aria-label={`${appName} commands`}>
    <strong>{appName}</strong>
    <button type="button" onClick={(event) => command('new', event.currentTarget)}>{office.translate('New')}</button>
    <button type="button" onClick={(event) => command('open', event.currentTarget)}>{office.translate('Open')}</button>
    <button type="button" onClick={(event) => command('save', event.currentTarget)}>{office.translate('Save')}</button>
    <span className="vibe-editor__chrome-spacer" />
    <label>{office.translate('Look')} <select aria-label={office.translate('Look')} value={office.look} onChange={(event) => setLook(event.currentTarget.value as typeof office.look, event.currentTarget)}>{office.looks.map((look) => <option key={look.id} value={look.id}>{office.translate(look.name)}</option>)}</select></label>
    <label>{office.translate('Language')} <select aria-label={office.translate('Language')} value={office.language} onChange={(event) => setLanguage(event.currentTarget.value as typeof office.language, event.currentTarget)}>{office.languages.map((language) => <option key={language.id} value={language.id}>{office.translate(language.name)}</option>)}</select></label>
  </nav>;
}

export function EditorFrame({ appName, editor, className, children, ...attributes }: EditorFrameProps) {
  const classes = ['vibe-editor', className].filter(Boolean).join(' ');

  return (
    <section
      {...attributes}
      aria-label={`${appName} editor`}
      className={classes}
      data-vibe-editor={editor}
    >
      {children}
    </section>
  );
}
