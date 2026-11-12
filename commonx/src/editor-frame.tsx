import { useCallback, useEffect, useRef, useState, type FormEvent, type HTMLAttributes, type ReactNode } from 'react';
import { AboutDialog, type AboutRequest } from './about-dialog.js';
import { MessageDialog, type MessageRequest } from './message-dialog.js';
import { Toast, type ToastRequest } from './toast.js';
import { BusyOverlay, type BusyRequest } from './busy-overlay.js';
import { EditorTooltip } from './tooltip.js';
import { PromptDialog, type PromptRequest } from './prompt-dialog.js';
import { PasswordDialog, type PasswordRequest } from './password-dialog.js';
import { LinkDialog, type LinkRequest } from './link-dialog.js';
import { CompatibilityDialog, type CompatibilityRequest } from './compatibility-dialog.js';
import { ColorMenu, type ColorMenuRequest } from './color-menu.js';
import { MoreColorsDialog, type MoreColorsRequest } from './more-colors-dialog.js';
import { MenuLayer } from './menu-layer.js';
import { useVibeOffice } from './react-bridge.js';
import { CommandSearch } from './command-search.js';
import { EditorRibbon } from './ribbon.js';
import { EditorToolbarLayer } from './toolbar.js';
import { EditorStatusMessage } from './status-message.js';
import { EditorComboLayer } from './combo.js';
import { EditorSelectLayer } from './select.js';
import { EditorSpinLayer } from './spin.js';
import { EditorChoiceLayer } from './choice.js';
import { EditorTabsLayer } from './tabs.js';
import { EditorButtonLayer } from './button.js';
import { EditorUiContentLayer } from './ui-content.js';
import { OfficeDialog } from './dialog.js';
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

export interface EngineRuntimeOptions {
  url: string | URL;
  launch?: 'new' | 'sample';
}

export interface EngineEditorFrameProps extends EditorFrameProps {
  mount?: EditorMount;
  onMountError?: (error: unknown) => void;
  /** App-owned React layers that consume requests from the mounted engine. */
  appLayers?: (host: HTMLElement) => ReactNode;
}

export interface AppearanceControlsProps {
  onLookChange?: (look: ReturnType<typeof useVibeOffice>['look']) => void;
  onLanguageChange?: (language: ReturnType<typeof useVibeOffice>['language']) => void;
}

/** Shared look and language selectors used by the Start Center and every editor package. */
export function AppearanceControls({ onLookChange, onLanguageChange }: AppearanceControlsProps) {
  const office = useVibeOffice();
  return <>
    <label>{office.translate('Look')} <select aria-label={office.translate('Look')} value={office.look} onChange={(event) => { const look = event.currentTarget.value as typeof office.look; office.setLook(look); onLookChange?.(look); }}>{office.looks.map((look) => <option key={look.id} value={look.id}>{office.translate(look.name)}</option>)}</select></label>
    <label>{office.translate('Language')} <select aria-label={office.translate('Language')} value={office.language} onChange={(event) => { const language = event.currentTarget.value as typeof office.language; office.setLanguage(language); onLanguageChange?.(language); }}>{office.languages.map((language) => <option key={language.id} value={language.id}>{office.translate(language.name)}</option>)}</select></label>
  </>;
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
export function EngineEditorFrame({ mount, onMountError, appLayers, children, ...props }: EngineEditorFrameProps) {
  const hostRef = useRef<HTMLDivElement>(null);
  const [engineHost, setEngineHost] = useState<HTMLDivElement | null>(null);
  const bindHost = useCallback((host: HTMLDivElement | null) => {
    hostRef.current = host;
    setEngineHost(host);
  }, []);

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

  return <EditorFrame {...props}><EditorChrome appName={props.appName} /><div className="vibe-editor__engine" ref={bindHost}>{children}</div>{engineHost && <><EditorTooltip host={engineHost} /><MenuLayer host={engineHost} /><EditorRibbon host={engineHost} /><EditorToolbarLayer host={engineHost} /><EditorStatusMessage host={engineHost} /><EditorComboLayer host={engineHost} /><EditorSelectLayer host={engineHost} /><EditorSpinLayer host={engineHost} /><EditorChoiceLayer host={engineHost} /><EditorTabsLayer host={engineHost} /><EditorButtonLayer host={engineHost} /><EditorUiContentLayer host={engineHost} />{appLayers?.(engineHost)}</>}</EditorFrame>;
}

function EditorChrome({ appName }: { appName: string }) {
  const office = useVibeOffice();
  const chromeRef = useRef<HTMLElement>(null);
  const [showUrl, setShowUrl] = useState(false);
  const [url, setUrl] = useState('');
  const [urlError, setUrlError] = useState('');
  const [urlPending, setUrlPending] = useState(false);
  const [about, setAbout] = useState<AboutRequest | null>(null);
  const [message, setMessage] = useState<MessageRequest | null>(null);
  const [toast, setToast] = useState<ToastRequest | null>(null);
  const [busy, setBusy] = useState<BusyRequest | null>(null);
  const [prompt, setPrompt] = useState<PromptRequest | null>(null);
  const [password, setPassword] = useState<PasswordRequest | null>(null);
  const [link, setLink] = useState<LinkRequest | null>(null);
  const [compatibility, setCompatibility] = useState<CompatibilityRequest | null>(null);
  const [colorMenu, setColorMenu] = useState<ColorMenuRequest | null>(null);
  const [moreColors, setMoreColors] = useState<MoreColorsRequest | null>(null);
  const closeAbout = () => setAbout(null);
  const closeMessage = () => setMessage(null);
  const dismissToast = useCallback(() => setToast(null), []);
  useEffect(() => {
    const engine = chromeRef.current?.closest('.vibe-editor')?.querySelector('.vibe-editor__engine');
    if (!engine) return;
    const onAbout = (event: Event) => {
      const request = (event as CustomEvent<AboutRequest>).detail;
      if (request && typeof request.name === 'string' && Array.isArray(request.text)) setAbout(request);
    };
    const onMessage = (event: Event) => {
      const request = (event as CustomEvent<MessageRequest>).detail;
      if (request && typeof request.text === 'string' && Array.isArray(request.buttons) && typeof request.respond === 'function') setMessage(request);
    };
    const onToast = (event: Event) => {
      const request = (event as CustomEvent<ToastRequest>).detail;
      if (request && typeof request.text === 'string' && typeof request.duration === 'number') setToast(request);
    };
    const onBusy = (event: Event) => {
      const request = (event as CustomEvent<BusyRequest | null>).detail;
      setBusy(request && typeof request.text === 'string' ? request : null);
    };
    const onPrompt = (event: Event) => {
      const request = (event as CustomEvent<PromptRequest>).detail;
      if (request && typeof request.title === 'string' && typeof request.label === 'string' && typeof request.value === 'string' && typeof request.respond === 'function') setPrompt(request);
    };
    const onPassword = (event: Event) => {
      const request = (event as CustomEvent<PasswordRequest>).detail;
      if (request && typeof request.text === 'string' && typeof request.wrong === 'boolean' && typeof request.respond === 'function') setPassword(request);
    };
    const onLink = (event: Event) => {
      const request = (event as CustomEvent<LinkRequest>).detail;
      if (request && typeof request.url === 'string' && typeof request.mail === 'boolean' && typeof request.respond === 'function') setLink(request);
    };
    const onCompatibility = (event: Event) => {
      const request = (event as CustomEvent<CompatibilityRequest>).detail;
      if (request && Array.isArray(request.groups) && typeof request.respond === 'function') setCompatibility(request);
    };
    const onColorMenu = (event: Event) => {
      const request = (event as CustomEvent<ColorMenuRequest>).detail;
      if (request && Array.isArray(request.sections) && typeof request.choose === 'function' && typeof request.more === 'function') setColorMenu(request);
    };
    const onMoreColors = (event: Event) => {
      const request = (event as CustomEvent<MoreColorsRequest>).detail;
      if (request && typeof request.initial === 'string' && Array.isArray(request.rows) && Array.isArray(request.grays) && typeof request.respond === 'function') setMoreColors(request);
    };
    engine.addEventListener('vibeoffice-react-about', onAbout);
    engine.addEventListener('vibeoffice-react-message', onMessage);
    engine.addEventListener('vibeoffice-react-toast', onToast);
    engine.addEventListener('vibeoffice-react-busy', onBusy);
    engine.addEventListener('vibeoffice-react-prompt', onPrompt);
    engine.addEventListener('vibeoffice-react-password', onPassword);
    engine.addEventListener('vibeoffice-react-link', onLink);
    engine.addEventListener('vibeoffice-react-compatibility', onCompatibility);
    engine.addEventListener('vibeoffice-react-color-menu', onColorMenu);
    engine.addEventListener('vibeoffice-react-more-colors', onMoreColors);
    return () => {
      engine.removeEventListener('vibeoffice-react-about', onAbout);
      engine.removeEventListener('vibeoffice-react-message', onMessage);
      engine.removeEventListener('vibeoffice-react-toast', onToast);
      engine.removeEventListener('vibeoffice-react-busy', onBusy);
      engine.removeEventListener('vibeoffice-react-prompt', onPrompt);
      engine.removeEventListener('vibeoffice-react-password', onPassword);
      engine.removeEventListener('vibeoffice-react-link', onLink);
      engine.removeEventListener('vibeoffice-react-compatibility', onCompatibility);
      engine.removeEventListener('vibeoffice-react-color-menu', onColorMenu);
      engine.removeEventListener('vibeoffice-react-more-colors', onMoreColors);
    };
  }, []);
  const sendToEngine = (element: HTMLElement, message: Record<string, unknown>) => {
    const host = element.closest('.vibe-editor')?.querySelector('.vibe-editor__engine');
    host?.dispatchEvent(new CustomEvent(message.type === 'vibeoffice-react-command' ? 'vibeoffice-react-command' : 'vibeoffice-react-chrome', { detail: message, bubbles: false }));
  };
  const command = (name: string, element: HTMLElement) => sendToEngine(element, { type: 'vibeoffice-react-command', command: name });
  const appearanceChanged = (detail: { look?: typeof office.look; language?: typeof office.language }) => {
    const host = chromeRef.current?.closest('.vibe-editor')?.querySelector<HTMLElement>('.vibe-editor__engine');
    host?.dispatchEvent(new CustomEvent('vibeoffice-react-chrome', { detail, bubbles: false }));
  };
  const submitUrl = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const host = event.currentTarget.closest('.vibe-editor')?.querySelector('.vibe-editor__engine');
    if (!host) { setUrlPending(false); return; }
    setUrlError('');
    setUrlPending(true);
    host.dispatchEvent(new CustomEvent('vibeoffice-react-open-url', {
      detail: {
        url,
        respond: (error?: string) => {
          setUrlPending(false);
          if (error) setUrlError(error);
          else { setShowUrl(false); setUrl(''); }
        },
      },
      bubbles: false,
    }));
  };
  return <>
    <nav ref={chromeRef} className="vibe-editor__chrome" aria-label={`${appName} commands`}>
    <strong>{appName}</strong>
    <button type="button" onClick={(event) => command('new', event.currentTarget)}>{office.translate('New')}</button>
    <button type="button" onClick={(event) => command('open', event.currentTarget)}>{office.translate('Open')}</button>
    <button type="button" aria-expanded={showUrl} onClick={() => setShowUrl((shown) => !shown)}>URL</button>
    {showUrl && <OfficeDialog title="URL" className="vibe-editor__url-dialog" onRequestClose={() => { setShowUrl(false); setUrlError(''); }}>
      <form className="vibe-editor__url-open" onSubmit={submitUrl}>
        <input type="url" aria-label="URL" required value={url} aria-busy={urlPending} onChange={(event) => { setUrl(event.currentTarget.value); setUrlError(''); }} />
        {urlError && <p className="vibe-editor__url-error" role="alert">{urlError}</p>}
        <button type="submit" disabled={urlPending}>{office.translate('Open')}</button>
        <button type="button" aria-label={office.translate('Cancel')} disabled={urlPending} onClick={() => setShowUrl(false)}>{office.translate('Cancel')}</button>
      </form>
    </OfficeDialog>}
    <button type="button" onClick={(event) => command('save', event.currentTarget)}>{office.translate('Save')}</button>
    <CommandSearch />
    <span className="vibe-editor__chrome-spacer" />
    <AppearanceControls onLookChange={(look) => appearanceChanged({ look })} onLanguageChange={(language) => appearanceChanged({ language })} />
    </nav>
    {about && <AboutDialog request={about} onClose={closeAbout} />}
    {message && <MessageDialog request={message} onClose={closeMessage} />}
    {toast && <Toast request={toast} onDismiss={dismissToast} />}
    {busy && <BusyOverlay request={busy} />}
    {prompt && <PromptDialog request={prompt} onClose={() => setPrompt(null)} />}
    {password && <PasswordDialog request={password} onClose={() => setPassword(null)} />}
    {link && <LinkDialog request={link} onClose={() => setLink(null)} />}
    {compatibility && <CompatibilityDialog request={compatibility} onClose={() => setCompatibility(null)} />}
    {colorMenu && <ColorMenu request={colorMenu} onClose={() => setColorMenu(null)} />}
    {moreColors && <MoreColorsDialog request={moreColors} onClose={() => setMoreColors(null)} />}
  </>;
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
