import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent } from 'react';
import { createRoot } from 'react-dom/client';
import { initializeSharedRuntime } from '@vibeoffice/core';
import { AboutDialog, AppearanceFields, createUrlReadProvider, createWebDavProvider, OfficeDialog, saveLocalDocument, VibeOfficeProvider, useVibeOffice } from '@vibeoffice/core';
import { createIndexedDbDraftStore } from '@vibeoffice/core';
import type { LocalDocument, LocalDraft, RemoteFile, RemoteStorageProvider, SaveFileOperation } from '@vibeoffice/core';
import { WordxEditor } from '@vibeoffice/wordx';
import { ExcelxEditor } from '@vibeoffice/excelx';
import { PptxEditor } from '@vibeoffice/pptx';
import '@vibeoffice/wordx/style.css';
import '@vibeoffice/excelx/style.css';
import '@vibeoffice/pptx/style.css';
import './style.css';

initializeSharedRuntime();
const mark = (message: string) => window.VO?.N?.(message) ?? message;

type AppId = 'wordx' | 'excelx' | 'pptx';
type Entry = { id: AppId; label: string; description: string; folder: string; extensions: string[] };
type Template = { id: string; name: string; img?: string | null };
type InstallPromptEvent = Event & { prompt: () => Promise<void>; userChoice: Promise<{ outcome: 'accepted' | 'dismissed'; platform?: string }> };
const apps: Entry[] = [
  { id: 'wordx', label: 'Wordx', description: 'Word processor', folder: 'wordx', extensions: ['docx', 'docm', 'dotx', 'dotm', 'rtf', 'txt', 'md', 'markdown', 'html', 'htm', 'zip'] },
  { id: 'excelx', label: 'Excelx', description: 'Spreadsheet', folder: 'excelx', extensions: ['xlsx', 'xlsm', 'xltx', 'xltm', 'csv', 'tsv', 'txt', 'xml', 'html', 'htm'] },
  { id: 'pptx', label: 'PPTX', description: 'Presentations', folder: 'pptx', extensions: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm'] },
];
const UI_TEXT = {
  openUrl: mark('Open URL…'), webdav: mark('WebDAV…'), changed: mark('The remote file changed since it was opened. Replace it anyway?'),
  unsupported: mark('This file type is not supported.'), openError: mark('The file could not be opened: {0}'), davUrl: mark('WebDAV collection URL'),
  saveAs: mark('Save As'), fileName: mark('File name'), save: mark('Save'), replace: mark('&Replace It'),
  empty: mark('This WebDAV folder is empty.'), davFiles: mark('WebDAV files'), choose: mark('Enter a number, or Cancel to stop.'),
  username: mark('Username'), password: mark('Password'), continue: mark('Continue'), cancel: mark('Cancel'),
  recoverTitle: mark('Recover unsaved work'), recover: mark('Recover'), dismiss: mark('Dismiss'), savedAt: mark('Saved {0}'),
};
const draftStore = createIndexedDbDraftStore<Blob>();
function appForName(name: string): Entry | undefined {
  const extension = name.split('.').pop()?.toLowerCase();
  return apps.find((candidate) => candidate.extensions.includes(extension || ''));
}

function AppIcon({ app }: { app: AppId }) {
  const color = 'var(--p-accent)';
  return <svg className={`suite-app-icon ${app}`} viewBox="0 0 80 100" width="24" height="30" aria-hidden="true">
    <path d="M4 2h54l18 18v76H4z" fill="var(--surface)" stroke="var(--field-border)" />
    <path d="M58 2v18h18" fill="var(--p-accent-soft)" stroke="var(--field-border)" />
    {app === 'excelx'
      ? <><rect x="14" y="22" width="52" height="54" fill="var(--surface)" stroke={color} /><path d="M14 31h52M14 40h52M14 49h52M14 58h52M14 67h52M30 22v54M48 22v54" stroke={color} strokeOpacity=".55" /><rect x="14" y="22" width="52" height="9" fill="var(--p-accent-mid)" /></>
      : app === 'pptx'
        ? <><rect x="12" y="28" width="56" height="40" rx="2" fill="var(--p-accent-soft)" stroke={color} /><rect x="18" y="34" width="44" height="8" fill={color} /><path d="M20 50h40M20 57h30" stroke={color} strokeWidth="3" strokeLinecap="round" strokeOpacity=".6" /></>
        : <path d="M18 26h44M18 34h44M18 42h44M18 50h44M18 58h30" stroke={color} strokeWidth="3" strokeLinecap="round" />}
  </svg>;
}

function StartActionIcon({ action }: { action: 'open' | 'create' }) {
  return action === 'open'
    ? <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true"><path d="M1.5 5.5v11h14l3-8h-14l-3 8" fill="var(--p-accent-mid)" stroke="var(--p-accent-ink)" /><path d="M1.5 16V3.5h5l1.5 2h6.5v3" fill="var(--p-accent-soft)" stroke="var(--p-accent-ink)" /></svg>
    : <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true"><path d="M3.5 1.5h9l4 4v13h-13z" fill="var(--surface)" stroke="var(--field-border)" /><path d="M12.5 1.5v4h4" fill="var(--p-accent-soft)" stroke="var(--field-border)" /><circle cx="14" cy="14.5" r="4.5" fill="var(--p-accent)" /><path d="M14 12v5M11.5 14.5h5" stroke="var(--surface)" strokeWidth="1.8" /></svg>;
}

function StartFooterIcon({ kind }: { kind: 'install' | 'feedback' | 'settings' | 'themes' | 'about' }) {
  if (kind === 'install') return <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true"><rect x="2" y="3" width="16" height="12" rx="1.5" fill="var(--surface)" stroke="var(--field-border)" /><path d="M10 5.5v6M7 9l3 3 3-3" fill="none" stroke="var(--p-accent)" strokeWidth="1.8" /><path d="M6 17.5h8" stroke="var(--field-border)" /></svg>;
  if (kind === 'feedback') return <svg width="20" height="20" viewBox="0 0 16 16" aria-hidden="true"><path fill="var(--pane-ink)" d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z" /></svg>;
  if (kind === 'settings') return <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true"><circle cx="8" cy="8" r="5.6" fill="none" stroke="var(--muted)" strokeWidth="2.6" strokeDasharray="2.2 2.2" /><circle cx="8" cy="8" r="4" fill="var(--p-accent-soft)" stroke="var(--field-border)" strokeWidth="1.2" /><circle cx="8" cy="8" r="1.6" fill="var(--surface)" stroke="var(--field-border)" /></svg>;
  if (kind === 'themes') return <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true"><path d="M8 1.5a6.5 6.5 0 1 0 0 13c1 0 1.4-.8 1-1.6-.5-1 .1-2 1.2-2h1.6a2.7 2.7 0 0 0 2.7-2.9C14.3 4.3 11.5 1.5 8 1.5z" fill="var(--surface)" stroke="var(--field-border)" /><circle cx="5" cy="6.5" r="1.3" fill="var(--p-accent)" /><circle cx="8.2" cy="4.6" r="1.3" fill="var(--p-accent-mid)" /><circle cx="11.3" cy="6.4" r="1.3" fill="var(--p-accent-soft)" /><circle cx="5.2" cy="10" r="1.3" fill="var(--p-accent-ink)" /></svg>;
  return <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true"><circle cx="8" cy="8" r="6.8" fill="var(--p-accent)" /><path d="M8 7v4.6" stroke="var(--surface)" strokeWidth="1.9" /><circle cx="8" cy="4.7" r="1.15" fill="var(--surface)" /></svg>;
}

function requestEditorClose(): Promise<boolean> {
  if (!document.querySelector('.suite-frame .vibe-editor__engine > *')) return Promise.resolve(true);
  return window.VO?.reactClose ? window.VO.reactClose().catch(() => false) : Promise.resolve(true);
}

function ActiveEditor({ app, file, document, template, allowedRemoteOrigins, onSaveFile, onDraftsAvailable, onTitleChange }: { app: Entry; file?: File; document?: LocalDocument; template?: string; allowedRemoteOrigins: readonly string[]; onSaveFile: (name: string, blob: Blob, current?: LocalDocument, signal?: AbortSignal) => Promise<LocalDocument | void | null>; onDraftsAvailable: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void; onTitleChange: (title: string) => void }) {
  const query = template ? `template=${encodeURIComponent(template)}` : 'new';
  const key = `${app.id}:${file?.name || template || 'blank'}`;
  const engine = useMemo(() => ({ url: `/${app.folder}/?${query}` }), [app.folder, query]);
  const recovery = { draftStore, onDraftsAvailable };
  if (app.id === 'wordx') return <WordxEditor key={key} className="suite-frame" engine={engine} allowedRemoteOrigins={allowedRemoteOrigins} onSaveFile={onSaveFile} onTitleChange={onTitleChange} {...recovery} {...(document ? { initialDocument: document } : file ? { initialFile: file } : {})} />;
  if (app.id === 'excelx') return <ExcelxEditor key={key} className="suite-frame" engine={engine} allowedRemoteOrigins={allowedRemoteOrigins} onSaveFile={onSaveFile} onTitleChange={onTitleChange} {...recovery} {...(document ? { initialDocument: document } : file ? { initialFile: file } : {})} />;
  return <PptxEditor key={key} className="suite-frame" engine={engine} allowedRemoteOrigins={allowedRemoteOrigins} onSaveFile={onSaveFile} onTitleChange={onTitleChange} {...recovery} {...(document ? { initialDocument: document } : file ? { initialFile: file } : {})} />;
}

function Suite() {
  const office = useVibeOffice();
  const { translate } = office;
  const [active, setActive] = useState<{ app: Entry; file?: File; document?: LocalDocument; template?: string } | null>(null);
  const initialTitle = useRef(document.title);
  const [recoverable, setRecoverable] = useState<LocalDraft<Blob>[]>([]);
  const restoreDraft = useRef<((draft: LocalDraft<Blob>) => void) | null>(null);
  const onDraftsAvailable = useCallback((drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => {
    restoreDraft.current = restore;
    setRecoverable(drafts);
  }, []);
  const onTitleChange = useCallback((title: string) => { document.title = title; }, []);
  useEffect(() => { if (!active) document.title = initialTitle.current; }, [active]);
  const [templates, setTemplates] = useState<Record<string, Template[]>>({});
  const [canInstall, setCanInstall] = useState(false);
  const installPrompt = useRef<InstallPromptEvent | null>(null);
  const [aboutOpen, setAboutOpen] = useState(false);
  const [appearanceOpen, setAppearanceOpen] = useState<'settings' | 'themes' | null>(null);
  const picker = useRef<HTMLInputElement>(null);
  const [credentialAsk, setCredentialAsk] = useState<{ label: string; secret: boolean; finish: (value: string | null) => void } | null>(null);
  const [credentialInput, setCredentialInput] = useState('');
  const [saveAsAsk, setSaveAsAsk] = useState<{ finish: (value: string | null) => void } | null>(null);
  const [saveAsName, setSaveAsName] = useState('');
  const [conflictAsk, setConflictAsk] = useState<{ finish: (replace: boolean) => void } | null>(null);
  const [urlOpen, setUrlOpen] = useState(false);
  const [urlValue, setUrlValue] = useState('');
  const [urlError, setUrlError] = useState('');
  const [urlBusy, setUrlBusy] = useState(false);
  const [webDavOpen, setWebDavOpen] = useState(false);
  const [webDavMode, setWebDavMode] = useState<'open' | 'saveAs'>('open');
  const [webDavUrl, setWebDavUrl] = useState('');
  const [webDavProvider, setWebDavProvider] = useState<RemoteStorageProvider | null>(null);
  const [webDavParent, setWebDavParent] = useState<string | undefined>();
  const [webDavTrail, setWebDavTrail] = useState<{ id?: string; name: string }[]>([]);
  const [webDavEntries, setWebDavEntries] = useState<RemoteFile[]>([]);
  const [webDavBusy, setWebDavBusy] = useState(false);
  const [webDavError, setWebDavError] = useState('');
  const remoteTarget = useRef<{ provider: RemoteStorageProvider; parentId?: string } | null>(null);
  const remoteSaveTarget = useRef<{ provider: RemoteStorageProvider; parentId?: string } | null>(null);
  const allowedOrigins = useMemo(() => [...new Set([window.location.origin, ...(document.querySelector('meta[name="vibeoffice-remote-origins"]')?.getAttribute('content') || '').split(/\s+/).filter(Boolean)])], []);

  useEffect(() => {
    void fetch('/common/templates.json').then((response) => response.json()).then(setTemplates).catch(() => setTemplates({}));
    if ('serviceWorker' in navigator && !navigator.webdriver && window.isSecureContext) {
      void navigator.serviceWorker.register('/sw.js', { updateViaCache: 'none' }).catch((error: unknown) => console.warn('Service worker:', error));
    }
  }, []);

  useEffect(() => {
    const supportedOrigin = (location.protocol === 'https:' && location.hostname === 'vibeoffice.work')
      || (location.protocol === 'http:' && ['localhost', '127.0.0.1'].includes(location.hostname));
    if (!supportedOrigin
      || matchMedia('(display-mode: standalone), (display-mode: window-controls-overlay)').matches) return;
    const onBeforeInstall = (event: Event) => {
      event.preventDefault();
      installPrompt.current = event as InstallPromptEvent;
      setCanInstall(true);
    };
    const onInstalled = () => { installPrompt.current = null; setCanInstall(false); };
    window.addEventListener('beforeinstallprompt', onBeforeInstall);
    window.addEventListener('appinstalled', onInstalled);
    return () => {
      window.removeEventListener('beforeinstallprompt', onBeforeInstall);
      window.removeEventListener('appinstalled', onInstalled);
    };
  }, []);

  useEffect(() => {
    const shortcutApp = apps.find((app) => app.id === new URLSearchParams(window.location.search).get('new'));
    if (!shortcutApp) return;
    window.history.replaceState(window.history.state, '', window.location.pathname + window.location.hash);
    setActive({ app: shortcutApp });
  }, []);

  const openFile = (file?: File) => {
    if (!file) return;
    const ext = file.name.split('.').pop()?.toLowerCase();
    const app = apps.find((candidate) => candidate.extensions.includes(ext || ''));
    if (app) setActive({ app, file });
  };
  const create = (app: Entry, template?: string) => setActive({ app, ...(template ? { template } : {}) });
  const askCredential = (label: string, secret: boolean) => new Promise<string | null>((resolve) => {
    setCredentialInput('');
    setCredentialAsk({ label, secret, finish: resolve });
  });
  const finishCredential = (value: string | null) => {
    credentialAsk?.finish(value);
    setCredentialAsk(null);
    setCredentialInput('');
  };
  const requestSaveAsName = useCallback((suggestedName: string, signal?: AbortSignal) => new Promise<string | null>((resolve) => {
    let settled = false;
    const finish = (value: string | null) => {
      if (settled) return;
      settled = true;
      signal?.removeEventListener('abort', onAbort);
      setSaveAsAsk(null);
      setSaveAsName('');
      resolve(value);
    };
    const onAbort = () => finish(null);
    if (signal?.aborted) { finish(null); return; }
    setSaveAsName(suggestedName);
    setSaveAsAsk({ finish });
    signal?.addEventListener('abort', onAbort, { once: true });
  }), []);
  const requestConflictChoice = useCallback((signal?: AbortSignal) => new Promise<boolean>((resolve) => {
    let settled = false;
    const finish = (replace: boolean) => {
      if (settled) return;
      settled = true;
      signal?.removeEventListener('abort', onAbort);
      setConflictAsk(null);
      resolve(replace);
    };
    const onAbort = () => finish(false);
    if (signal?.aborted) { finish(false); return; }
    setConflictAsk({ finish });
    signal?.addEventListener('abort', onAbort, { once: true });
  }), []);
  const onSaveFile = useCallback(async (name: string, blob: Blob, current?: LocalDocument, signal?: AbortSignal, operation?: SaveFileOperation) => {
    if (operation?.saveAs) {
      remoteSaveTarget.current = null;
      const selectedName = (await requestSaveAsName(name, signal))?.trim();
      if (!selectedName) return null;
      const existingRemote = current?.remote;
      const target = remoteSaveTarget.current || (remoteTarget.current?.provider === existingRemote?.provider ? remoteTarget.current : existingRemote ? null : remoteTarget.current);
      const provider = target?.provider || (existingRemote?.provider.create ? existingRemote.provider : undefined);
      if (provider?.create) {
        const file = await provider.create(target?.parentId, selectedName, blob, signal);
        remoteSaveTarget.current = null;
        return { file: new File([blob], file.name || selectedName, { type: blob.type }), remote: { provider, file } };
      }
      return saveLocalDocument({ file: new File([], selectedName) }, blob, selectedName, window, signal);
    }
    const remote = current?.remote;
    if (remote?.provider.save) {
      try {
        const file = await remote.provider.save(remote.file.id, blob, { ...(remote.file.revision ? { ifRevision: remote.file.revision } : {}), ...(signal ? { signal } : {}) });
        return { file: new File([blob], file.name || name, { type: blob.type }), remote: { provider: remote.provider, file } };
      } catch (error) {
        if ((error as { code?: string })?.code !== 'revision-conflict') throw error;
        if (!(await requestConflictChoice(signal))) return null;
        const file = await remote.provider.save(remote.file.id, blob, signal ? { signal } : undefined);
        return { file: new File([blob], file.name || name, { type: blob.type }), remote: { provider: remote.provider, file } };
      }
    }
    if (!remote && remoteTarget.current?.provider.create) {
      const target = remoteTarget.current;
      const file = await target.provider.create(target.parentId, name, blob, signal);
      return { file: new File([blob], file.name || name, { type: blob.type }), remote: { provider: target.provider, file } };
    }
    return saveLocalDocument(current ?? { file: new File([], name) }, blob, name, window, signal);
  }, [requestSaveAsName, requestConflictChoice, translate]);

  const openUrl = () => { setUrlValue(''); setUrlError(''); setUrlOpen(true); };
  const submitUrl = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault(); setUrlBusy(true); setUrlError('');
    try {
      const reader = createUrlReadProvider(allowedOrigins);
      const { file, blob } = await reader.read(urlValue);
      const app = appForName(file.name);
      if (!app) throw new Error(translate(UI_TEXT.unsupported));
      const opened = new File([blob], file.name, { type: file.mimeType || blob.type });
      setUrlOpen(false);
      setActive({ app, document: { file: opened, remote: { provider: reader, file } } });
    } catch (error) { setUrlError(translate(UI_TEXT.openError, error instanceof Error ? error.message : String(error))); }
    finally { setUrlBusy(false); }
  };

  const openWebDav = (mode: 'open' | 'saveAs' = 'open') => {
    setWebDavMode(mode); setWebDavError(''); setWebDavOpen(true);
    if (mode === 'saveAs' && webDavProvider) {
      setWebDavTrail([{ name: webDavUrl || translate(UI_TEXT.davFiles) }]);
      void listWebDav(webDavProvider);
    } else {
      setWebDavUrl(''); setWebDavProvider(null); setWebDavParent(undefined); setWebDavTrail([]); setWebDavEntries([]);
    }
  };
  const listWebDav = async (provider: RemoteStorageProvider, parentId?: string) => {
    setWebDavBusy(true); setWebDavError('');
    try {
      const entries = (await provider.list(parentId)).filter((entry) => entry.name);
      setWebDavProvider(provider); setWebDavParent(parentId); setWebDavEntries(entries);
      if (!entries.length) setWebDavError(translate(UI_TEXT.empty));
    } catch (error) { setWebDavError(translate(UI_TEXT.openError, error instanceof Error ? error.message : String(error))); }
    finally { setWebDavBusy(false); }
  };
  const connectWebDav = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    try {
      let credentials: { username: string; password: string } | undefined;
      const provider = createWebDavProvider({ baseUrl: webDavUrl, allowedOrigins, credentials: async () => {
        if (credentials) return credentials;
        const username = await askCredential(translate(UI_TEXT.username), false);
        if (username === null) return undefined;
        const password = await askCredential(translate(UI_TEXT.password), true);
        if (password === null) return undefined;
        credentials = { username, password };
        return credentials;
      } });
      setWebDavTrail([{ name: webDavUrl }]);
      await listWebDav(provider);
    } catch (error) { setWebDavError(translate(UI_TEXT.openError, error instanceof Error ? error.message : String(error))); }
  };
  const chooseWebDavEntry = async (entry: RemoteFile) => {
    if (!webDavProvider) return;
    if (entry.isDirectory) {
      setWebDavTrail((trail) => [...trail, { id: entry.id, name: entry.name }]);
      await listWebDav(webDavProvider, entry.id);
      return;
    }
    setWebDavBusy(true); setWebDavError('');
    try {
      const { file, blob } = await webDavProvider.read(entry.id);
      const app = appForName(file.name);
      if (!app) throw new Error(translate(UI_TEXT.unsupported));
      const opened = new File([blob], file.name, { type: file.mimeType || blob.type });
      remoteTarget.current = { provider: webDavProvider, ...(webDavParent !== undefined ? { parentId: webDavParent } : {}) };
      setWebDavOpen(false);
      setActive({ app, document: { file: opened, remote: { provider: webDavProvider, file } } });
    } catch (error) { setWebDavError(translate(UI_TEXT.openError, error instanceof Error ? error.message : String(error))); }
    finally { setWebDavBusy(false); }
  };

  const webDavDialog = webDavOpen && <OfficeDialog title={translate(UI_TEXT.davFiles)} className="credential" onRequestClose={() => setWebDavOpen(false)}>
    <form onSubmit={(event) => void connectWebDav(event)}><label>{translate(UI_TEXT.davUrl)}<input autoFocus required type="url" value={webDavUrl} onChange={(event) => setWebDavUrl(event.currentTarget.value)} /></label><button type="submit" disabled={webDavBusy}>{translate(UI_TEXT.continue)}</button></form>
    {webDavProvider && <section aria-label={translate(UI_TEXT.davFiles)}><p>{webDavTrail.at(-1)?.name}</p>{webDavTrail.length > 1 && <button type="button" disabled={webDavBusy} onClick={() => { const trail = webDavTrail.slice(0, -1); setWebDavTrail(trail); void listWebDav(webDavProvider, trail.length > 1 ? trail.at(-1)?.id : undefined); }}>{translate('Back')}</button>}{webDavMode === 'saveAs' ? <button type="button" disabled={webDavBusy} onClick={() => { remoteSaveTarget.current = { provider: webDavProvider, ...(webDavParent !== undefined ? { parentId: webDavParent } : {}) }; setWebDavOpen(false); }}>{translate(UI_TEXT.saveAs)}</button> : <button type="button" disabled={webDavBusy} onClick={() => { remoteTarget.current = { provider: webDavProvider, ...(webDavParent !== undefined ? { parentId: webDavParent } : {}) }; setWebDavOpen(false); }}>{translate('Create File')}</button>}<ul>{webDavEntries.map((entry) => <li key={entry.id}><button type="button" disabled={webDavBusy || (webDavMode === 'saveAs' && !entry.isDirectory)} onClick={() => void chooseWebDavEntry(entry)}>{entry.name}{entry.isDirectory ? '/' : ''}</button></li>)}</ul></section>}
    {webDavError && <p role="alert">{webDavError}</p>}<button type="button" onClick={() => setWebDavOpen(false)}>{translate(UI_TEXT.cancel)}</button>
  </OfficeDialog>;

  if (active) {
    const { app, file, template } = active;
    return <main className="suite-editor">
      <button className="back" type="button" onClick={() => { void requestEditorClose().then((closed) => { if (closed) setActive(null); }); }}>{translate('Create File')}</button>
      {recoverable.length > 0 && <aside className="draft-recovery" aria-label={translate(UI_TEXT.recoverTitle)}><strong>{translate(UI_TEXT.recoverTitle)}</strong>{recoverable.map((draft) => <div key={draft.id}><span>{draft.name} · {translate(UI_TEXT.savedAt, new Date(draft.updatedAt).toLocaleString())}</span><button type="button" onClick={() => restoreDraft.current?.(draft)}>{translate(UI_TEXT.recover)}</button><button type="button" onClick={() => { void draftStore.remove(draft.id).then(() => setRecoverable((items) => items.filter((item) => item.id !== draft.id))); }}>{translate(UI_TEXT.dismiss)}</button></div>)}</aside>}
      <ActiveEditor app={app} allowedRemoteOrigins={allowedOrigins} onSaveFile={onSaveFile} onDraftsAvailable={onDraftsAvailable} onTitleChange={onTitleChange} {...(file ? { file } : {})} {...(active.document ? { document: active.document } : {})} {...(template ? { template } : {})} />
      {conflictAsk && <OfficeDialog title={translate(UI_TEXT.changed)} className="credential" onRequestClose={() => conflictAsk.finish(false)}><p>{translate(UI_TEXT.changed)}</p><button type="button" onClick={() => conflictAsk.finish(false)}>{translate(UI_TEXT.cancel)}</button><button type="button" onClick={() => conflictAsk.finish(true)}>{translate(UI_TEXT.replace)}</button></OfficeDialog>}
      {saveAsAsk && <OfficeDialog title={translate(UI_TEXT.saveAs)} className="credential" onRequestClose={() => saveAsAsk.finish(null)}><form onSubmit={(event) => { event.preventDefault(); saveAsAsk.finish(saveAsName); }}><label>{translate(UI_TEXT.fileName)}<input autoFocus required value={saveAsName} onChange={(event) => setSaveAsName(event.currentTarget.value)} /></label><button type="button" onClick={() => openWebDav('saveAs')}>{translate(UI_TEXT.webdav)}</button><button type="button" onClick={() => saveAsAsk.finish(null)}>{translate(UI_TEXT.cancel)}</button><button type="submit">{translate(UI_TEXT.save)}</button></form></OfficeDialog>}
      {webDavDialog}
    </main>;
  }

  return <div className="suite-start">
    <nav className="suite-side" aria-label={translate('Start')}>
      <div className="suite-brand"><img src="/icons/icon.svg" alt="" /><span><b translate="no">VibeOffice.work</b><small>{translate('Office suite in your browser')}</small></span></div>
      <div className="suite-actions">
        <button className="suite-action" type="button" onClick={() => picker.current?.click()}><StartActionIcon action="open" />{translate('Open File…')}</button>
        <button className="suite-action suite-action-on" type="button" aria-current="page"><StartActionIcon action="create" />{translate('Create File')}</button>
      </div>
      <section className="suite-app-links" aria-label={translate('Apps')}>
        <h2>{translate('Apps')}</h2>
        {apps.map((app) => <button type="button" key={app.id} onClick={() => create(app)}><AppIcon app={app.id} /><span><b>{translate(app.label)}</b><small>{translate(app.description)}</small></span></button>)}
      </section>
      <div className="suite-side-grow" />
      <div className="suite-side-links">
        {canInstall && <button type="button" onClick={() => {
          const prompt = installPrompt.current;
          if (!prompt) return;
          void prompt.prompt().then(() => prompt.userChoice).finally(() => { installPrompt.current = null; setCanInstall(false); });
        }}><StartFooterIcon kind="install" />{translate('Install VibeOffice.work')}</button>}
        <a href="https://github.com/tcchen2026/VibeOffice" target="_blank" rel="noopener"><StartFooterIcon kind="feedback" />{translate('Feedback on our GitHub')}</a>
      </div>
      <p className="suite-local-note">{translate('Your files stay on this computer.')}<br /><span className="suite-version" translate="no">{window.VO?.VERSION ?? 'dev'}</span></p>
      <div className="suite-side-links suite-footer-links">
        <button type="button" onClick={() => setAppearanceOpen('settings')}><StartFooterIcon kind="settings" />{translate('Settings')}</button>
        <button type="button" onClick={() => setAppearanceOpen('themes')}><StartFooterIcon kind="themes" />{translate('Themes')}</button>
        <button type="button" onClick={() => setAboutOpen(true)}><StartFooterIcon kind="about" />{translate('About')}</button>
      </div>
      <input ref={picker} type="file" hidden onChange={(event) => { const selected = event.currentTarget.files?.[0]; event.currentTarget.value = ''; openFile(selected); }} />
    </nav>
    <main className="suite-main">
      <header className="suite-bar"><h1>{translate('Create File')}</h1><span className="suite-bar-spacer" />
        <button type="button" onClick={openUrl}>{translate(UI_TEXT.openUrl)}</button><button type="button" onClick={() => openWebDav()}>{translate(UI_TEXT.webdav)}</button>
      </header>
      <section className="suite-gallery" aria-label={translate('Create File')}>
      {apps.map((app) => <section className="app" key={app.id}><h2><span className={`suite-app-mark ${app.id}`} aria-hidden="true" />{translate(app.id === 'wordx' ? 'Documents' : app.id === 'excelx' ? 'Spreadsheets' : 'Presentations')}<small>· {translate(app.label)}</small></h2><div className="cards">
        {(templates[app.id] || []).map((template) => <button type="button" key={template.id || 'blank'} title={translate(template.id ? 'New document from “{0}”' : app.id === 'wordx' ? 'New blank document' : app.id === 'excelx' ? 'New blank spreadsheet' : 'New blank presentation', ...(template.id ? [translate(template.name)] : []))} onClick={() => create(app, template.id || undefined)}><span className="suite-template-thumb">{template.img ? <img src={`/${template.img}`} alt="" loading="lazy" /> : <span className={`suite-blank-preview ${app.id}`} aria-hidden="true"><span /></span>}</span><span className="suite-template-name">{translate(template.name)}</span></button>)}
        {!(templates[app.id] || []).some((template) => !template.id) && <button type="button" title={translate(app.id === 'wordx' ? 'New blank document' : app.id === 'excelx' ? 'New blank spreadsheet' : 'New blank presentation')} onClick={() => create(app)}><span className="suite-template-thumb"><span className={`suite-blank-preview ${app.id}`} aria-hidden="true"><span /></span></span><span className="suite-template-name">{translate(app.id === 'wordx' ? 'Blank Document' : app.id === 'excelx' ? 'Blank Workbook' : 'Blank Presentation')}</span></button>}
      </div></section>)}
      </section>
    </main>
    {urlOpen && <OfficeDialog title={translate(UI_TEXT.openUrl)} className="credential" onRequestClose={() => setUrlOpen(false)}><form onSubmit={(event) => void submitUrl(event)}><label>{translate('URL')}<input autoFocus required type="url" aria-label={translate('URL')} value={urlValue} onChange={(event) => setUrlValue(event.currentTarget.value)} /></label><button type="button" onClick={() => setUrlOpen(false)}>{translate(UI_TEXT.cancel)}</button><button type="submit" disabled={urlBusy}>{translate('Open')}</button></form>{urlError && <p role="alert">{urlError}</p>}</OfficeDialog>}
    {appearanceOpen && <OfficeDialog title={translate(appearanceOpen === 'settings' ? 'Settings' : 'Themes')} className="credential suite-preferences" onRequestClose={() => setAppearanceOpen(null)}>
      <AppearanceFields
        look={office.look}
        looks={appearanceOpen === 'themes' ? office.looks.map(({ id, name }) => ({ value: id, label: translate(name) })) : []}
        language={office.language}
        languages={appearanceOpen === 'settings' ? office.languages.map(({ id, name }) => ({ value: id, label: translate(name) })) : []}
        onLookChange={(look) => office.setLook(look as typeof office.look)}
        onLanguageChange={(language) => { office.setLanguage(language as typeof office.language); location.reload(); }}
        translate={translate}
        lookLabel="Look"
        languageLabel="Language"
      />
      <button type="button" onClick={() => setAppearanceOpen(null)}>{translate('OK')}</button>
    </OfficeDialog>}
    {aboutOpen && <AboutDialog request={{ name: 'VibeOffice.work', version: translate('Web edition'), text: ['A browser office suite with Wordx, Excelx, and PPTX. Your documents stay on this computer.'], licensesUrl: '/oss-licenses.html' }} onClose={() => setAboutOpen(false)} />}
    {webDavDialog}
    {credentialAsk && <OfficeDialog title={credentialAsk.label} className="credential" onRequestClose={() => finishCredential(null)}><form onSubmit={(event) => { event.preventDefault(); finishCredential(credentialInput); }}><label>{credentialAsk.label}<input autoFocus type={credentialAsk.secret ? 'password' : 'text'} value={credentialInput} onChange={(event) => setCredentialInput(event.currentTarget.value)} /></label><button type="button" onClick={() => finishCredential(null)}>{translate(UI_TEXT.cancel)}</button><button type="submit">{translate(UI_TEXT.continue)}</button></form></OfficeDialog>}
  </div>;
}

const root = document.getElementById('root');
if (!root) throw new Error('Missing #root mount element');
createRoot(root).render(<VibeOfficeProvider><Suite /></VibeOfficeProvider>);
