import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent } from 'react';
import { createRoot } from 'react-dom/client';
import { AppearanceControls, createUrlReadProvider, createWebDavProvider, OfficeDialog, saveLocalDocument, VibeOfficeProvider, useVibeOffice } from '@vibeoffice/core';
import { createIndexedDbDraftStore } from '@vibeoffice/core';
import type { LocalDocument, LocalDraft, RemoteFile, RemoteStorageProvider, SaveFileOperation } from '@vibeoffice/core';
import { WordxEditor } from '@vibeoffice/wordx';
import { ExcelxEditor } from '@vibeoffice/excelx';
import { PptxEditor } from '@vibeoffice/pptx';
import './style.css';

const mark = (message: string) => window.VO?.N?.(message) ?? message;

type AppId = 'wordx' | 'excelx' | 'pptx';
type Entry = { id: AppId; label: string; folder: string; extensions: string[] };
type Template = { id: string; name: string; img?: string | null };
const apps: Entry[] = [
  { id: 'wordx', label: 'Wordx', folder: 'wordx', extensions: ['docx', 'docm', 'dotx', 'dotm', 'rtf', 'txt', 'md', 'markdown', 'html', 'htm', 'zip'] },
  { id: 'excelx', label: 'Excelx', folder: 'excelx', extensions: ['xlsx', 'xlsm', 'xltx', 'xltm', 'csv', 'tsv', 'txt', 'xml', 'html', 'htm'] },
  { id: 'pptx', label: 'PPTX', folder: 'pptx', extensions: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm'] },
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

function requestEditorClose(): Promise<boolean> {
  if (!document.querySelector('.suite-frame .vibe-editor__engine > *')) return Promise.resolve(true);
  return window.VO?.reactClose ? window.VO.reactClose().catch(() => false) : Promise.resolve(true);
}

function ActiveEditor({ app, file, document, template, allowedRemoteOrigins, onSaveFile, onDraftsAvailable }: { app: Entry; file?: File; document?: LocalDocument; template?: string; allowedRemoteOrigins: readonly string[]; onSaveFile: (name: string, blob: Blob, current?: LocalDocument, signal?: AbortSignal) => Promise<LocalDocument | void | null>; onDraftsAvailable: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void }) {
  const query = template ? `template=${encodeURIComponent(template)}` : 'new';
  const key = `${app.id}:${file?.name || template || 'blank'}`;
  const engine = useMemo(() => ({ url: `/${app.folder}/?${query}` }), [app.folder, query]);
  const recovery = { draftStore, onDraftsAvailable };
  if (app.id === 'wordx') return <WordxEditor key={key} className="suite-frame" engine={engine} allowedRemoteOrigins={allowedRemoteOrigins} onSaveFile={onSaveFile} {...recovery} {...(document ? { initialDocument: document } : file ? { initialFile: file } : {})} />;
  if (app.id === 'excelx') return <ExcelxEditor key={key} className="suite-frame" engine={engine} allowedRemoteOrigins={allowedRemoteOrigins} onSaveFile={onSaveFile} {...recovery} {...(document ? { initialDocument: document } : file ? { initialFile: file } : {})} />;
  return <PptxEditor key={key} className="suite-frame" engine={engine} allowedRemoteOrigins={allowedRemoteOrigins} onSaveFile={onSaveFile} {...recovery} {...(document ? { initialDocument: document } : file ? { initialFile: file } : {})} />;
}

function Suite() {
  const office = useVibeOffice();
  const { translate } = office;
  const [active, setActive] = useState<{ app: Entry; file?: File; document?: LocalDocument; template?: string } | null>(null);
  const [recoverable, setRecoverable] = useState<LocalDraft<Blob>[]>([]);
  const restoreDraft = useRef<((draft: LocalDraft<Blob>) => void) | null>(null);
  const onDraftsAvailable = useCallback((drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => {
    restoreDraft.current = restore;
    setRecoverable(drafts);
  }, []);
  const [templates, setTemplates] = useState<Record<string, Template[]>>({});
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
      <ActiveEditor app={app} allowedRemoteOrigins={allowedOrigins} onSaveFile={onSaveFile} onDraftsAvailable={onDraftsAvailable} {...(file ? { file } : {})} {...(active.document ? { document: active.document } : {})} {...(template ? { template } : {})} />
      {conflictAsk && <OfficeDialog title={translate(UI_TEXT.changed)} className="credential" onRequestClose={() => conflictAsk.finish(false)}><p>{translate(UI_TEXT.changed)}</p><button type="button" onClick={() => conflictAsk.finish(false)}>{translate(UI_TEXT.cancel)}</button><button type="button" onClick={() => conflictAsk.finish(true)}>{translate(UI_TEXT.replace)}</button></OfficeDialog>}
      {saveAsAsk && <OfficeDialog title={translate(UI_TEXT.saveAs)} className="credential" onRequestClose={() => saveAsAsk.finish(null)}><form onSubmit={(event) => { event.preventDefault(); saveAsAsk.finish(saveAsName); }}><label>{translate(UI_TEXT.fileName)}<input autoFocus required value={saveAsName} onChange={(event) => setSaveAsName(event.currentTarget.value)} /></label><button type="button" onClick={() => openWebDav('saveAs')}>{translate(UI_TEXT.webdav)}</button><button type="button" onClick={() => saveAsAsk.finish(null)}>{translate(UI_TEXT.cancel)}</button><button type="submit">{translate(UI_TEXT.save)}</button></form></OfficeDialog>}
      {webDavDialog}
    </main>;
  }

  return <main className="suite-start">
    <header><img src="/icons/icon.svg" alt="" /><h1 translate="no">VibeOffice</h1>
      <AppearanceControls onLanguageChange={() => location.reload()} />
      <button type="button" onClick={() => picker.current?.click()}>{translate('Open File…')}</button><button type="button" onClick={openUrl}>{translate(UI_TEXT.openUrl)}</button><button type="button" onClick={() => openWebDav()}>{translate(UI_TEXT.webdav)}</button><input ref={picker} type="file" hidden onChange={(event) => { const selected = event.currentTarget.files?.[0]; event.currentTarget.value = ''; openFile(selected); }} /></header>
    <section><h2>{translate('Create File')}</h2>{apps.map((app) => <section className="app" key={app.id}><h3>{translate(app.label)}</h3><div className="cards">
      <button type="button" onClick={() => create(app)}><b>{translate('Blank')}</b></button>
      {(templates[app.id] || []).map((template) => <button type="button" key={template.id} onClick={() => create(app, template.id)}>{template.img && <img src={template.img} alt="" loading="lazy" />}<b>{translate(template.name)}</b></button>)}
    </div></section>)}</section>
    <footer>{apps.map((app) => <button type="button" key={app.id} onClick={() => create(app)}>{translate(app.label)}</button>)}</footer>
    {urlOpen && <OfficeDialog title={translate(UI_TEXT.openUrl)} className="credential" onRequestClose={() => setUrlOpen(false)}><form onSubmit={(event) => void submitUrl(event)}><label>{translate('URL')}<input autoFocus required type="url" aria-label={translate('URL')} value={urlValue} onChange={(event) => setUrlValue(event.currentTarget.value)} /></label><button type="button" onClick={() => setUrlOpen(false)}>{translate(UI_TEXT.cancel)}</button><button type="submit" disabled={urlBusy}>{translate('Open')}</button></form>{urlError && <p role="alert">{urlError}</p>}</OfficeDialog>}
    {webDavDialog}
    {credentialAsk && <OfficeDialog title={credentialAsk.label} className="credential" onRequestClose={() => finishCredential(null)}><form onSubmit={(event) => { event.preventDefault(); finishCredential(credentialInput); }}><label>{credentialAsk.label}<input autoFocus type={credentialAsk.secret ? 'password' : 'text'} value={credentialInput} onChange={(event) => setCredentialInput(event.currentTarget.value)} /></label><button type="button" onClick={() => finishCredential(null)}>{translate(UI_TEXT.cancel)}</button><button type="submit">{translate(UI_TEXT.continue)}</button></form></OfficeDialog>}
  </main>;
}

const root = document.getElementById('root');
if (!root) throw new Error('Missing #root mount element');
createRoot(root).render(<VibeOfficeProvider><Suite /></VibeOfficeProvider>);
