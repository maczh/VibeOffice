import { useCallback, useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { createUrlReadProvider, createWebDavProvider, saveLocalDocument, VibeOfficeProvider, useVibeOffice } from '@vibeoffice/core';
import { createIndexedDbDraftStore } from '@vibeoffice/core';
import type { LocalDocument, LocalDraft, RemoteFile, RemoteStorageProvider } from '@vibeoffice/core';
import { WordxEditor } from '@vibeoffice/wordx';
import { ExcelxEditor } from '@vibeoffice/excelx';
import { PptxEditor } from '@vibeoffice/pptx';
import './style.css';

const mark = (message: string) => window.VO?.N?.(message) ?? message;

type AppId = 'quire' | 'ledger' | 'lectern';
type Entry = { id: AppId; label: string; folder: string; extensions: string[] };
type Template = { id: string; name: string; img?: string | null };
const apps: Entry[] = [
  { id: 'quire', label: 'Quire', folder: 'quire', extensions: ['docx', 'docm', 'dotx', 'dotm', 'rtf', 'txt', 'md', 'markdown', 'html', 'htm', 'zip'] },
  { id: 'ledger', label: 'Ledger', folder: 'ledger', extensions: ['xlsx', 'xlsm', 'xltx', 'xltm', 'csv', 'tsv', 'txt', 'xml', 'html', 'htm'] },
  { id: 'lectern', label: 'Lectern', folder: 'lectern', extensions: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm'] },
];
const UI_TEXT = {
  openUrl: mark('Open URL…'), webdav: mark('WebDAV…'), changed: mark('The remote file changed since it was opened. Replace it anyway?'),
  unsupported: mark('This file type is not supported.'), openError: mark('The file could not be opened: {0}'), davUrl: mark('WebDAV collection URL'),
  empty: mark('This WebDAV folder is empty.'), davFiles: mark('WebDAV files'), choose: mark('Enter a number, or Cancel to stop.'),
  username: mark('Username'), password: mark('Password'), continue: mark('Continue'), cancel: mark('Cancel'),
  recoverTitle: mark('Recover unsaved work'), recover: mark('Recover'), dismiss: mark('Dismiss'), savedAt: mark('Saved {0}'),
};
const draftStore = createIndexedDbDraftStore<Blob>();
function appForName(name: string): Entry | undefined {
  const extension = name.split('.').pop()?.toLowerCase();
  return apps.find((candidate) => candidate.extensions.includes(extension || ''));
}

function ActiveEditor({ app, file, document, template, onSaveFile, onDraftsAvailable }: { app: Entry; file?: File; document?: LocalDocument; template?: string; onSaveFile: (name: string, blob: Blob, current?: LocalDocument) => Promise<LocalDocument | void | null>; onDraftsAvailable: (drafts: LocalDraft<Blob>[], restore: (draft: LocalDraft<Blob>) => void) => void }) {
  const query = template ? `template=${encodeURIComponent(template)}` : 'new';
  const key = `${app.id}:${file?.name || template || 'blank'}`;
  const engine = { url: `/${app.folder}/?${query}` };
  const recovery = { draftStore, onDraftsAvailable };
  if (app.id === 'quire') return <WordxEditor key={key} className="suite-frame" engine={engine} onSaveFile={onSaveFile} {...recovery} {...(document ? { initialDocument: document } : file ? { initialFile: file } : {})} />;
  if (app.id === 'ledger') return <ExcelxEditor key={key} className="suite-frame" engine={engine} onSaveFile={onSaveFile} {...recovery} {...(document ? { initialDocument: document } : file ? { initialFile: file } : {})} />;
  return <PptxEditor key={key} className="suite-frame" engine={engine} onSaveFile={onSaveFile} {...recovery} {...(document ? { initialDocument: document } : file ? { initialFile: file } : {})} />;
}

function Suite() {
  const office = useVibeOffice();
  const { translate } = office;
  const [active, setActive] = useState<{ app: Entry; file?: File; document?: LocalDocument; template?: string } | null>(null);
  const [recoverable, setRecoverable] = useState<LocalDraft<Blob>[]>([]);
  const restoreDraft = useRef<((draft: LocalDraft<Blob>) => void) | null>(null);
  const [templates, setTemplates] = useState<Record<string, Template[]>>({});
  const picker = useRef<HTMLInputElement>(null);
  const [credentialAsk, setCredentialAsk] = useState<{ label: string; secret: boolean; finish: (value: string | null) => void } | null>(null);
  const [credentialInput, setCredentialInput] = useState('');
  const remoteTarget = useRef<{ provider: RemoteStorageProvider; parentId?: string } | null>(null);
  const allowedOrigins = (document.querySelector('meta[name="vibeoffice-remote-origins"]')?.getAttribute('content') || '').split(/\s+/).filter(Boolean);

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
  const onSaveFile = useCallback(async (name: string, blob: Blob, current?: LocalDocument) => {
    const remote = current?.remote;
    if (remote?.provider.save) {
      try {
        const file = await remote.provider.save(remote.file.id, blob, remote.file.revision ? { ifRevision: remote.file.revision } : {});
        return { file: new File([blob], file.name || name, { type: blob.type }), remote: { provider: remote.provider, file } };
      } catch (error) {
        if ((error as { code?: string })?.code !== 'revision-conflict' || !window.confirm(translate(UI_TEXT.changed))) throw error;
        const file = await remote.provider.save(remote.file.id, blob);
        return { file: new File([blob], file.name || name, { type: blob.type }), remote: { provider: remote.provider, file } };
      }
    }
    if (!remote && remoteTarget.current?.provider.create) {
      const target = remoteTarget.current;
      const file = await target.provider.create(target.parentId, name, blob);
      return { file: new File([blob], file.name || name, { type: blob.type }), remote: { provider: target.provider, file } };
    }
    return saveLocalDocument(current ?? { file: new File([], name) }, blob, name);
  }, [translate]);

  const openUrl = async () => {
    try {
      const url = window.prompt(translate(UI_TEXT.openUrl));
      if (!url) return;
      const reader = createUrlReadProvider(allowedOrigins);
      const { file, blob } = await reader.read(url);
      const app = appForName(file.name);
      if (!app) throw new Error(translate(UI_TEXT.unsupported));
      const opened = new File([blob], file.name, { type: file.mimeType || blob.type });
      setActive({ app, document: { file: opened, remote: { provider: reader, file } } });
    } catch (error) { window.alert(translate(UI_TEXT.openError, error instanceof Error ? error.message : String(error))); }
  };

  const openWebDav = async () => {
    try {
      const baseUrl = window.prompt(translate(UI_TEXT.davUrl));
      if (!baseUrl) return;
      let credentials: { username: string; password: string } | undefined;
      const provider = createWebDavProvider({ baseUrl, allowedOrigins, credentials: async () => {
        if (credentials) return credentials;
        const username = await askCredential(translate(UI_TEXT.username), false);
        if (username === null) return undefined;
        const password = await askCredential(translate(UI_TEXT.password), true);
        if (password === null) return undefined;
        credentials = { username, password };
        return credentials;
      } });
      remoteTarget.current = { provider };
      let parentId: string | undefined;
      let selected: RemoteFile | undefined;
      for (let depth = 0; depth < 12; depth++) {
        const entries = (await provider.list(parentId)).filter((entry) => entry.name);
        if (!entries.length) throw new Error(translate(UI_TEXT.empty));
        const choice = window.prompt(`${translate(UI_TEXT.davFiles)}\n${entries.map((entry, i) => `${i + 1}. ${entry.name}${entry.isDirectory ? '/' : ''}`).join('\n')}\n${translate(UI_TEXT.choose)}`);
        if (choice === null) return;
        selected = entries[Number(choice) - 1];
        if (!selected) continue;
        if (selected.isDirectory) { parentId = selected.id; remoteTarget.current = { provider, parentId }; continue; }
        break;
      }
      if (!selected || selected.isDirectory) return;
      const { file, blob } = await provider.read(selected.id);
      const app = appForName(file.name);
      if (!app) throw new Error(translate(UI_TEXT.unsupported));
      const opened = new File([blob], file.name, { type: file.mimeType || blob.type });
      remoteTarget.current = { provider, ...(parentId !== undefined ? { parentId } : {}) };
      setActive({ app, document: { file: opened, remote: { provider, file } } });
    } catch (error) { window.alert(translate(UI_TEXT.openError, error instanceof Error ? error.message : String(error))); }
  };

  if (active) {
    const { app, file, template } = active;
    return <main className="suite-editor">
      <button className="back" type="button" onClick={() => setActive(null)}>{translate('Create File')}</button>
      {recoverable.length > 0 && <aside className="draft-recovery" aria-label={translate(UI_TEXT.recoverTitle)}><strong>{translate(UI_TEXT.recoverTitle)}</strong>{recoverable.map((draft) => <div key={draft.id}><span>{draft.name} · {translate(UI_TEXT.savedAt, new Date(draft.updatedAt).toLocaleString())}</span><button type="button" onClick={() => restoreDraft.current?.(draft)}>{translate(UI_TEXT.recover)}</button><button type="button" onClick={() => { void draftStore.remove(draft.id).then(() => setRecoverable((items) => items.filter((item) => item.id !== draft.id))); }}>{translate(UI_TEXT.dismiss)}</button></div>)}</aside>}
      <ActiveEditor app={app} onSaveFile={onSaveFile} onDraftsAvailable={(drafts, restore) => { setRecoverable(drafts); restoreDraft.current = restore; }} {...(file ? { file } : {})} {...(active.document ? { document: active.document } : {})} {...(template ? { template } : {})} />
    </main>;
  }

  return <main className="suite-start">
    <header><img src="/icons/icon.svg" alt="" /><h1 translate="no">VibeOffice</h1>
      <label>{translate('Look')} <select aria-label={translate('Look')} value={office.look} onChange={(event) => office.setLook(event.currentTarget.value as typeof office.look)}>{office.looks.map((look) => <option key={look.id} value={look.id}>{translate(look.name)}</option>)}</select></label>
      <label>{translate('Language')} <select aria-label={translate('Language')} value={office.language} onChange={(event) => { localStorage.setItem('vibeoffice.lang', event.currentTarget.value); location.reload(); }}>{office.languages.map((language) => <option key={language.id} value={language.id}>{translate(language.name)}</option>)}</select></label>
      <button type="button" onClick={() => picker.current?.click()}>{translate('Open File…')}</button><button type="button" onClick={() => void openUrl()}>{translate(UI_TEXT.openUrl)}</button><button type="button" onClick={() => void openWebDav()}>{translate(UI_TEXT.webdav)}</button><input ref={picker} type="file" hidden onChange={(event) => { const selected = event.currentTarget.files?.[0]; event.currentTarget.value = ''; openFile(selected); }} /></header>
    <section><h2>{translate('Create File')}</h2>{apps.map((app) => <section className="app" key={app.id}><h3>{translate(app.label)}</h3><div className="cards">
      <button type="button" onClick={() => create(app)}><b>{translate('Blank')}</b></button>
      {(templates[app.id] || []).map((template) => <button type="button" key={template.id} onClick={() => create(app, template.id)}>{template.img && <img src={template.img} alt="" loading="lazy" />}<b>{translate(template.name)}</b></button>)}
    </div></section>)}</section>
    <footer>{apps.map((app) => <button type="button" key={app.id} onClick={() => create(app)}>{translate(app.label)}</button>)}</footer>
    {credentialAsk && <dialog open className="credential"><form onSubmit={(event) => { event.preventDefault(); finishCredential(credentialInput); }}><label>{credentialAsk.label}<input autoFocus type={credentialAsk.secret ? 'password' : 'text'} value={credentialInput} onChange={(event) => setCredentialInput(event.currentTarget.value)} /></label><button type="button" onClick={() => finishCredential(null)}>{translate(UI_TEXT.cancel)}</button><button type="submit">{translate(UI_TEXT.continue)}</button></form></dialog>}
  </main>;
}

const root = document.getElementById('root');
if (!root) throw new Error('Missing #root mount element');
createRoot(root).render(<VibeOfficeProvider><Suite /></VibeOfficeProvider>);
