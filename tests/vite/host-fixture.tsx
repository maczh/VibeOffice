import { createRoot } from 'react-dom/client';
import { useCallback, useMemo, useState } from 'react';
import { WordxEditor } from '@vibeoffice/wordx';
import { ExcelxEditor } from '@vibeoffice/excelx';
import { PptxEditor } from '@vibeoffice/pptx';
import { createRemoteFileCallbacks, useVibeOffice, VibeOfficeProvider } from '@vibeoffice/core';
import type { RemoteFile, RemoteStorageProvider } from '@vibeoffice/core';

type EditorId = 'wordx' | 'excelx' | 'pptx';
const ENGINE = { wordx: { url: '/wordx/?new' }, excelx: { url: '/excelx/?new' }, pptx: { url: '/pptx/?new' } } as const;

function HostFixture() {
  const office = useVibeOffice();
  const [editor, setEditor] = useState<EditorId>('wordx');
  const [engineReady, setEngineReady] = useState(false);
  const [editorTitle, setEditorTitle] = useState('');
  const [providerStatus, setProviderStatus] = useState('Provider ready');
  const onEngineReady = useCallback(() => setEngineReady(true), []);
  const remote = useMemo(() => {
    let revision = 1;
    const file: RemoteFile = { id: 'sample.txt', name: 'sample.txt', mimeType: 'text/plain', revision: { value: '1', kind: 'version' } };
    const currentFile = (): RemoteFile => ({ ...file, revision: { value: String(revision), kind: 'version' } });
    const provider: RemoteStorageProvider = {
      async list() { return [currentFile()]; },
      async read(id) {
        if (id !== file.id) throw new Error(`Unknown host-provider file: ${id}`);
        setProviderStatus('Read sample.txt from the host provider');
        return { file: currentFile(), blob: new Blob(['Opened through the host-provided storage adapter.'], { type: 'text/plain' }) };
      },
      async create(_parentId, name) {
        revision++;
        setProviderStatus(`Created ${name} at revision ${revision}`);
        return { id: name, name, revision: { value: String(revision), kind: 'version' } };
      },
      async save(id, _blob, options) {
        if (options?.ifRevision?.value !== String(revision)) throw new Error('The host provider detected a stale revision.');
        revision++;
        setProviderStatus(`Saved ${id} at revision ${revision}`);
        return { ...file, id, name: id, revision: { value: String(revision), kind: 'version' } };
      },
    };
    return createRemoteFileCallbacks(provider, async () => currentFile(), async (name) => ({ id: file.id, name, revision: currentFile().revision }));
  }, []);
  const engineWindow = window as Window & { VO?: { reactDirectHost?: { open?(): Promise<void>; save?(name: string, blob: Blob): Promise<string> } } };
  const setLook = (look: typeof office.look) => {
    document.documentElement.dataset.look = look;
    document.documentElement.dataset.layout = look === 'paper2016' ? 'ribbon' : 'bars';
    window.dispatchEvent(new Event('vo-look'));
  };
  const renderErrorCase = new URLSearchParams(window.location.search).has('error-boundary');
  const failingLayer = renderErrorCase ? () => <FailingHostLayer /> : undefined;

  return (
    <main className="fixture">
      <header className="fixture-head">
        <h1>{office.translate('VibeOffice editors')}</h1>
        <label>
          {office.translate('Look')}{' '}
          <select aria-label={office.translate('Look')} value={office.look} onChange={(event) => setLook(event.target.value as typeof office.look)}>
            {office.looks.map(({ id, name }) => <option key={id} value={id}>{office.translate(name)}</option>)}
          </select>
        </label>
        <label>
          <select aria-label={office.translate('VibeOffice editors')} value={editor} onChange={(event) => { setEngineReady(false); setEditor(event.target.value as EditorId); }}>
            <option value="wordx">Wordx</option>
            <option value="excelx">Excelx</option>
            <option value="pptx">PPTX</option>
          </select>
        </label>
        <button type="button" disabled={!engineReady} onClick={() => engineWindow.VO?.reactDirectHost?.open?.()}>Open host-provider sample</button>
        <button type="button" disabled={!engineReady} onClick={() => engineWindow.VO?.reactDirectHost?.save?.('sample.txt', new Blob(['Saved through the host-provided storage adapter.'], { type: 'text/plain' }))}>Save through host provider</button>
        <output aria-label="Host provider status">{providerStatus}</output>
        <output aria-label="Editor title">{editorTitle}</output>
      </header>
      <section className="fixture-editors">
        {editor === 'wordx' && <WordxEditor key="wordx" engine={ENGINE.wordx} onEngineReady={onEngineReady} onTitleChange={setEditorTitle} onOpenFile={remote.onOpenFile} onSaveFile={remote.onSaveFile} {...(renderErrorCase ? { appLayers: failingLayer, onRenderError: (error: Error) => setProviderStatus(error.message), renderError: () => <div role="alert" data-host-render-fallback="true">Editor layer unavailable</div> } : {})} />}
        {editor === 'excelx' && <ExcelxEditor key="excelx" engine={ENGINE.excelx} onEngineReady={onEngineReady} onTitleChange={setEditorTitle} onOpenFile={remote.onOpenFile} onSaveFile={remote.onSaveFile} />}
        {editor === 'pptx' && <PptxEditor key="pptx" engine={ENGINE.pptx} onEngineReady={onEngineReady} onTitleChange={setEditorTitle} onOpenFile={remote.onOpenFile} onSaveFile={remote.onSaveFile} />}
      </section>
    </main>
  );
}

function FailingHostLayer(): never {
  throw new Error('Host-provided React layer failed');
}

const root = document.getElementById('root');
if (!root) throw new Error('Missing #root mount element');
createRoot(root).render(<VibeOfficeProvider><HostFixture /></VibeOfficeProvider>);
