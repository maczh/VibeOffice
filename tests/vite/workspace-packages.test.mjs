import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { WordxEditor } from '@vibeoffice/wordx';
import { ExcelxEditor } from '@vibeoffice/excelx';
import { PptxEditor } from '@vibeoffice/pptx';
import { AppearanceControls, OfficeDialog, VibeOfficeProvider } from '@vibeoffice/core';
import { createDirectEngineBridge, createIndexedDbDraftStore, openLocalDocument, resolveUnsavedChanges, saveLocalDocument } from '@vibeoffice/core';
import { IDBFactory } from 'fake-indexeddb';
import { createRemoteFileCallbacks, createUrlReadProvider, createWebDavProvider, remoteConflict } from '@vibeoffice/core';
import { reactCsp } from '../../tools/vite/react-csp.mjs';

const editors = [
  { component: WordxEditor, label: 'Wordx editor', id: 'wordx' },
  { component: ExcelxEditor, label: 'Excelx editor', id: 'excelx' },
  { component: PptxEditor, label: 'PPTX editor', id: 'pptx' },
];

test('a host can mount each editor package through its public React export', () => {
  for (const { component, label, id } of editors) {
    const html = renderToStaticMarkup(createElement(component, { className: 'host-embed', services: { translate: (message) => message } }));
    assert.match(html, new RegExp(`aria-label="${label}"`));
    assert.match(html, new RegExp(`data-vibe-editor="${id}"`));
    assert.match(html, /class="vibe-editor host-embed"/);
  }
});

test('Wordx exposes a React-owned engine mount point inside its editor frame', () => {
  const html = renderToStaticMarkup(createElement(WordxEditor, {
    mountEngine: () => {},
    children: createElement('p', null, 'Engine fixture'),
  }));
  assert.match(html, /data-vibe-editor="wordx"/);
  assert.match(html, /class="vibe-editor__engine"><p>Engine fixture<\/p><\/div>/);
});

test('async editor mounts dispose late resources and report only live mount failures', async () => {
  const { settleEditorMount } = await import('@vibeoffice/core');
  let disposed = false, disposedLateMount = 0, reported = 0;
  let resolveMount;
  const pending = new Promise((resolve) => { resolveMount = resolve; });
  settleEditorMount(pending, () => disposed, () => assert.fail('disposed mount must not be retained'), () => assert.fail('successful mount must not report an error'));
  disposed = true;
  resolveMount(() => disposedLateMount++);
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(disposedLateMount, 1, 'a disposer returned after unmount runs immediately');

  let rejectMount;
  const lateFailure = new Promise((_resolve, reject) => { rejectMount = reject; });
  settleEditorMount(lateFailure, () => true, () => assert.fail('rejected mount must not be retained'), () => reported++);
  rejectMount(new Error('late failure'));
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(reported, 0, 'rejection after unmount is consumed without calling stale UI');

  settleEditorMount(Promise.reject(new Error('active failure')), () => false, () => {}, () => reported++);
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(reported, 1, 'a live mount failure reaches its error handler');
});

test('React packages expose the same-document direct engine mount adapter', () => {
  assert.equal(typeof createDirectEngineBridge, 'function');
  const frame = readFileSync(new URL('../../commonx/src/editor-frame.tsx', import.meta.url), 'utf8');
  const runtime = readFileSync(new URL('../../commonx/src/direct-engine.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(frame, /iframe|contentWindow|postMessage/);
  assert.match(runtime, /mountLegacyEngineDocument/);
  assert.doesNotMatch(runtime, /createElement\(['"]iframe/);
});

test('shared provider exports the existing look and language catalogs and React hooks', async () => {
  const html = renderToStaticMarkup(createElement(VibeOfficeProvider, {
    services: { translate: (message) => message },
    children: createElement('span', null, 'ready'),
  }));
  assert.match(html, /ready/);
  const source = await import('@vibeoffice/core');
  assert.equal(typeof source.useVibeOffice, 'function');
  const bridge = readFileSync(new URL('../../commonx/src/react-bridge.tsx', import.meta.url), 'utf8');
  assert.match(bridge, /id: 'paper2016'/);
  assert.match(bridge, /id: 'zh-Hant'/);
});

test('shared appearance controls render the same look and language choices', () => {
  const html = renderToStaticMarkup(createElement(VibeOfficeProvider, {
    services: { translate: (message) => message },
    children: createElement(AppearanceControls),
  }));
  assert.match(html, /aria-label="Look"/);
  assert.match(html, /value="paper2016">Paper 2016/);
  assert.match(html, /aria-label="Language"/);
  assert.match(html, /value="zh-Hant">繁體中文/);
});

test('shared native dialog exposes an accessible title and request-close affordance', () => {
  const html = renderToStaticMarkup(createElement(OfficeDialog, { title: 'Credentials', onRequestClose() {}, children: createElement('button', null, 'Continue') }));
  assert.match(html, /<dialog[^>]*aria-modal="true"[^>]*aria-labelledby="[^"]+"/);
  assert.match(html, /<h2 id="[^"]+">Credentials<\/h2>/);
  assert.match(html, /Continue/);
});

test('local drafts use their own database and support recovery without touching Recent Files', async () => {
  const store = createIndexedDbDraftStore(new IDBFactory());
  const draft = { id: 'wordx:local-1', editor: 'wordx', name: 'Plan.docx', updatedAt: 100, snapshot: { text: 'unsaved' } };
  await store.put(draft);
  assert.deepEqual(await store.list('wordx'), [draft]);
  assert.deepEqual(await store.list('excelx'), []);
  await store.remove(draft.id);
  assert.deepEqual(await store.list(), []);
});

test('local open/save honors a writable file handle and returns the updated file', async () => {
  let written;
  const handle = {
    async getFile() { return new File([written ?? 'original'], 'Plan.docx'); },
    async createWritable() {
      return { async write(blob) { written = await blob.text(); }, async close() {} };
    },
  };
  const opened = await openLocalDocument(undefined, { showOpenFilePicker: async () => [handle] });
  assert.equal(await opened.file.text(), 'original');
  const saved = await saveLocalDocument(opened, new Blob(['updated']), 'Plan.docx', { document: { createElement() { throw new Error('download fallback should not run'); } } });
  assert.equal(written, 'updated');
  assert.equal(await saved.file.text(), 'updated');
  assert.equal(saved.handle, handle);
});

test('local file open rejects promptly when its mount AbortSignal is cancelled', async () => {
  const controller = new AbortController();
  let pickerSettled = false;
  let finishPicker;
  const picker = { showOpenFilePicker: () => new Promise((resolve) => { finishPicker = () => { pickerSettled = true; resolve([]); }; }) };
  const opening = openLocalDocument(undefined, picker, controller.signal);
  controller.abort();
  await assert.rejects(opening, { name: 'AbortError' });
  assert.equal(pickerSettled, false, 'the host stops waiting even if the browser-owned picker cannot be programmatically dismissed');
  finishPicker();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(pickerSettled, true, 'the eventual native picker completion is safely ignored');
});

test('local file-input fallback removes its picker when the mount is aborted', async () => {
  const previousDocument = globalThis.document;
  const listeners = new Map();
  const input = { type: '', accept: '', files: [], removed: false, addEventListener(type, listener) { listeners.set(type, listener); }, removeEventListener(type) { listeners.delete(type); }, click() {}, remove() { this.removed = true; } };
  globalThis.document = { createElement: (tag) => { assert.equal(tag, 'input'); return input; } };
  try {
    const controller = new AbortController();
    const opening = openLocalDocument(undefined, {}, controller.signal);
    assert.equal(input.type, 'file');
    controller.abort();
    await assert.rejects(opening, { name: 'AbortError' });
    assert.equal(input.removed, true);
    assert.equal(listeners.size, 0);
  } finally { globalThis.document = previousDocument; }
});

test('local save picker and draft transactions abort with their editor mount', async () => {
  const pickerController = new AbortController();
  const picker = { showSaveFilePicker: () => new Promise(() => {}) };
  const saving = saveLocalDocument({ file: new File([], 'Plan.docx') }, new Blob(['save']), 'Plan.docx', picker, pickerController.signal);
  pickerController.abort();
  await assert.rejects(saving, { name: 'AbortError' });

  const drafts = createIndexedDbDraftStore(new IDBFactory());
  const draftController = new AbortController();
  const writing = drafts.put({ id: 'wordx:abort', editor: 'wordx', name: 'Plan.docx', updatedAt: 1, snapshot: new Blob(['draft']) }, draftController.signal);
  draftController.abort();
  await assert.rejects(writing, { name: 'AbortError' });
});

test('local save aborts an in-progress writable stream', async () => {
  const controller = new AbortController();
  let startWrite;
  let rejectWrite;
  let abortCalls = 0;
  const writeStarted = new Promise((resolve) => { startWrite = resolve; });
  const handle = {
    async createWritable() {
      return {
        write() {
          startWrite();
          return new Promise((_, reject) => { rejectWrite = reject; });
        },
        async abort() {
          abortCalls++;
          rejectWrite?.(new DOMException('Aborted', 'AbortError'));
        },
        async close() {},
      };
    },
  };
  const saving = saveLocalDocument({ file: new File([], 'Plan.docx'), handle }, new Blob(['save']), 'Plan.docx', {}, controller.signal);
  await writeStarted;
  controller.abort();
  await assert.rejects(saving, { name: 'AbortError' });
  assert.ok(abortCalls > 0, 'cancelling the mount aborts the active writable stream');
});

test('an aborted mount never starts the local download fallback', async () => {
  const previousWindow = globalThis.window;
  let anchorsCreated = 0;
  globalThis.window = { document: { createElement() { anchorsCreated++; throw new Error('download must not start after unmount'); } } };
  try {
    const controller = new AbortController();
    controller.abort();
    await assert.rejects(saveLocalDocument({ file: new File([], 'Plan.docx') }, new Blob(['save']), 'Plan.docx', {}, controller.signal), { name: 'AbortError' });
    assert.equal(anchorsCreated, 0);
  } finally { globalThis.window = previousWindow; }
});

test('dirty local documents preserve Save, Discard, and Cancel behavior', async () => {
  for (const [choice, expected, expectedCalls] of [
    ['save', true, ['save']],
    ['discard', true, ['discard']],
    ['cancel', false, []],
  ]) {
    const calls = [];
    let dirty = true;
    const allowed = await resolveUnsavedChanges({
      isDirty: () => dirty,
      choose: async () => choice,
      async save() { calls.push('save'); dirty = false; },
      async discard() { calls.push('discard'); dirty = false; },
    });
    assert.equal(allowed, expected);
    assert.deepEqual(calls, expectedCalls);
  }
  await assert.rejects(resolveUnsavedChanges({
    isDirty: () => true,
    choose: async () => 'save',
    async save() { throw new Error('save failed'); },
    async discard() {},
  }), /save failed/);
});

test('remote URL provider is read-only and rejects unconfigured origins', async () => {
  let requested;
  const provider = createUrlReadProvider(['https://files.example'], async (url) => {
    requested = String(url);
    return new Response('document', { headers: { ETag: '"v1"' } });
  });
  const result = await provider.read('https://files.example/a.docx');
  assert.equal(requested, 'https://files.example/a.docx');
  assert.equal(await result.blob.text(), 'document');
  assert.deepEqual(result.file.revision, { value: '"v1"', kind: 'etag' });
  await assert.rejects(provider.read('https://other.example/a.docx'), /not allowed/);
  await assert.rejects(provider.read('https://user:secret@files.example/a.docx'), /not allowed/);
  assert.equal(remoteConflict().code, 'revision-conflict');
});

test('WebDAV provider limits origin and path, keeps credentials callback-scoped, and uses conditional writes', async () => {
  const calls = [];
  let credentialCalls = 0;
  const provider = createWebDavProvider({
    baseUrl: 'https://dav.example/docs/', allowedOrigins: ['https://dav.example'],
    credentials: async () => { credentialCalls++; return { username: 'u', password: 'p' }; },
    fetcher: async (url, init) => {
      calls.push({ url: String(url), init });
      if (String(url).includes('/conflict')) return new Response('', { status: 412 });
      return new Response('', { status: init.method === 'PROPFIND' ? 207 : 201, headers: { ETag: '"v2"' } });
    },
  });
  await provider.read('/docs/a.docx');
  await provider.save('/docs/a.docx', new Blob(['saved']), { ifRevision: { value: '"v1"', kind: 'etag' } });
  assert.equal(calls[0].init.method, 'GET');
  assert.equal(calls[1].init.headers['If-Match'], '"v1"');
  assert.match(calls[0].init.headers.Authorization, /^Basic /);
  assert.equal(credentialCalls, 2);
  await assert.rejects(provider.read('https://evil.example/a.docx'), /escaped/);
  await assert.rejects(provider.read('/docs-escape/a.docx'), /escaped/);
  await assert.rejects(provider.read('https://user:secret@dav.example/docs/a.docx'), /escaped/);
  await assert.rejects(provider.save('/docs/conflict.docx', new Blob(['x']), { ifRevision: { value: '"stale"', kind: 'etag' } }), (error) => error.code === 'revision-conflict');
  assert.throws(() => createWebDavProvider({ baseUrl: 'https://evil.example/', allowedOrigins: ['https://dav.example'] }), /not allowed/);
  assert.throws(() => createWebDavProvider({ baseUrl: 'https://user:secret@dav.example/docs/', allowedOrigins: ['https://dav.example'] }), /not allowed/);
});

test('remote editor callbacks retain file revision and use conditional save or create', async () => {
  const calls = [];
  const provider = {
    async read(id) { calls.push(['read', id]); return { file: { id, name: 'a.docx', revision: { value: 'v1', kind: 'etag' } }, blob: new Blob(['opened']) }; },
    async save(id, blob, options) { calls.push(['save', id, options]); return { id, name: 'a.docx', revision: { value: 'v2', kind: 'etag' } }; },
    async create(parentId, name) { calls.push(['create', parentId, name]); return { id: 'created', name }; },
  };
  const callbacks = createRemoteFileCallbacks(provider, async () => ({ id: 'file-1', name: 'a.docx' }), async () => ({ parentId: 'folder', name: 'new.docx' }));
  const opened = await callbacks.onOpenFile();
  assert.equal(await opened.file.text(), 'opened');
  const saved = await callbacks.onSaveFile('a.docx', new Blob(['updated']), opened);
  assert.equal(saved.remote.file.revision.value, 'v2');
  await callbacks.onSaveFile('new.docx', new Blob(['new']));
  assert.deepEqual(calls.map(([op]) => op), ['read', 'save', 'create']);
  assert.deepEqual(calls[1][2], { ifRevision: { value: 'v1', kind: 'etag' } });
});

test('suite CSP only adds explicitly configured remote origins', () => {
  const plugin = reactCsp('https://dav.example,https://api.example');
  const html = plugin.transformIndexHtml('<head></head>');
  assert.match(html, /connect-src 'self' https:\/\/dav\.example https:\/\/api\.example/);
  assert.match(html, /script-src 'self'/);
  assert.throws(() => reactCsp('https://dav.example/path'), /Invalid remote origin/);
});
