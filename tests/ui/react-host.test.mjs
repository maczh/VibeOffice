import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open } from '../lib/input.mjs';

test('Quire React-host mode starts the editor without opening the legacy Recent Files database', async () => {
  const t = await open('quire', {
    persistence: true,
    query: { reactHost: '' },
    preload: `window.__legacyDatabaseOpens = 0;
      const open = indexedDB.open.bind(indexedDB);
      indexedDB.open = (...args) => { window.__legacyDatabaseOpens++; return open(...args); };`,
  });
  try {
    assert.equal(await t.evaluate('VO.reactHost'), true);
    assert.equal(await t.evaluate('window.__legacyDatabaseOpens'), 0);
    assert.deepEqual(await t.evaluate('VO.recent.list()'), []);
    assert.deepEqual(await t.evaluate('VO.recentMenuItems()'), []);
    assert.equal(await t.evaluate('!!L.D.doc'), true, 'the Quire document engine initialized');
    await t.evaluate(`window.postMessage({ type: 'vibeoffice-react-restore', id: 'wordx:test-recovery', name: 'Recovered.txt', blob: new Blob(['Recovered text']), lossState: null }, location.origin)`);
    await new Promise((resolve) => setTimeout(resolve, 500));
    assert.equal(await t.evaluate('L.D.plainText(L.D.doc.main.blocks[0])'), 'Recovered text');
    assert.equal(await t.evaluate('VO.changeCount(L.D.doc) > 0'), true, 'restored work is dirty and autosave is re-armed');
    const openRequested = await t.evaluate(`new Promise((resolve) => {
      const listener = (event) => { if (event.data?.type === 'vibeoffice-react-request-open') { removeEventListener('message', listener); resolve(true); } };
      addEventListener('message', listener); L.app.openDialog(); setTimeout(() => resolve(false), 500);
    })`);
    assert.equal(openRequested, true, 'Quire Open requests use the React host');
    const saveResult = await t.evaluate(`(async () => {
      const response = (event) => { if (event.data?.type === 'vibeoffice-react-request-save') postMessage({ type: 'vibeoffice-react-save-result', requestId: event.data.requestId, result: 'saved' }, location.origin); };
      addEventListener('message', response); const result = await VO.reactFileSave('Recovered.docx', new Blob(['bytes'])); removeEventListener('message', response); return result;
    })()`);
    assert.equal(saveResult, 'saved', 'save requests complete through the host response');
    const appSaveResult = await t.evaluate(`(async () => {
      const response = (event) => { if (event.data?.type === 'vibeoffice-react-request-save') postMessage({ type: 'vibeoffice-react-save-result', requestId: event.data.requestId, result: 'saved' }, location.origin); };
      addEventListener('message', response); const result = await L.app.saveAs(); removeEventListener('message', response); return result;
    })()`);
    assert.equal(appSaveResult, 'saved', 'Quire Save As writes through the React host adapter');
    t.clean();
  } finally { await t.close(); }
});
