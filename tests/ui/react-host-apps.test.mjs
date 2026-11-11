import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open } from '../lib/input.mjs';

for (const [app, editor] of [['ledger', 'excelx'], ['lectern', 'pptx']]) {
  test(`${app} React-host mode routes Open and Save through the parent without Recent Files`, async () => {
    const t = await open(app, { persistence: true, query: { reactHost: '', new: '' } });
    try {
      assert.equal(await t.evaluate('VO.reactHost'), true);
      assert.deepEqual(await t.evaluate('VO.recent.list()'), []);
      assert.deepEqual(await t.evaluate('VO.recentMenuItems()'), []);
      const openRequested = await t.evaluate(`new Promise((resolve) => {
        const listener = (event) => { if (event.data?.type === 'vibeoffice-react-request-open') { removeEventListener('message', listener); resolve(true); } };
        addEventListener('message', listener); L.app.openDialog(); setTimeout(() => resolve(false), 700);
      })`);
      assert.equal(openRequested, true, `${app} Open requests the host to choose a file`);
      const saveRequested = await t.evaluate(`new Promise((resolve) => {
        const response = (event) => { if (event.data?.type === 'vibeoffice-react-request-save') { postMessage({ type: 'vibeoffice-react-save-result', requestId: event.data.requestId, result: 'saved' }, location.origin); removeEventListener('message', response); resolve(true); } };
        addEventListener('message', response); L.app.save(); setTimeout(() => resolve(false), 1500);
      })`);
      assert.equal(saveRequested, true, `${app} saves through the host adapter`);
    } finally { await t.close(); }
  });
}
