import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open } from '../lib/input.mjs';

for (const app of ['ledger', 'lectern']) {
  test(`${app} standalone ignores the obsolete iframe-host query`, async () => {
    const page = await open(app, { query: { reactHost: '', new: '' } });
    try {
      assert.equal(await page.evaluate('VO.reactHost'), false, 'a query parameter cannot activate a React host bridge');
      assert.equal(await page.evaluate('VO.reactDirectHost'), null);
      assert.equal(await page.evaluate("document.querySelectorAll('iframe').length"), 0, 'standalone editor contains no nested browsing context');
      const expectedId = app === 'ledger' ? 'excelx' : 'pptx', expectedName = app === 'ledger' ? 'Excelx' : 'PPTX';
      assert.deepEqual(await page.evaluate('({ root: document.documentElement.dataset.app, id: L.APP_ID, name: L.APP_NAME, app: L.APP })'), { root: expectedId, id: expectedId, name: expectedName, app: expectedName }, 'standalone browser identity uses the new web product namespace');
      assert.equal(await page.evaluate('!!L.app'), true, 'the classic standalone editor still initializes');
    } finally { await page.close(); }
  });
}
