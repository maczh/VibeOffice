import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage } from '../../tools/ooxml/cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

async function unusedPort() {
  const probe = net.createServer();
  await new Promise((resolve) => probe.listen(0, '127.0.0.1', resolve));
  const port = probe.address().port;
  await new Promise((resolve) => probe.close(resolve));
  return port;
}

test('closing an old Wordx history menu does not remove callbacks installed by a newer menu', { timeout: 60000 }, async () => {
  const previousBase = process.env.BASE;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'wordx/standalone')], { stdio: 'ignore' });
  process.env.BASE = base;
  let page;
  try {
    for (let n = 0; n < 40 && !await fetch(base).then((response) => response.ok).catch(() => false); n++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { timeout: 15000, bypassServiceWorker: true });
    const ready = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+10000;const poll=()=>window.L?.app?.historyMenu?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,50);poll()})");
    assert.equal(ready, true, `Wordx production editor mounted: ${JSON.stringify({ errors: page.errors, requests: page.requests })}`);

    await page.evaluate(`(() => {
      window.__oldUndoList = L.D.undoList;
      L.D.undoList = () => ['Set font', 'Insert text'];
      L.app.historyMenu({ left: 80, bottom: 80 }, 'undo');
    })()`);
    const menuReady = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+2000;const poll=()=>document.querySelector('#react-wordx-history-menu')?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,20);poll()})");
    assert.equal(menuReady, true, 'the first React history menu mounted');

    const superseded = await page.evaluate(`(() => {
      const app=L.app;
      window.__newHistoryCatalog=()=>({kind:'undo',items:[{label:'Newer action',count:1}],initialStatus:'Undo 1 Action'});
      window.__newHistoryAction=()=>{window.__newHistoryActionCalled=true};
      app.reactHistoryMenuCatalog=window.__newHistoryCatalog;
      app.reactHistoryMenuAction=window.__newHistoryAction;
      L.ui.closeMenus();
      return true;
    })()`);
    assert.equal(superseded, true);
    const closed = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+2000;const poll=()=>!document.querySelector('#react-wordx-history-menu-root')?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,20);poll()})");
    assert.equal(closed, true, 'the old React menu unmounted');
    assert.deepEqual(await page.evaluate("({catalog:L.app.reactHistoryMenuCatalog===window.__newHistoryCatalog,action:L.app.reactHistoryMenuAction===window.__newHistoryAction,items:L.app.reactHistoryMenuCatalog?.().items.map(({label})=>label)})"), {
      catalog: true, action: true, items: ['Newer action'],
    }, 'old cleanup is identity-scoped and preserves a newer menu instance callback set');
    assert.deepEqual(page.errors, []);
    assert.deepEqual(page.requests, []);
  } finally {
    if (page) {
      await page.evaluate("(() => { if(window.__oldUndoList)L.D.undoList=window.__oldUndoList; if(window.__newHistoryCatalog&&L.app.reactHistoryMenuCatalog===window.__newHistoryCatalog)delete L.app.reactHistoryMenuCatalog; if(window.__newHistoryAction&&L.app.reactHistoryMenuAction===window.__newHistoryAction)delete L.app.reactHistoryMenuAction; })()").catch(() => {});
      await page.close();
    }
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});
