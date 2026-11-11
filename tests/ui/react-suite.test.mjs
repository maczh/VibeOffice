import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { openPage } from '../../tools/ooxml/cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

test('React suite starts at Create File and mounts each full editor package', { timeout: 60000 }, async () => {
  const probe = net.createServer();
  await new Promise((resolve) => probe.listen(0, '127.0.0.1', resolve));
  const port = probe.address().port;
  await new Promise((resolve) => probe.close(resolve));
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'suite/dist')], { stdio: 'ignore' });
  const previousBase = process.env.BASE;
  process.env.BASE = `http://127.0.0.1:${port}`;
  let page;
  try {
    for (let n = 0; n < 40; n++) {
      if (await fetch(process.env.BASE).then((response) => response.ok).catch(() => false)) break;
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    page = await openPage(undefined, { timeout: 15000 });
    const ready = await page.evaluate(`new Promise((resolve) => { const end = Date.now() + 8000; const poll = () => document.querySelector('.suite-start') ? resolve(true) : Date.now() > end ? resolve(false) : setTimeout(poll, 50); poll(); })`);
    assert.equal(ready, true, 'the React Start Center rendered');
    assert.equal(await page.evaluate("document.querySelector('.suite-start h2')?.textContent"), 'Create File');
    assert.equal(await page.evaluate("document.body.innerText.includes('Recent Files')"), false);
    const workerState = await page.evaluate("({worker:!!navigator.serviceWorker, script:navigator.serviceWorker.controller?.scriptURL})");
    assert.ok(workerState.worker, 'the production suite exposes service-worker support');
    await page.close();
    page = await openPage(undefined, { timeout: 15000 });
    await page.evaluate("new Promise((resolve) => { const end=Date.now()+8000; const poll=()=>document.querySelector('input[type=file]')?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,50); poll(); })");
    const fixture = await (await import('node:fs/promises')).readFile(path.join(root, 'tests/vite/react-suite-sample.txt'), 'utf8');
    await page.evaluate(`(() => { const input=document.querySelector('input[type=file]'); const file=new File([${JSON.stringify(fixture)}],'react-suite-sample.txt',{type:'text/plain'}); Object.defineProperty(input,'files',{configurable:true,value:[file]}); input.dispatchEvent(new Event('change',{bubbles:true})); })()`);
    const localOpened = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+10000; const poll=()=>{ const f=document.querySelector('iframe'); const n=f?.contentWindow?.L?.app?.fileName||''; n.includes('react-suite-sample')?resolve(n):Date.now()>end?resolve(n):setTimeout(poll,100); }; poll(); })`);
    assert.match(localOpened, /react-suite-sample/);
    await page.evaluate("document.querySelector('.back').click()");
    for (const [index, label] of [[0, 'Quire editor'], [1, 'Ledger editor'], [2, 'Lectern editor']]) {
      await page.evaluate(`document.querySelectorAll('.app')[${index}].querySelector('button').click()`);
      const mounted = await page.evaluate(`new Promise((resolve) => { const end = Date.now() + 12000; const poll = () => { const frame = document.querySelector('iframe'); const ok = frame?.contentWindow?.L?.app && frame.contentDocument?.readyState === 'complete'; ok ? resolve(frame.title) : Date.now() > end ? resolve(false) : setTimeout(poll, 100); }; poll(); })`);
      assert.equal(mounted, label);
      if (index === 0) {
        await page.evaluate(`new Promise((resolve, reject) => { const request = indexedDB.open('vibeoffice-react-drafts', 1); request.onupgradeneeded = () => request.result.createObjectStore('documents', {keyPath:'id'}); request.onerror = () => reject(request.error); request.onsuccess = () => { const db=request.result; const tx=db.transaction('documents','readwrite'); tx.objectStore('documents').put({id:'wordx:recovery-test',editor:'wordx',name:'Recovered.docx',updatedAt:Date.now(),snapshot:new Blob(['draft'])}); tx.oncomplete=()=>{db.close();resolve(true)}; tx.onerror=()=>reject(tx.error); }; })`);
        await page.evaluate("document.querySelector('.back').click()");
        await page.evaluate("document.querySelectorAll('.app')[0].querySelector('button').click()");
        const recoverable = await page.evaluate("new Promise((resolve) => { const end=Date.now()+8000; const poll=()=>{ const panel=document.querySelector('.draft-recovery'); panel ? resolve(panel.innerText) : Date.now()>end ? resolve(false) : setTimeout(poll,50); }; poll(); })");
        assert.match(recoverable, /Recovered\.docx/);
        const hasRecover = await page.evaluate("document.querySelector('.draft-recovery button')?.textContent");
        assert.equal(hasRecover, 'Recover');
      }
      await page.evaluate("document.querySelector('.back').click()");
    }
    assert.deepEqual(page.errors, []);
    assert.deepEqual(page.requests, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});
