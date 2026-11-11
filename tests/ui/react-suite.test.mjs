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
    const waitForStart = async () => { const state = await page.evaluate("new Promise((resolve) => { const end=Date.now()+5000; const poll=()=>document.querySelector('.suite-start')?resolve({start:true}):Date.now()>end?resolve({start:false,body:document.body.innerText,dialogs:[...document.querySelectorAll('iframe')].map(f=>f.contentDocument?.body?.innerText.slice(-300))}):setTimeout(poll,50); poll(); })"); assert.equal(state.start, true, `React host confirmed the document close: ${JSON.stringify(state)}`); };
    const closeEditor = async () => {
      await page.evaluate("document.querySelector('.back').click()");
      const state = await page.evaluate("new Promise((resolve) => { const end=Date.now()+5000; const poll=()=>{ if(document.querySelector('.suite-start')) resolve('closed'); else { const doc=document.querySelector('iframe')?.contentDocument; if(doc?.querySelector('.dlg[role=dialog]')) resolve('prompt'); else if(Date.now()>end) resolve('timeout'); else setTimeout(poll,50); } }; poll(); })");
      if (state === 'prompt') await page.evaluate("[...document.querySelector('iframe').contentDocument.querySelectorAll('.dlg-foot button')].find((button)=>button.textContent.trim()==='No')?.click()");
      assert.notEqual(state, 'timeout', 'the embedded editor answers its close request');
      await waitForStart();
    };
    const ready = await page.evaluate(`new Promise((resolve) => { const end = Date.now() + 8000; const poll = () => document.querySelector('.suite-start') ? resolve(true) : Date.now() > end ? resolve(false) : setTimeout(poll, 50); poll(); })`);
    assert.equal(ready, true, 'the React Start Center rendered');
    assert.equal(await page.evaluate("document.querySelector('.suite-start h2')?.textContent"), 'Create File');
    assert.equal(await page.evaluate("document.body.innerText.includes('Recent Files')"), false);
    const workerState = await page.evaluate("({worker:!!navigator.serviceWorker, script:navigator.serviceWorker.controller?.scriptURL})");
    assert.ok(workerState.worker, 'the production suite exposes service-worker support');
    await page.evaluate("document.querySelectorAll('.app')[0].querySelector('button').click()");
    const chromeReady = await page.evaluate("new Promise((resolve) => { const end=Date.now()+10000; const poll=()=>{ const f=document.querySelector('iframe'); f?.contentWindow?.L?.app && f.contentWindow.L.D?.doc && document.querySelector('.vibe-editor__chrome') ? resolve(true) : Date.now()>end ? resolve(false) : setTimeout(poll,100); }; poll(); })");
    assert.equal(chromeReady, true, 'the standalone Quire package mounts its React chrome');
    await page.evaluate("(() => { const select=document.querySelector('.vibe-editor__chrome select[aria-label=Look]'); select.value='paper2016'; select.dispatchEvent(new Event('change',{bubbles:true})); })()");
    const lookSynced = await page.evaluate("new Promise((resolve) => { const end=Date.now()+3000; const poll=()=>{ const look=document.querySelector('iframe')?.contentDocument?.documentElement?.dataset?.look; look==='paper2016'?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,50); }; poll(); })");
    assert.equal(lookSynced, true, 'React look control synchronizes to the hosted engine');
    await page.evaluate("document.querySelector('.vibe-editor__chrome button').click()");
    const commandReady = await page.evaluate("new Promise((resolve) => { const end=Date.now()+3000; const poll=()=>{ const f=document.querySelector('iframe'); f?.contentWindow?.L?.app ? resolve(true) : Date.now()>end ? resolve(false) : setTimeout(poll,50); }; poll(); })");
    assert.equal(commandReady, true, 'React New command remains connected to the engine');
    const blankCreated = await page.evaluate("new Promise((resolve) => { const end=Date.now()+5000; const poll=()=>{ const f=document.querySelector('iframe'); const text=f?.contentWindow?.L?.D?.doc?.main?.blocks?.map((b)=>f.contentWindow.L.D.ptext(b)).join('') ?? null; text===''?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,50); }; poll(); })");
    assert.equal(blankCreated, true, 'React New creates a blank document directly, without the legacy template dialog');
    await closeEditor();
    const { root: documentNode } = await page.send('DOM.getDocument', { depth: -1 });
    const { nodeId: fileInput } = await page.send('DOM.querySelector', { nodeId: documentNode.nodeId, selector: 'input[type=file]' });
    assert.ok(fileInput, 'the Start Center exposes its local file picker');
    await page.send('DOM.setFileInputFiles', { nodeId: fileInput, files: [path.join(root, 'tests/vite/react-suite-sample.txt')] });
    const localOpened = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+10000; const poll=()=>{ const n=document.querySelector('iframe')?.contentWindow?.L?.app?.fileName||''; n.includes('react-suite-sample')?resolve(n):Date.now()>end?resolve(n):setTimeout(poll,100); }; poll(); })`);
    assert.match(localOpened, /react-suite-sample/);
    await page.evaluate("document.querySelector('.back').click()");
    const discardPrompt = await page.evaluate("new Promise((resolve) => { const end=Date.now()+3000; const poll=()=>{ const dlg=document.querySelector('iframe')?.contentDocument?.querySelector('.dlg[role=dialog]'); dlg?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,50); }; poll(); })");
    if (discardPrompt) await page.evaluate("[...document.querySelector('iframe').contentDocument.querySelectorAll('.dlg-foot button')].find((button)=>button.textContent.trim()==='No')?.click()");
    await waitForStart();
    for (const [index, label] of [[0, 'Quire editor'], [1, 'Ledger editor'], [2, 'Lectern editor']]) {
      await page.evaluate(`document.querySelectorAll('.app')[${index}].querySelector('button').click()`);
      const mounted = await page.evaluate(`new Promise((resolve) => { const end = Date.now() + 12000; const poll = () => { const frame = document.querySelector('iframe'); const ok = frame?.contentWindow?.L?.app && frame.contentDocument?.readyState === 'complete'; ok ? resolve(frame.title) : Date.now() > end ? resolve(false) : setTimeout(poll, 100); }; poll(); })`);
      assert.equal(mounted, label);
      assert.equal(await page.evaluate("!!document.querySelector('.vibe-editor__chrome')"), true, `each editor package mounts the shared React chrome (${await page.evaluate("document.body.innerText.slice(0,160)")})`);
      if (index === 0) {
        await page.evaluate(`new Promise((resolve, reject) => { const request = indexedDB.open('vibeoffice-react-drafts', 1); request.onupgradeneeded = () => request.result.createObjectStore('documents', {keyPath:'id'}); request.onerror = () => reject(request.error); request.onsuccess = () => { const db=request.result; const tx=db.transaction('documents','readwrite'); tx.objectStore('documents').put({id:'wordx:recovery-test',editor:'wordx',name:'Recovered.docx',updatedAt:Date.now(),snapshot:new Blob(['draft'])}); tx.oncomplete=()=>{db.close();resolve(true)}; tx.onerror=()=>reject(tx.error); }; })`);
        await closeEditor();
        await page.evaluate("document.querySelectorAll('.app')[0].querySelector('button').click()");
        const recoverable = await page.evaluate("new Promise((resolve) => { const end=Date.now()+8000; const poll=()=>{ const panel=document.querySelector('.draft-recovery'); panel ? resolve(panel.innerText) : Date.now()>end ? resolve(false) : setTimeout(poll,50); }; poll(); })");
        assert.match(recoverable, /Recovered\.docx/);
        const hasRecover = await page.evaluate("document.querySelector('.draft-recovery button')?.textContent");
        assert.equal(hasRecover, 'Recover');
      }
      await closeEditor();
    }
    assert.deepEqual(page.errors, []);
    assert.deepEqual(page.requests, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});
