import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage } from '../../tools/ooxml/cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const editors = [
  { id: 'wordx', ready: '!!window.L?.D?.doc?.main' },
  { id: 'excelx', ready: '!!window.L?.app?.books?.[window.L.app.cur]?.wb?.sheets?.length' },
  { id: 'pptx', ready: '!!window.L?.pres?.slides?.length' },
];

async function unusedPort() {
  const probe = net.createServer();
  await new Promise((resolve) => probe.listen(0, '127.0.0.1', resolve));
  const port = probe.address().port;
  await new Promise((resolve) => probe.close(resolve));
  return port;
}

test('React editor default surfaces omit Recent Files, preserve legacy data and never read its database', { timeout: 90000 }, async () => {
  const previousBase = process.env.BASE;
  try {
    for (const editor of editors) {
      const port = await unusedPort();
      const base = `http://127.0.0.1:${port}`;
      const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, editor.id, 'standalone')], { stdio: 'ignore' });
      process.env.BASE = base;
      let page;
      try {
        for (let n = 0; n < 40 && !await fetch(base).then((response) => response.ok).catch(() => false); n++) await new Promise((resolve) => setTimeout(resolve, 100));
        page = await openPage(undefined, {
          timeout: 15000,
          bypassServiceWorker: true,
          url: `${base}/icons/icon.svg`,
          ready: "document.readyState === 'complete'",
          preload: "(() => { const open=IDBFactory.prototype.open; window.__vibeofficeOriginalDbOpen=open; window.__vibeofficeDbOpens=[]; IDBFactory.prototype.open=function(name,...args){window.__vibeofficeDbOpens.push(String(name));return open.call(this,name,...args)} })()",
        });
        await page.evaluate(`new Promise((resolve,reject)=>{
          const request=indexedDB.open('vibeoffice',1);
          request.onupgradeneeded=()=>{request.result.createObjectStore('recent',{keyPath:'key'});request.result.createObjectStore('handoff',{keyPath:'id'});};
          request.onerror=()=>reject(request.error);
          request.onsuccess=()=>{const db=request.result,tx=db.transaction('recent','readwrite');tx.objectStore('recent').put({key:'wordx:pre-existing',app:'wordx',name:'Pre-existing.docx',time:1,file:new Blob(['existing'])});tx.oncomplete=()=>{db.close();window.__vibeofficeDbOpens=[];resolve(true)};tx.onerror=()=>reject(tx.error)};
        })`);
        const editorUrl = `${base}/?new=1&lang=en`;
        await page.send('Page.navigate', { url: editorUrl });
        await page.waitForNavigation(editorUrl);
        const state = await page.evaluate(`new Promise((resolve) => {
          const end=Date.now()+12000;
          const poll=()=>{
            if (${editor.ready}) {
              const host=document.querySelector('.vibeoffice-legacy-root');
              const productOpens=window.__vibeofficeDbOpens.slice();
              window.__vibeofficeDbOpens=[];
              const request=window.__vibeofficeOriginalDbOpen.call(indexedDB,'vibeoffice',1);
              request.onerror=()=>resolve({ready:true,dbReadError:String(request.error),productOpens,recentVisible:/Recent Files/i.test(host?.innerText||'')});
              request.onsuccess=()=>{const db=request.result,tx=db.transaction('recent','readonly'),get=tx.objectStore('recent').get('wordx:pre-existing');get.onsuccess=()=>{const preserved=!!get.result&&get.result.name==='Pre-existing.docx';db.close();window.L?.app?.menuBar?.openIndex?.(0);setTimeout(()=>{const menu=document.querySelector('.menu.vibeoffice-menu-layer[role=menu]');resolve({ready:true,recentVisible:/Recent Files/i.test(host?.innerText||''),fileMenuOpened:!!menu,fileMenuHasRecent:/Recent Files/i.test(menu?.innerText||''),productOpens,preserved})},80)}};
            } else if(Date.now()>end) resolve({ready:false,body:document.body.innerText.slice(0,500),opens:window.__vibeofficeDbOpens});
            else setTimeout(poll,50);
          };
          poll();
        })`);
        assert.equal(state.ready, true, `${editor.id} production editor started: ${JSON.stringify(state)}`);
        assert.equal(state.recentVisible, false, `${editor.id} default editor surface contains no Recent Files UI: ${JSON.stringify(state)}`);
        assert.equal(state.fileMenuOpened, true, `${editor.id} opens the React-owned File menu for the no-Recent-Files check: ${JSON.stringify(state)}`);
        assert.equal(state.fileMenuHasRecent, false, `${editor.id} File menu has no Recent Files entry: ${JSON.stringify(state)}`);
        assert.equal(state.preserved, true, `${editor.id} leaves pre-existing Recent Files data untouched: ${JSON.stringify(state)}`);
        assert.equal(state.productOpens.includes('vibeoffice'), false, `${editor.id} never opens the legacy Recent Files database: ${JSON.stringify(state)}`);
        assert.deepEqual(page.errors, [], `${editor.id} has no browser errors while omitting Recent Files`);
      } finally {
        await page?.close().catch(() => {});
        server.kill('SIGTERM');
      }
    }
  } finally {
    if (previousBase === undefined) delete process.env.BASE;
    else process.env.BASE = previousBase;
  }
});
