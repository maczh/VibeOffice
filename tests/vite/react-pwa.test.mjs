import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage } from '../../tools/ooxml/cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

test('the production React suite can edit and save Office files entirely offline', { timeout: 120000 }, async (t) => {
  const probe = net.createServer();
  await new Promise((resolve) => probe.listen(0, '127.0.0.1', resolve));
  const port = probe.address().port;
  await new Promise((resolve) => probe.close(resolve));
  const origin = `http://127.0.0.1:${port}`;
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'suite/dist')], { stdio: 'ignore' });
  const previousBase = process.env.BASE;
  process.env.BASE = origin;
  let page;
  try {
    for (let attempt = 0; attempt < 80; attempt++) {
      if (await fetch(origin).then((response) => response.ok).catch(() => false)) break;
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    page = await openPage(undefined, { url: origin + '/?lang=en', timeout: 60000 });
    await page.send('Storage.clearDataForOrigin', { origin, storageTypes: 'all' });
    await page.evaluate("caches.open('legacy-vibeoffice-cache').then((cache)=>cache.put('/index.html',new Response('<!doctype html><title>stale legacy shell</title>',{headers:{'Content-Type':'text/html'}})))");
    await page.evaluate("navigator.serviceWorker.register('/sw.js', { updateViaCache: 'none' })");
    await page.evaluate("new Promise((resolve,reject)=>{const end=Date.now()+30000;const poll=()=>navigator.serviceWorker.ready.then(()=>navigator.serviceWorker.controller?resolve(true):Date.now()>end?reject(new Error('Vite worker did not control the page')):setTimeout(poll,50));poll()})");
    const cached = await page.evaluate("caches.keys().then(async(keys)=>({keys,entries:await Promise.all(['/index.html','/wordx/index.html','/excelx/index.html','/pptx/index.html'].map(async(url)=>({url,hit:!!await caches.match(url)})))}))");
    assert.ok(cached.keys.some((key) => key.startsWith('vibeoffice-react-')), `the production Vite worker creates its own cache: ${JSON.stringify(cached.keys)}`);
    assert.deepEqual(cached.entries, ['/index.html', '/wordx/index.html', '/excelx/index.html', '/pptx/index.html'].map((url) => ({ url, hit: true })), 'Start Center and all independent React editor entries are precached');
    assert.equal(await page.evaluate("caches.open('legacy-vibeoffice-cache').then((cache)=>cache.match('/index.html').then((response)=>response.text()))"), '<!doctype html><title>stale legacy shell</title>', 'the independent legacy cache remains untouched during React worker activation');
    const shortcutTargets = await page.evaluate("fetch('/manifest.webmanifest').then((response)=>response.json()).then((manifest)=>manifest.shortcuts.map((shortcut)=>{const target=new URL(shortcut.url,location.origin+'/manifest.webmanifest');return target.pathname+target.search}))");
    assert.deepEqual(shortcutTargets, ['/?new=wordx', '/?new=excelx', '/?new=pptx'], 'installed app shortcuts resolve to the React Start Center deep links rather than legacy editor pages');

    await page.send('Network.emulateNetworkConditions', { offline: true, latency: 0, downloadThroughput: 0, uploadThroughput: 0, connectionType: 'none' });
    await page.send('Page.navigate', { url: origin + '/?lang=en' });
    const offlineStart = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+10000;const poll=()=>document.querySelector('.suite-start')?resolve({title:document.title,text:document.body.innerText}):Date.now()>end?resolve(false):setTimeout(poll,30);poll()})");
    assert.ok(offlineStart && !offlineStart.title.includes('stale legacy shell') && offlineStart.text.includes('Create File'), JSON.stringify(offlineStart));
    for (const [index, app] of ['wordx', 'excelx', 'pptx'].entries()) {
      if (!index) {
        await page.send('Page.navigate', { url: new URL(`./?new=${app}`, `${origin}/manifest.webmanifest`).href });
      } else {
        await page.evaluate("document.querySelector('.suite-editor > .back')?.click()");
        const closeState = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+5000;const poll=()=>document.querySelector('.suite-start .app .cards button')?resolve('start'):document.querySelector('dialog:modal,.dlg[role=dialog]')?resolve('prompt'):Date.now()>end?resolve('timeout'):setTimeout(poll,30);poll()})");
        if (closeState === 'prompt') await page.evaluate("[...document.querySelectorAll('dialog:modal button,.dlg-foot button')].find((button)=>/^(?:No|否|Discard|Don’t Save|Don't Save)/.test(button.textContent.trim()))?.click()");
        const returned = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+8000;const poll=()=>document.querySelector('.suite-start .app .cards button')?resolve(true):Date.now()>end?resolve({body:document.body.innerText.slice(-500),dialogs:[...document.querySelectorAll('dialog:modal,.dlg[role=dialog]')].map((dialog)=>dialog.innerText)}):setTimeout(poll,30);poll()})");
        assert.equal(returned, true, `the React Start Center returns without reloading the offline page after close state ${closeState}: ${JSON.stringify(returned)}`);
      }
      if (index) {
        const startReady = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+10000;const poll=()=>document.querySelector('.suite-start .app .cards button')?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,30);poll()})");
        assert.equal(startReady, true, 'the cached React Start Center returns while offline');
        t.diagnostic(`${app}: Start Center available offline`);
        await page.evaluate(`(() => [...document.querySelectorAll('.suite-start .app')].find((section)=>section.querySelector('.suite-app-mark.${app}'))?.querySelector('.cards button')?.click())()`);
      }
      const documentReady = app === 'wordx' ? '!!L.D?.doc?.main' : app === 'excelx' ? '!!L.grid?.wb?.sheets?.length' : '!!L.pres?.slides?.length';
      const ready = await page.evaluate(`new Promise((resolve)=>{const end=Date.now()+15000;const poll=()=>document.querySelector('.vibe-editor__engine #app')&&window.L?.APP_ID==='${app}'&&window.L?.app&&(${documentReady})?resolve({app:L.APP_ID,name:L.APP_NAME}):Date.now()>end?resolve(false):setTimeout(poll,50);poll()})`);
      const diagnostic = ready ? null : await page.evaluate("({url:location.href,start:!!document.querySelector('.suite-start'),body:document.body.innerText.slice(0,600),host:!!document.querySelector('.vibe-editor__engine #app'),app:window.L?.APP_ID,requests:performance.getEntriesByType('resource').slice(-20).map((entry)=>entry.name)})");
      assert.deepEqual(ready, { app, name: app === 'pptx' ? 'PPTX' : app[0].toUpperCase() + app.slice(1) }, `${app} React package opens from cached production assets while offline: ${JSON.stringify({ diagnostic, errors: page.errors })}`);
      assert.equal(await page.evaluate('location.search'), '', `${app} consumes its PWA shortcut query before mounting the editor`);
      t.diagnostic(`${app}: React editor mounted offline`);

      const edited = app === 'wordx'
        ? await page.evaluate("(()=>{const paragraph=L.D.firstPara(L.D.doc.main.blocks);L.ed.edit('Offline edit',()=>L.O.insertText(L.D.pos(paragraph,L.D.plen(paragraph)),' offline',{offline:true}));return L.D.ptext(paragraph).includes('offline')})()")
        : app === 'excelx'
          ? await page.evaluate("(()=>{L.ops.tx(L.grid.wb,'Offline edit',()=>L.ops.enter(L.grid.sheet(),0,0,'Offline edit'));return L.editor.cellText(L.grid.sheet(),0,0)==='Offline edit'})()")
          : await page.evaluate("(()=>{const before=L.pres.slides.length;L.ui.exec('newSlide');return L.pres.slides.length===before+1})()");
      assert.equal(edited, true, `${app} can edit its document model while offline`);
      t.diagnostic(`${app}: document model edited offline`);
      await page.evaluate(`(() => {
        window.__offlineSaved = null;
        window.__offlinePickerCalls = 0;
        Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: async ({ suggestedName }) => {
          window.__offlinePickerCalls++;
          return {
            async createWritable() { return {
              async write(blob) { window.__offlineSaved = { name: suggestedName, size: blob.size, type: blob.type, signature: Array.from(new Uint8Array(await blob.slice(0, 4).arrayBuffer())) }; },
              async close() {}, async abort() {},
            }; },
            async getFile() { return new File([new Uint8Array([1])], suggestedName); },
          };
        } });
      })()`);
      const writer = app === 'wordx'
        ? "L.docx.write(L.D.doc,{stats:L.app.docStats()})"
        : app === 'excelx'
          ? "L.xlsxWrite.write(L.grid.wb,{type:'xlsx'})"
          : "L.pptx.write(L.pres,{format:'pptx'})";
      const name = app === 'wordx' ? 'Offline.docx' : app === 'excelx' ? 'Offline.xlsx' : 'Offline.pptx';
      await page.evaluate(`(() => { window.__offlineSaveState='writing'; window.__offlineSavePromise=${writer}.then(async(blob)=>{window.__offlineBlob={size:blob.size,type:blob.type};window.__offlineSaveState=await VO.reactFileSave('${name}',blob)}).catch((error)=>{window.__offlineSaveState={error:String(error),stack:error?.stack}});return true; })()`);
      t.diagnostic(`${app}: OOXML writer started`);
      let saved = 'writing';
      for (let attempt = 0; attempt < 240 && saved === 'writing'; attempt++) {
        await new Promise((resolve) => setTimeout(resolve, 50));
        saved = await page.evaluate('window.__offlineSaveState');
      }
      assert.equal(saved, 'saved', `${app} saves the edited OOXML document through its local file picker while offline: ${JSON.stringify({ saved, pickerCalls: await page.evaluate('window.__offlinePickerCalls'), blob: await page.evaluate('window.__offlineBlob') })}`);
      t.diagnostic(`${app}: local file save completed offline`);
      const localWrite = await page.evaluate('window.__offlineSaved');
      assert.ok(localWrite && localWrite.name.endsWith(app === 'wordx' ? '.docx' : app === 'excelx' ? '.xlsx' : '.pptx') && localWrite.size > 100 && localWrite.signature.join(',') === '80,75,3,4', `${app} writes a non-empty ZIP-based Office document to local storage: ${JSON.stringify({ localWrite, pickerCalls: await page.evaluate('window.__offlinePickerCalls') })}`);
    }
  } finally {
    if (page) {
      await page.send('Network.emulateNetworkConditions', { offline: false, latency: 0, downloadThroughput: -1, uploadThroughput: -1 }).catch(() => {});
      await page.send('Storage.clearDataForOrigin', { origin, storageTypes: 'all' }).catch(() => {});
      await page.close();
    }
    server.kill();
    if (previousBase === undefined) delete process.env.BASE;
    else process.env.BASE = previousBase;
  }
});
