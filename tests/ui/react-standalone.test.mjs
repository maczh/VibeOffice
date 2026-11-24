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

test('each editor production package runs independently from its standalone directory', { timeout: 60000 }, async () => {
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
        page = await openPage(undefined, { timeout: 15000, bypassServiceWorker: true });
        const mounted = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+10000; const poll=()=>document.querySelector('.vibe-editor__engine #app') && (${editor.ready}) ? resolve(true) : Date.now()>end ? resolve(false) : setTimeout(poll,50); poll(); })`);
        const diagnostic = mounted ? null : await page.evaluate("({url:location.href,title:document.title,body:document.body.innerText.slice(0,1000),editor:!!document.querySelector('.vibe-editor__engine #app'),scriptErrors:window.__capturedErrors||[]})");
        assert.equal(mounted, true, `${editor.id} standalone loads its React editor and document engine: ${JSON.stringify({ diagnostic, errors: page.errors, requests: page.requests })}`);
        const lookState = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+3000;const poll=()=>{const root=window.__VIBE_OFFICE_DIRECT_HOST__?.root,select=document.querySelector('.vibe-editor__chrome select[aria-label=Look]');if(window.VO?.look&&root?.isConnected&&select){const hostLook=document.documentElement.dataset.look;const paper=[...select.options].find((option)=>option.value==='paper2016');if(!paper)return resolve({missingPaper:true});select.value=paper.value;select.dispatchEvent(new Event('change',{bubbles:true}));return resolve({look:VO.look.get(),layout:VO.look.layout(),root:root.dataset.look,hostLookBefore:hostLook,hostLookAfter:document.documentElement.dataset.look})}if(Date.now()>end)return resolve(false);setTimeout(poll,30)};poll()})");
        assert.deepEqual({ look: lookState.look, layout: lookState.layout, root: lookState.root }, { look: 'paper2016', layout: 'ribbon', root: 'paper2016' }, `${editor.id} installs the shared look service and scopes look changes to its directly mounted editor`);
        assert.equal(lookState.hostLookAfter, lookState.hostLookBefore, `${editor.id} look changes do not mutate the host document's existing look`);
        assert.deepEqual(await page.evaluate(`window.__VIBE_OFFICE_DIRECT_HOST__?.loadedLegacyScripts.filter((path)=>path.startsWith('/${editor.id}/js/')) || []`), [], `${editor.id} standalone does not load the app-specific legacy JavaScript runtime`);
        assert.equal(await page.evaluate("document.querySelectorAll('iframe').length"), 0, `${editor.id} standalone mounts directly without an iframe`);
        if (editor.id === 'wordx') {
          await page.evaluate("(()=>{const d=L.D.newDoc();d.main.blocks=[L.D.para([L.D.text('Sep')])];L.app.opts.autoComplete=true;L.app.loadDoc(d,'Standalone AutoComplete');const p=L.D.firstPara(L.D.doc.main.blocks);L.ed.setSel(L.D.pos(p,L.D.plen(p)));L.ed.focus()})()");
          const caretReady = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+3000;const poll=()=>{L.ed.caretRect()?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,20)};poll()})");
          assert.equal(caretReady, true, 'standalone Wordx has a laid-out caret before typing');
          await page.send('Input.insertText', { text: 't' });
          const suggestion = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+1200;const poll=()=>{const tip=document.querySelector('[data-react-autocomplete=wordx]');tip?resolve(tip.textContent):Date.now()>end?resolve(false):setTimeout(poll,20)};poll()})");
          assert.equal(suggestion, 'September (Press ENTER to Insert)', `standalone Wordx renders its AutoComplete suggestion through the React package: ${JSON.stringify({ errors: page.errors })}`);
          await page.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13 });
          await page.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13 });
          const completed = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+1000;const poll=()=>{const text=L.D.ptext(L.D.firstPara(L.D.doc.main.blocks));text==='September'?resolve(text):Date.now()>end?resolve(text):setTimeout(poll,20)};poll()})");
          assert.equal(completed, 'September', 'standalone React Wordx accepts the suggestion using the existing document transaction');
        }
        assert.equal(await page.evaluate('location.pathname'), '/');
        assert.equal(await page.evaluate("document.querySelectorAll('iframe').length"), 0);
        assert.deepEqual(page.errors, []);
        assert.deepEqual(page.requests, []);
      } finally {
        if (page) await page.close();
        server.kill();
      }
    }
  } finally {
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});
