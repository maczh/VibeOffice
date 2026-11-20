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

test('Wordx React Ribbon switches tabs through real pointer input', { timeout: 60000 }, async () => {
  const previousBase = process.env.BASE;
  const previousCdp = process.env.CDP;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'suite/dist')], { stdio: 'ignore' });
  process.env.BASE = base;
  let page;
  try {
    for (let attempt = 0; attempt < 40 && !await fetch(base).then((response) => response.ok).catch(() => false); attempt++) {
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    page = await openPage(undefined, { bypassServiceWorker: true });
    await page.send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 800, deviceScaleFactor: 1, mobile: false });
    const startReady = await page.evaluate("new Promise((resolve) => { const end=Date.now()+6000; const poll=()=>document.querySelector('.suite-start')?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,25); poll(); })");
    assert.equal(startReady, true, `React Start Center loaded: ${JSON.stringify({ errors: page.errors, requests: page.requests })}`);

    await page.evaluate("document.querySelector('.suite-start .app button')?.click()");
    const ribbonReady = await page.evaluate("new Promise((resolve) => { const end=Date.now()+12000; const poll=()=>document.querySelector('.vibeoffice-react-ribbon .rb-tab[data-tab=insert]')?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,25); poll(); })");
    assert.equal(ribbonReady, true, `Wordx React Ribbon mounted: ${JSON.stringify({ errors: page.errors })}`);

    const switchedLook = await page.evaluate("(() => { const select=document.querySelector('.vibe-editor__chrome select[aria-label=Look]'); if(!select) return false; select.value='paper2016'; select.dispatchEvent(new Event('change',{bubbles:true})); return true; })()");
    assert.equal(switchedLook, true, 'the editor exposes its look selector');
    const insert = await page.evaluate("new Promise((resolve) => { const end=Date.now()+2500; const poll=()=>{ const button=document.querySelector('.vibeoffice-react-ribbon .rb-tab[data-tab=insert]'); const rect=button?.getBoundingClientRect(); if(rect?.width>0&&rect.height>0) resolve({x:rect.left+rect.width/2,y:rect.top+rect.height/2}); else if(Date.now()>end) resolve(false); else setTimeout(poll,20); }; poll(); })");
    assert.ok(insert, `the Insert tab is visible for pointer input: ${JSON.stringify({ errors: page.errors })}`);

    await page.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: insert.x, y: insert.y });
    await page.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: insert.x, y: insert.y, button: 'left', buttons: 1, clickCount: 1 });
    await page.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: insert.x, y: insert.y, button: 'left', buttons: 0, clickCount: 1 });
    const selected = await page.evaluate("new Promise((resolve) => { const end=Date.now()+1500; const poll=()=>{ const tab=document.querySelector('.vibeoffice-react-ribbon .rb-tab[aria-selected=true]')?.dataset.tab; const commands=document.querySelectorAll('.vibeoffice-react-ribbon .rb-panel [data-cmd]').length; if(tab==='insert'&&commands>0) resolve({tab,commands}); else if(Date.now()>end) resolve({tab,commands}); else setTimeout(poll,15); }; poll(); })");
    assert.equal(selected.tab, 'insert', `real pointer input selects Insert: ${JSON.stringify(selected)}`);
    assert.ok(selected.commands > 0, `the Insert tab renders its command controls: ${JSON.stringify(selected)}`);
    assert.deepEqual(page.errors, []);
    assert.deepEqual(page.requests, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
    if (previousCdp === undefined) delete process.env.CDP; else process.env.CDP = previousCdp;
  }
});
