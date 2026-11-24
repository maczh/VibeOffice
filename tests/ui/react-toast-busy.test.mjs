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

test('React Busy bridge emits transition-only host notifications and closes its overlay', { timeout: 60000 }, async () => {
  const previousBase = process.env.BASE;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'wordx/standalone')], { stdio: 'ignore' });
  process.env.BASE = base;
  let page;
  try {
    for (let attempt = 0; attempt < 40 && !await fetch(base).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: `${base}/`, timeout: 15000, bypassServiceWorker: true });
    const ready = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+10000;const poll=()=>window.L?.D?.doc?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,30);poll()})");
    assert.equal(ready, true, `Wordx standalone is ready: ${JSON.stringify({ errors: page.errors, requests: page.requests })}`);
    const signals = await page.evaluate("(()=>{const signals=[];window.addEventListener('vo-busy',event=>signals.push(event.detail));window.L.ui.busy(true,'Working test');window.L.ui.busy(true,'Repeated on');window.L.ui.busy(false);window.L.ui.busy(false);return signals})()");
    assert.deepEqual(signals, [{ on: true, text: 'Working test' }, { on: false, text: '' }]);
    const closed = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+1000;const poll=()=>!document.querySelector('.vibeoffice-busy')?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,20);poll()})");
    assert.equal(closed, true, 'the React busy overlay is removed when the engine finishes');
    assert.deepEqual(page.errors, []);
    assert.deepEqual(page.requests, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});
