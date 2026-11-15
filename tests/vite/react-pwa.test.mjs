import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage } from '../../tools/ooxml/cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

test('the production React suite and standalone editors start from the Vite service-worker cache offline', { timeout: 120000 }, async (t) => {
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
    page = await openPage(undefined, { url: origin + '/', timeout: 15000 });
    await page.send('Storage.clearDataForOrigin', { origin, storageTypes: 'all' });
    await page.evaluate("navigator.serviceWorker.register('/sw.js', { updateViaCache: 'none' })");
    await page.evaluate("new Promise((resolve,reject)=>{const end=Date.now()+30000;const poll=()=>navigator.serviceWorker.ready.then(()=>navigator.serviceWorker.controller?resolve(true):Date.now()>end?reject(new Error('Vite worker did not control the page')):setTimeout(poll,50));poll()})");
    const cached = await page.evaluate("caches.keys().then(async(keys)=>({keys,entries:await Promise.all(['/index.html','/wordx/index.html','/excelx/index.html','/pptx/index.html'].map(async(url)=>({url,hit:!!await caches.match(url)})))}))");
    assert.ok(cached.keys.some((key) => key.startsWith('vibeoffice-react-')), `the production Vite worker creates its own cache: ${JSON.stringify(cached.keys)}`);
    assert.deepEqual(cached.entries, ['/index.html', '/wordx/index.html', '/excelx/index.html', '/pptx/index.html'].map((url) => ({ url, hit: true })), 'Start Center and all independent React editor entries are precached');

    await page.send('Network.emulateNetworkConditions', { offline: true, latency: 0, downloadThroughput: 0, uploadThroughput: 0, connectionType: 'none' });
    for (const app of ['wordx', 'excelx', 'pptx']) {
      await page.send('Page.navigate', { url: `${origin}/${app}/` });
      const ready = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+15000;const poll=()=>window.L?.app?resolve({app:L.APP_ID,name:L.APP_NAME}):Date.now()>end?resolve(false):setTimeout(poll,50);poll()})");
      assert.deepEqual(ready, { app, name: app === 'pptx' ? 'PPTX' : app[0].toUpperCase() + app.slice(1) }, `${app} opens from cached production assets while offline`);
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
