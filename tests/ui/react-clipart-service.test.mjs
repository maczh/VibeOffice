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

test('Wordx and PPTX install the React clip-art service without loading the legacy catalog script', { timeout: 60000 }, async () => {
  const previousBase = process.env.BASE;
  try {
    for (const app of ['wordx', 'pptx']) {
      const port = await unusedPort();
      const base = `http://127.0.0.1:${port}`;
      const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, app, 'standalone')], { stdio: 'ignore' });
      process.env.BASE = base;
      let page;
      try {
        for (let attempt = 0; attempt < 40 && !await fetch(base).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
        page = await openPage(undefined, { url: `${base}/`, timeout: 15000, bypassServiceWorker: true });
        const result = await page.evaluate(`new Promise((resolve)=>{const end=Date.now()+10000;const poll=()=>{if(window.L?.clipart){const all=L.clipart.search(''),match=L.clipart.search('people team'),none=L.clipart.search('not-a-clip');resolve({count:all.length,first:all[0]?.name,matched:match.map(x=>x.name),none:none.length,svgUrl:L.clipart.url(all[0]).startsWith('data:image/svg+xml;utf8,'),png:typeof L.clipart.toPNG})}else if(Date.now()>end)resolve(false);else setTimeout(poll,50)};poll()})`);
        assert.deepEqual(result, { count: 24, first: 'Light bulb', matched: ['Team'], none: 0, svgUrl: true, png: 'function' }, `${app} exposes the complete clip-art service`);
        const loaded = await page.evaluate("window.__VIBE_OFFICE_DIRECT_HOST__?.loadedLegacyScripts || []");
        assert.equal(loaded.some((script) => script.endsWith('/common/clipart.js')), false, `${app} skips the legacy clip-art script`);
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
