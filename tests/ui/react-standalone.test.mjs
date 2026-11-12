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
        page = await openPage(undefined, { timeout: 15000 });
        const mounted = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+10000; const poll=()=>document.querySelector('.vibe-editor__engine #app') && (${editor.ready}) ? resolve(true) : Date.now()>end ? resolve(false) : setTimeout(poll,50); poll(); })`);
        assert.equal(mounted, true, `${editor.id} standalone loads its React editor and document engine`);
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
