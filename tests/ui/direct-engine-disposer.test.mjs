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

test('direct-engine registers and invokes before-script service disposers exactly once', { timeout: 30000 }, async () => {
  const previousBase = process.env.BASE;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, [path.join(root, 'node_modules/vite/bin/vite.js'), '--config', 'tests/vite/host.vite.config.ts', '--host', '127.0.0.1', '--port', String(port), '--strictPort'], { cwd: root, stdio: 'ignore' });
  process.env.BASE = base;
  let page;
  try {
    for (let attempt = 0; attempt < 80 && !await fetch(`${base}/tests/vite/direct-engine-disposer.html`).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: `${base}/tests/vite/direct-engine-disposer.html`, ready: 'window.mountDisposerFixture !== undefined', timeout: 15000, bypassServiceWorker: true });
    const result = await page.evaluate(`(async() => {
      const host = document.querySelector('#engine-host');
      const dispose = await window.mountDisposerFixture(host);
      const afterMount = { calls: window.disposerFixtureCalls, mounted: !!host.querySelector('[data-vibeoffice-engine-root]') };
      dispose();
      dispose();
      return { afterMount, afterDispose: window.disposerFixtureCalls, rootRemoved: !host.querySelector('[data-vibeoffice-engine-root]') };
    })()`);
    assert.deepEqual(result, { afterMount: { calls: 1, mounted: true }, afterDispose: 2, rootRemoved: true });
    assert.deepEqual(page.errors, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});
