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

test('WebDAV listing resolves relative hrefs against the requested collection and removes its own response', { timeout: 30000 }, async () => {
  const previousBase = process.env.BASE;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, [path.join(root, 'node_modules/vite/bin/vite.js'), '--config', 'tests/vite/host.vite.config.ts', '--host', '127.0.0.1', '--port', String(port), '--strictPort'], { cwd: root, stdio: 'ignore' });
  process.env.BASE = base;
  let page;
  try {
    for (let attempt = 0; attempt < 80 && !await fetch(`${base}/tests/vite/webdav-list-fixture.html`).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: `${base}/tests/vite/webdav-list-fixture.html`, ready: "document.querySelector('#result')?.textContent !== 'waiting'", timeout: 15000, bypassServiceWorker: true });
    const files = JSON.parse(await page.evaluate("document.querySelector('#result')?.textContent || 'null'"));
    assert.deepEqual(files, [
      { id: '/docs/2026/reports/annual%20summary.xlsx', name: 'Annual Summary', isDirectory: false, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', revision: { value: '"rev-1"', kind: 'etag' } },
      { id: '/docs/2026/notes.docx', name: 'Notes', isDirectory: false },
    ]);
    assert.deepEqual(page.errors, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});
