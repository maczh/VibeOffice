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

test('React workspace settings and branding never call or preserve a desktop bridge', { timeout: 60000 }, async () => {
  const previousBase = process.env.BASE;
  const previousCdp = process.env.CDP;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, [path.join(root, 'node_modules/vite/bin/vite.js'), '--config', 'tests/vite/host.vite.config.ts', '--host', '127.0.0.1', '--port', String(port), '--strictPort'], { cwd: root, stdio: 'ignore' });
  process.env.BASE = base;
  process.env.CDP ||= 'http://127.0.0.1:9223';
  let page;
  try {
    for (let attempt = 0; attempt < 80 && !await fetch(`${base}/tests/vite/runtime-services-web-only.html`).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: `${base}/tests/vite/runtime-services-web-only.html`, ready: 'window.checkWorkspaceRuntimeServices !== undefined', timeout: 20000, bypassServiceWorker: true });
    const result = await page.evaluate('window.checkWorkspaceRuntimeServices()');
    assert.deepEqual(result, {
      settings: { stored: true, value: 'browser-storage', look: 'classic', desktopInvocations: 0, desktopSettingsUntouched: true },
      brand: 'Welcome to Wordx / Excelx / PPTX',
    });
    assert.deepEqual(page.errors, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
    if (previousCdp === undefined) delete process.env.CDP; else process.env.CDP = previousCdp;
  }
});
