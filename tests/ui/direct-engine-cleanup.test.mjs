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

test('direct engine disposal restores host state after a cleanup action throws', { timeout: 60000 }, async () => {
  const previousBase = process.env.BASE;
  const previousCdp = process.env.CDP;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, [path.join(root, 'node_modules/vite/bin/vite.js'), '--config', 'tests/vite/host.vite.config.ts', '--host', '127.0.0.1', '--port', String(port), '--strictPort'], { cwd: root, stdio: 'ignore' });
  process.env.BASE = base;
  process.env.CDP ||= 'http://127.0.0.1:9223';
  let page;
  try {
    for (let attempt = 0; attempt < 80 && !await fetch(`${base}/tests/vite/direct-engine-cleanup.html`).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: `${base}/tests/vite/direct-engine-cleanup.html`, ready: 'window.mountCleanupFixture !== undefined', timeout: 20000, bypassServiceWorker: true });
    const result = await page.evaluate(`(async()=>{
      const host=document.querySelector('#engine-host');
      const hostApp={name:'embedding app'};
      Object.defineProperty(window,'L',{configurable:true,enumerable:true,writable:true,value:hostApp});
      const reports=[];
      window.reportError=(error)=>reports.push({name:error.name,message:error.message,errors:error.errors?.map((item)=>item.message)});
      const dispose=await window.mountCleanupFixture(host);
      const root=host.querySelector('[data-vibeoffice-engine-root]');
      const style=root.querySelector('style');
      style.remove=()=>{throw new Error('simulated style cleanup failure')};
      const mounted={root:!!root,engineOwnsNamespace:window.L!==hostApp};
      dispose();
      const afterFirst={root:!!host.querySelector('[data-vibeoffice-engine-root]'),hostNamespaceRestored:window.L===hostApp,directRootCleared:window.__VIBE_OFFICE_DIRECT_HOST__?.root===null,reports:reports.slice()};
      dispose();
      return {mounted,afterFirst,reportsAfterSecondDispose:reports.length};
    })()`);
    assert.deepEqual(result, {
      mounted: { root: true, engineOwnsNamespace: true },
      afterFirst: {
        root: false,
        hostNamespaceRestored: true,
        directRootCleared: true,
        reports: [{ name: 'AggregateError', message: 'One or more VibeOffice editor cleanup actions failed', errors: ['simulated style cleanup failure'] }],
      },
      reportsAfterSecondDispose: 1,
    });
    assert.deepEqual(page.errors, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
    if (previousCdp === undefined) delete process.env.CDP; else process.env.CDP = previousCdp;
  }
});
