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

test('embedded React editors retain real Ribbon pointer and Save keyboard workflows', { timeout: 90000 }, async () => {
  const previousBase = process.env.BASE;
  const previousCdp = process.env.CDP;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, [path.join(root, 'node_modules/vite/bin/vite.js'), '--config', 'tests/vite/host.vite.config.ts', '--host', '127.0.0.1', '--port', String(port), '--strictPort'], { cwd: root, stdio: 'ignore' });
  process.env.BASE = base;
  process.env.CDP ||= 'http://127.0.0.1:9250';
  let page;

  const waitFor = (expression, timeout = 15000) => page.evaluate(`new Promise((resolve) => { const end = Date.now() + ${timeout}; const poll = () => { const value = (${expression}); value ? resolve(true) : Date.now() > end ? resolve(false) : setTimeout(poll, 40); }; poll(); })`);
  const clickAt = async ({ x, y }) => {
    await page.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y });
    await page.send('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', buttons: 1, clickCount: 1 });
    await page.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', buttons: 0, clickCount: 1 });
  };

  try {
    for (let attempt = 0; attempt < 80 && !await fetch(`${base}/tests/vite/host-fixture.html`).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: `${base}/tests/vite/host-fixture.html`, ready: 'document.readyState === "complete"', timeout: 20000, bypassServiceWorker: true });
    await page.send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 800, deviceScaleFactor: 1, mobile: false });

    for (const app of ['wordx', 'excelx', 'pptx']) {
      if (app !== 'wordx') await page.evaluate(`(() => { const select = [...document.querySelectorAll('select')].find((item) => [...item.options].some((option) => option.value === ${JSON.stringify(app)})); select.value = ${JSON.stringify(app)}; select.dispatchEvent(new Event('change', { bubbles: true })); })()`);

      const ribbonReady = await waitFor(`window.L?.APP_ID === ${JSON.stringify(app)} && document.querySelector('.vibeoffice-react-ribbon .rb-tab[data-tab="insert"]')`);
      assert.ok(ribbonReady, `${app} mounts its embedded React editor and Ribbon: ${JSON.stringify({ errors: page.errors.slice(-4) })}`);
      await page.evaluate("window.VO?.look?.set?.('paper2016')");

      await page.evaluate(`[...document.querySelectorAll('button')].find((button) => button.textContent.includes('Open host-provider sample'))?.click()`);
      const opened = await waitFor(`document.querySelector('[aria-label="Host provider status"]')?.textContent.includes('Read sample')`);
      assert.ok(opened, `${app} opens a host-provider document before input: ${JSON.stringify({ status: await page.evaluate('document.querySelector(\'[aria-label="Host provider status"]\')?.textContent'), errors: page.errors.slice(-4) })}`);
      if (app === 'excelx') {
        const textImport = await waitFor(`document.querySelector('.excelx-text-import[open]')`);
        assert.ok(textImport, 'Excelx presents its text import wizard for the host-selected text file');
        await page.evaluate(`document.querySelector('.excelx-text-import__actions .primary')?.click()`);
        assert.ok(await waitFor(`!document.querySelector('.excelx-text-import[open]')`), 'the host-selected text document is committed to the workbook before keyboard Save');
      }

      const insertVisible = await waitFor(`(() => { const rect = document.querySelector('.vibeoffice-react-ribbon .rb-tab[data-tab="insert"]')?.getBoundingClientRect(); return !!rect && rect.width > 0 && rect.height > 0; })()`);
      const appearance = await page.evaluate(`(() => ({ global: window.VO?.look?.get?.(), attrs: [document.documentElement.dataset.look, document.documentElement.dataset.layout], root: window.__VIBE_OFFICE_DIRECT_HOST__?.root && { look: window.__VIBE_OFFICE_DIRECT_HOST__.root.dataset.look, layout: window.__VIBE_OFFICE_DIRECT_HOST__.root.dataset.layout }, editorLook: document.querySelector('.vibe-editor__chrome select[aria-label="Look"]')?.value, ribbon: document.querySelector('.vibeoffice-react-ribbon')?.getBoundingClientRect().toJSON(), errors: window.__capturedErrors || [] }))()`);
      assert.ok(insertVisible, `${app} renders the Insert tab before pointer input: ${JSON.stringify(appearance)}`);
      const insertPoint = await page.evaluate(`(() => { const rect = document.querySelector('.vibeoffice-react-ribbon .rb-tab[data-tab="insert"]').getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }; })()`);
      await clickAt(insertPoint);
      const selected = await page.evaluate(`document.querySelector('.vibeoffice-react-ribbon .rb-tab[aria-selected="true"]')?.dataset.tab`);
      const pointerState = await page.evaluate(`(() => ({ rect: document.querySelector('.vibeoffice-react-ribbon .rb-tab[data-tab="insert"]').getBoundingClientRect().toJSON(), hit: document.elementFromPoint(${insertPoint.x}, ${insertPoint.y})?.outerHTML.slice(0, 180), look: document.documentElement.dataset.look }))()`);
      assert.equal(selected, 'insert', `${app} selects the React Ribbon tab through real pointer input: ${JSON.stringify({ selected, pointerState, errors: page.errors.slice(-4) })}`);

      const editorPoint = await page.evaluate(`(() => { const rect = document.querySelector('.vibe-editor__engine').getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }; })()`);
      await clickAt(editorPoint);
      for (const [type, key, code, modifiers] of [
        ['keyDown', 'Control', 'ControlLeft', 2],
        ['keyDown', 's', 'KeyS', 2],
        ['keyUp', 's', 'KeyS', 2],
        ['keyUp', 'Control', 'ControlLeft', 0],
      ]) await page.send('Input.dispatchKeyEvent', { type, key, code, modifiers, ...(key === 's' ? { windowsVirtualKeyCode: 83, nativeVirtualKeyCode: 83 } : {}) });

      const saved = await waitFor(`document.querySelector('[aria-label="Host provider status"]')?.textContent.startsWith('Saved')`, 5000);
      assert.ok(saved, `${app} routes real Ctrl+S input through the embedded host provider: ${JSON.stringify({ status: await page.evaluate('document.querySelector(\'[aria-label="Host provider status"]\')?.textContent'), errors: page.errors.slice(-4) })}`);
    }

    assert.deepEqual(page.errors, []);
    assert.deepEqual(page.requests, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
    if (previousCdp === undefined) delete process.env.CDP; else process.env.CDP = previousCdp;
  }
});
