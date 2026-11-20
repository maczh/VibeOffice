import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage } from '../../tools/ooxml/cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const editors = [
  { id: 'wordx', extension: '.docx' },
  { id: 'excelx', extension: '.xlsx' },
  { id: 'pptx', extension: '.pptx' },
];

async function unusedPort() {
  const probe = net.createServer();
  await new Promise((resolve) => probe.listen(0, '127.0.0.1', resolve));
  const port = probe.address().port;
  await new Promise((resolve) => probe.close(resolve));
  return port;
}

test('standalone React editors retain real Ribbon pointer workflows and Wordx/Excelx local-save shortcuts', { timeout: 90000 }, async () => {
  const previousBase = process.env.BASE;
  const previousCdp = process.env.CDP;

  try {
  for (const editor of editors) {
    const port = await unusedPort();
    const base = `http://127.0.0.1:${port}`;
    const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, editor.id, 'standalone')], { stdio: 'ignore' });
    process.env.BASE = base;
    process.env.CDP ||= 'http://127.0.0.1:9250';
    let page;
    const waitFor = (expression, timeout = 15000) => page.evaluate(`new Promise((resolve) => { const end = Date.now() + ${timeout}; const poll = () => { const value = (${expression}); value ? resolve(true) : Date.now() > end ? resolve(false) : setTimeout(poll, 40); }; poll(); })`);

    try {
      for (let attempt = 0; attempt < 80 && !await fetch(base).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
      page = await openPage(undefined, { url: `${base}/?new`, ready: 'document.readyState === "complete"', timeout: 40000, bypassServiceWorker: true });
      await page.send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 800, deviceScaleFactor: 1, mobile: false });
      const ready = await waitFor(`document.querySelector('.vibe-editor__engine #app') && window.L?.APP_ID === ${JSON.stringify(editor.id)}`);
      assert.ok(ready, `${editor.id} standalone app mounts: ${JSON.stringify({ errors: page.errors.slice(-4) })}`);

      await page.evaluate("window.VO?.look?.set?.('paper2016')");
      const insertVisible = await waitFor(`(() => { const rect = document.querySelector('.vibeoffice-react-ribbon .rb-tab[data-tab="insert"]')?.getBoundingClientRect(); return !!rect && rect.width > 0 && rect.height > 0; })()`);
      assert.ok(insertVisible, `${editor.id} standalone renders the Paper 2016 Ribbon`);
      const point = await page.evaluate(`(() => { const rect = document.querySelector('.vibeoffice-react-ribbon .rb-tab[data-tab="insert"]').getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }; })()`);
      await page.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: point.x, y: point.y });
      await page.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'left', buttons: 1, clickCount: 1 });
      await page.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x, y: point.y, button: 'left', buttons: 0, clickCount: 1 });
      assert.equal(await page.evaluate("document.querySelector('.vibeoffice-react-ribbon .rb-tab[aria-selected=true]')?.dataset.tab"), 'insert', `${editor.id} selects the React Ribbon tab through real pointer input`);

      if (editor.id !== 'pptx') {
        await page.evaluate(`(() => { window.__standaloneSaved = null; window.__standalonePickerCalls = 0; Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: async ({ suggestedName }) => { window.__standalonePickerCalls++; return { async createWritable() { return { async write(blob) { window.__standaloneSaved = { name: suggestedName, size: blob.size, signature: Array.from(new Uint8Array(await blob.slice(0, 4).arrayBuffer())) }; }, async close() {}, async abort() {} }; }, async getFile() { return new File([new Uint8Array([1])], suggestedName); } }; } }); })()`);
      }
      const editorPoint = await page.evaluate(`(() => { const rect = document.querySelector('.vibe-editor__engine').getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }; })()`);
      await page.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: editorPoint.x, y: editorPoint.y, button: 'left', buttons: 1, clickCount: 1 });
      await page.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: editorPoint.x, y: editorPoint.y, button: 'left', buttons: 0, clickCount: 1 });
      if (editor.id !== 'pptx') {
        for (const [type, key, code, modifiers] of [
          ['keyDown', 'Control', 'ControlLeft', 2],
          ['keyDown', 's', 'KeyS', 2],
          ['keyUp', 's', 'KeyS', 2],
          ['keyUp', 'Control', 'ControlLeft', 0],
        ]) await page.send('Input.dispatchKeyEvent', { type, key, code, modifiers, ...(key === 's' ? { windowsVirtualKeyCode: 83, nativeVirtualKeyCode: 83 } : {}) });

        const saved = await waitFor('window.__standaloneSaved && window.__standaloneSaved.size > 100', 12000);
        const result = await page.evaluate('window.__standaloneSaved');
        assert.ok(saved && result.name.endsWith(editor.extension) && result.signature.join(',') === '80,75,3,4', `${editor.id} saves an Office package through real Ctrl+S input: ${JSON.stringify({ result, pickerCalls: await page.evaluate('window.__standalonePickerCalls'), app: await page.evaluate('({id:L?.APP_ID,file:L?.app?.fileName,dirty:L?.app?.dirty,command:!!L?.ui?.cmds?.save})'), errors: page.errors.slice(-4) })}`);
      }
      assert.deepEqual(page.errors, []);
      assert.deepEqual(page.requests, []);
    } finally {
      if (page) await page.close();
      server.kill();
    }
  }
  } finally {
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
    if (previousCdp === undefined) delete process.env.CDP; else process.env.CDP = previousCdp;
  }
});
