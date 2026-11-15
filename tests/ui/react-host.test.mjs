import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open } from '../lib/input.mjs';

test('legacy app keyboard handlers ignore events outside the mounted editor host', async () => {
  for (const app of ['quire', 'ledger', 'lectern']) {
    const page = await open(app, { preload: `document.addEventListener('DOMContentLoaded', () => {
      const host = document.createElement('div'); host.dataset.vibeofficeEngineRoot = '';
      while (document.body.firstChild) host.append(document.body.firstChild);
      document.body.append(host);
    });` });
    try {
      const calls = await page.evaluate(`(() => {
        const outside = document.createElement('button'); document.body.append(outside);
        let keyboard = 0, clipboard = 0, tableEscape = 0, textDrag = 0, extendEscape = 0, drawEscape = 0;
        const originalMenu = L.app.menuBar.openIndex;
        L.app.menuBar.openIndex = () => { keyboard++; };
        outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'F10', bubbles: true, cancelable: true }));
        L.app.menuBar.openIndex = originalMenu;
        if ('${app}' === 'quire') {
          const outsideEditor = document.createElement('div'); outsideEditor.className = 'ed-host';
          const paragraph = document.createElement('div'); paragraph.className = 'p'; outsideEditor.append(paragraph); document.body.append(outsideEditor);
          const originalDrag = L.ui.drag;
          L.ui.drag = (...args) => { textDrag++; return originalDrag(...args); };
          paragraph.dispatchEvent(new PointerEvent('pointerdown', { button: 0, bubbles: true, cancelable: true }));
          L.ui.drag = originalDrag; outsideEditor.remove();
          const originalSetTool = L.tables.setTool, originalTool = L.tables.tool, originalStatus = L.app.updateStatus;
          L.ed.extend = true; L.app.updateStatus = () => { extendEscape++; };
          outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
          L.app.updateStatus = originalStatus; L.ed.extend = false;
          const originalStop = L.drawtool.stop; L.drawtool.active = {};
          L.drawtool.stop = () => { drawEscape++; };
          outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
          L.drawtool.stop = originalStop; L.drawtool.active = null;
          L.tables.tool = 'draw'; L.tables.setTool = () => { tableEscape++; };
          outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
          L.tables.setTool = originalSetTool; L.tables.tool = originalTool;
        }
        if ('${app}' === 'lectern') {
          const originalCopy = L.app.copySlides; L.app.focusArea = 'slides';
          L.app.copySlides = () => { clipboard++; return null; };
          outside.dispatchEvent(new ClipboardEvent('copy', { bubbles: true, cancelable: true }));
          L.app.copySlides = originalCopy;
        }
        outside.remove(); return { keyboard, clipboard, tableEscape, textDrag, extendEscape, drawEscape };
      })()`);
      assert.equal(calls.keyboard, 0, `${app} must ignore keyboard input outside its mount`);
      if (app === 'lectern') assert.equal(calls.clipboard, 0, 'Lectern ignores clipboard input outside its mount');
      if (app === 'quire') {
        assert.equal(calls.tableEscape, 0);
        assert.equal(calls.textDrag, 0);
        assert.equal(calls.extendEscape, 0);
        assert.equal(calls.drawEscape, 0);
      }
    } finally { await page.close(); }
  }
});

test('each legacy editor exposes an explicit mount lifecycle with an idempotent disposer', async () => {
  for (const app of ['quire', 'ledger', 'lectern']) {
    const page = await open(app);
    try {
      const result = await page.evaluate(`(async () => {
        const originalInit = L.app.init;
        const host = document.createElement('section');
        const inside = document.createElement('i'); inside.id = 'mount-inside'; host.append(inside);
        const outside = document.createElement('i'); outside.id = 'mount-outside'; document.body.append(outside);
        let events = 0, busEvents = 0, queriesStayedInHost = false;
        L.app.init = async () => {
          queriesStayedInHost = L.$('#mount-inside') === inside && L.$('#mount-outside') === null;
          L.addMountListener(window, 'vibeoffice-explicit-mount-test', () => events++);
          L.bus.on('vibeoffice-explicit-mount-test', () => busEvents++);
        };
        try {
          const dispose = await L.app.mount(host);
          window.dispatchEvent(new Event('vibeoffice-explicit-mount-test'));
          const activeListenerWorks = events === 1;
          dispose(); dispose();
          window.dispatchEvent(new Event('vibeoffice-explicit-mount-test'));
          L.bus.emit('vibeoffice-explicit-mount-test');
          return { available: typeof dispose === 'function', queriesStayedInHost, activeListenerWorks, listenerRemoved: events === 1, busListenerRemoved: busEvents === 0 };
        } finally { L.app.init = originalInit; L.ui.unmount(host); host.remove(); outside.remove(); }
      })()`);
      assert.deepEqual(result, { available: true, queriesStayedInHost: true, activeListenerWorks: true, listenerRemoved: true, busListenerRemoved: true }, `${app} mount owns its DOM query and listener scope`);
    } finally { await page.close(); }
  }
});

test('file picker is owned by and cancelled with its mounted editor', async () => {
  const page = await open('quire');
  try {
    const result = await page.evaluate(`(async () => {
      const originalInit = L.app.init, originalClick = HTMLInputElement.prototype.click;
      const host = document.createElement('section');
      L.app.init = async () => {};
      HTMLInputElement.prototype.click = function () {};
      try {
        const dispose = await L.app.mount(host);
        const picked = L.pickFiles('.docx');
        const picker = host.querySelector('input[type=file]');
        const wasScoped = !!picker && picker.accept === '.docx' && !document.body.querySelector(':scope > input[type=file]');
        dispose();
        const files = await picked;
        return { wasScoped, removedOnDispose: !host.querySelector('input[type=file]'), cancelled: files.length === 0 };
      } finally {
        HTMLInputElement.prototype.click = originalClick;
        L.app.init = originalInit;
        L.ui.unmount(host);
        host.remove();
      }
    })()`);
    assert.deepEqual(result, { wasScoped: true, removedOnDispose: true, cancelled: true });
  } finally { await page.close(); }
});

test('fallback download is owned by and revoked with its mounted editor', async () => {
  const page = await open('quire');
  try {
    const result = await page.evaluate(`(async () => {
      const originalInit = L.app.init, originalClick = HTMLAnchorElement.prototype.click;
      const originalCreate = URL.createObjectURL, originalRevoke = URL.revokeObjectURL;
      const host = document.createElement('section'), revoked = [];
      L.app.init = async () => {};
      HTMLAnchorElement.prototype.click = function () {};
      URL.createObjectURL = () => 'blob:vibeoffice-host-test';
      URL.revokeObjectURL = (url) => revoked.push(url);
      try {
        const dispose = await L.app.mount(host);
        await L.saveFile('test.docx', new Blob(['test']));
        const anchor = host.querySelector('a[download="test.docx"]');
        const wasScoped = !!anchor && !document.body.querySelector(':scope > a[download="test.docx"]');
        dispose();
        return { wasScoped, removedOnDispose: !host.querySelector('a[download="test.docx"]'), revokedOnDispose: revoked.includes('blob:vibeoffice-host-test') };
      } finally {
        L.app.init = originalInit;
        HTMLAnchorElement.prototype.click = originalClick;
        URL.createObjectURL = originalCreate;
        URL.revokeObjectURL = originalRevoke;
        L.ui.unmount(host);
        host.remove();
      }
    })()`);
    assert.deepEqual(result, { wasScoped: true, removedOnDispose: true, revokedOnDispose: true });
  } finally { await page.close(); }
});
