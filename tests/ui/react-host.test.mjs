import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open, pause } from '../lib/input.mjs';

test('legacy app keyboard handlers ignore events outside the mounted editor host', async () => {
  for (const app of ['quire', 'ledger', 'lectern']) {
    const t = await open(app, { query: { reactHost: '' }, preload: `document.addEventListener('DOMContentLoaded', () => {
      const host = document.createElement('div'); host.dataset.vibeofficeEngineRoot = '';
      while (document.body.firstChild) host.append(document.body.firstChild);
      document.body.append(host);
    });` });
    try {
      const calls = await t.evaluate(`(() => {
        const outside = document.createElement('button');
        document.body.append(outside);
        let calls = 0, clipboardCalls = 0, tableEscapeCalls = 0, textDragCalls = 0, extendEscapeCalls = 0, drawEscapeCalls = 0;
        const original = L.app.menuBar.openIndex;
        L.app.menuBar.openIndex = () => { calls++; };
        outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'F10', bubbles: true, cancelable: true }));
        L.app.menuBar.openIndex = original;
        if ('${app}' === 'quire') {
          const outsideEditor = document.createElement('div');
          outsideEditor.className = 'ed-host';
          const paragraph = document.createElement('div'); paragraph.className = 'p'; outsideEditor.append(paragraph);
          document.body.append(outsideEditor);
          const originalDrag = L.ui.drag;
          L.ui.drag = (...args) => { textDragCalls++; return originalDrag(...args); };
          paragraph.dispatchEvent(new PointerEvent('pointerdown', { button: 0, bubbles: true, cancelable: true }));
          L.ui.drag = originalDrag;
          outsideEditor.remove();
          const originalSetTool = L.tables.setTool;
          const originalTool = L.tables.tool;
          const originalStatus = L.app.updateStatus;
          L.ed.extend = true;
          L.app.updateStatus = () => { extendEscapeCalls++; };
          outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
          L.app.updateStatus = originalStatus;
          L.ed.extend = false;
          const originalStop = L.drawtool.stop;
          L.drawtool.active = {};
          L.drawtool.stop = () => { drawEscapeCalls++; };
          outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
          L.drawtool.stop = originalStop;
          L.drawtool.active = null;
          L.tables.tool = 'draw';
          L.tables.setTool = () => { tableEscapeCalls++; };
          outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
          L.tables.setTool = originalSetTool;
          L.tables.tool = originalTool;
        }
        if ('${app}' === 'lectern') {
          const originalCopy = L.app.copySlides;
          L.app.focusArea = 'slides';
          L.app.copySlides = () => { clipboardCalls++; return null; };
          outside.dispatchEvent(new ClipboardEvent('copy', { bubbles: true, cancelable: true }));
          L.app.copySlides = originalCopy;
        }
        outside.remove();
        return { keyboard: calls, clipboard: clipboardCalls, tableEscape: tableEscapeCalls, textDrag: textDragCalls, extendEscape: extendEscapeCalls, drawEscape: drawEscapeCalls };
      })()`);
      assert.equal(calls.keyboard, 0, `${app} must not handle React chrome keyboard input outside its editor root`);
      if (app === 'lectern') assert.equal(calls.clipboard, 0, 'Lectern must not handle clipboard input outside its editor root');
      if (app === 'quire') assert.equal(calls.tableEscape, 0, 'Quire table tools must not handle Escape outside their editor root');
      if (app === 'quire') assert.equal(calls.textDrag, 0, 'Quire text selection tracking must ignore pointer input outside its editor root');
      if (app === 'quire') assert.equal(calls.extendEscape, 0, 'Quire extend-selection Escape handling must ignore keys outside its editor root');
      if (app === 'quire') assert.equal(calls.drawEscape, 0, 'Quire drawing-tool Escape handling must ignore keys outside its editor root');
    } finally { await t.close(); }
  }
});

test('each legacy editor exposes an explicit mount lifecycle with an idempotent disposer', async () => {
  for (const app of ['quire', 'ledger', 'lectern']) {
    const t = await open(app, { query: { reactHost: '' } });
    try {
      const result = await t.evaluate(`(async () => {
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
    } finally { await t.close(); }
  }
});

test('Quire React-host mode starts the editor without opening the legacy Recent Files database', async () => {
  const t = await open('quire', {
    persistence: true,
    query: { reactHost: '' },
    preload: `window.__legacyDatabaseOpens = 0;
      const open = indexedDB.open.bind(indexedDB);
      indexedDB.open = (...args) => { window.__legacyDatabaseOpens++; return open(...args); };`,
  });
  try {
    assert.equal(await t.evaluate('VO.reactHost'), true);
    assert.equal(await t.evaluate('window.__legacyDatabaseOpens'), 0);
    assert.deepEqual(await t.evaluate('VO.recent.list()'), []);
    assert.deepEqual(await t.evaluate('VO.recentMenuItems()'), []);
    assert.equal(await t.evaluate('!!L.D.doc'), true, 'the Quire document engine initialized');
    const domScope = await t.evaluate(`(() => {
      const outside = document.createElement('i'); outside.id = 'outside-scope'; document.body.append(outside);
      const first = document.createElement('div'); first.innerHTML = '<b id="first-scope"></b>'; document.body.append(first);
      const second = document.createElement('div'); second.innerHTML = '<b id="second-scope"></b>'; document.body.append(second);
      let firstEvents = 0, secondEvents = 0;
      const restoreFirst = L.setDomRoot(first);
      const firstScoped = L.$('#first-scope')?.parentElement === first && L.$('#outside-scope') === null;
      L.addMountListener(window, 'vibeoffice-mount-scope-test', () => firstEvents++);
      const restoreSecond = L.setDomRoot(second);
      const secondScoped = L.$('#second-scope')?.parentElement === second && L.$('#first-scope') === null;
      L.addMountListener(window, 'vibeoffice-mount-scope-test', () => secondEvents++);
      const unmountUi = L.ui.mount(second);
      L.ui.toast('Mounted editor notice');
      L.ui.busy(true, 'Mounted editor operation');
      const sharedLayersScoped = L.domHost() === second && second.querySelector('.toast')?.textContent === 'Mounted editor notice' && !!second.querySelector('.busy');
      L.ui.busy(false);
      const firstUiHost = document.createElement('div'); document.body.append(firstUiHost);
      const unmountFirstUi = L.ui.mount(firstUiHost);
      L.ui.toast('First mounted notice');
      const nestedHostsOwnLayers = firstUiHost.querySelector('.toast')?.textContent === 'First mounted notice' && second.querySelector('.toast')?.textContent === 'Mounted editor notice';
      unmountFirstUi();
      const firstUiLayersRemoved = !firstUiHost.querySelector('.toast');
      firstUiHost.remove();
      second.querySelector('.toast')?.remove();
      restoreFirst();
      const outOfOrderRestore = L.$('#second-scope')?.parentElement === second;
      window.dispatchEvent(new Event('vibeoffice-mount-scope-test'));
      const inactiveScopeListenerRemoved = firstEvents === 0 && secondEvents === 1;
      restoreSecond();
      window.dispatchEvent(new Event('vibeoffice-mount-scope-test'));
      const allScopeListenersRemoved = firstEvents === 0 && secondEvents === 1;
      second.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }));
      const mountedDelegationInstalled = second.querySelectorAll('[data-tip]').length === 0;
      unmountUi();
      const uiUnmountIdempotent = !L.ui.unmount(second);
      const scopedHost = document.createElement('div'); document.body.append(scopedHost);
      const secondScopedHost = document.createElement('div'); document.body.append(secondScopedHost);
      L.ui.unmount(document.body);
      const add = window.addEventListener.bind(window), remove = window.removeEventListener.bind(window);
      const added = [], removed = [];
      window.addEventListener = (type, ...args) => { if (['blur', 'resize', 'mousedown', 'mouseup', 'auxclick', 'contextmenu'].includes(type)) added.push(type); return add(type, ...args); };
      window.removeEventListener = (type, ...args) => { if (['blur', 'resize', 'mousedown', 'mouseup', 'auxclick', 'contextmenu'].includes(type)) removed.push(type); return remove(type, ...args); };
      const mountScope = L.createMountScope(scopedHost);
      const secondMountScope = L.createMountScope(secondScopedHost);
      const windowListenersShared = added.length === 6 && new Set(added).size === 6;
      mountScope.dispose();
      const windowListenersRetained = removed.length === 0;
      secondMountScope.dispose();
      window.addEventListener = add; window.removeEventListener = remove;
      const windowListenersReleased = removed.length === 6 && new Set(removed).size === 6;
      L.ui.mount(document.body);
      const mountScopeAgain = L.createMountScope(scopedHost);
      let scopedEvents = 0, deferredRan = false, cleanupsRan = 0;
      mountScopeAgain.listen(window, 'vibeoffice-editor-scope-test', () => scopedEvents++);
      mountScopeAgain.defer(() => { deferredRan = true; }, 0);
      L.addMountCleanup(() => { cleanupsRan++; });
      const scopeOwnsRootAndUi = L.domHost() === scopedHost && !!scopedHost.querySelector('.toast') === false;
      mountScopeAgain.dispose(); mountScopeAgain.dispose();
      window.dispatchEvent(new Event('vibeoffice-editor-scope-test'));
      const scopeDisposesListeners = scopedEvents === 0 && mountScopeAgain.disposed && !deferredRan && cleanupsRan === 1;
      scopedHost.remove(); secondScopedHost.remove();
      const restoredDocument = L.$('#outside-scope') === outside;
      first.remove(); second.remove(); outside.remove();
      return { firstScoped, secondScoped, sharedLayersScoped, nestedHostsOwnLayers, firstUiLayersRemoved, outOfOrderRestore, inactiveScopeListenerRemoved, allScopeListenersRemoved, mountedDelegationInstalled, uiUnmountIdempotent, windowListenersShared, windowListenersRetained, windowListenersReleased, scopeOwnsRootAndUi, scopeDisposesListeners, restoredDocument };
    })()`);
    assert.deepEqual(domScope, { firstScoped: true, secondScoped: true, sharedLayersScoped: true, nestedHostsOwnLayers: true, firstUiLayersRemoved: true, outOfOrderRestore: true, inactiveScopeListenerRemoved: true, allScopeListenersRemoved: true, mountedDelegationInstalled: true, uiUnmountIdempotent: true, windowListenersShared: true, windowListenersRetained: true, windowListenersReleased: true, scopeOwnsRootAndUi: true, scopeDisposesListeners: true, restoredDocument: true }, 'DOM queries, shared overlay layers, and listeners follow mounted scopes and restore safely');
    await t.evaluate(`L.app.newDocument({ build: (doc) => { doc.main.blocks = [L.D.para([L.D.text('React-host editing')])]; } }); const paragraph = L.D.firstPara(L.D.doc.main.blocks); L.ed.setSel(L.D.pos(paragraph, L.D.plen(paragraph))); L.ed.focus();`);
    await pause(300);
    await t.type(' works');
    await pause(350);
    assert.match(await t.evaluate('L.D.plainText(L.D.doc.main.blocks[0])'), /React-host editing works/);
    await t.key('z', { ctrl: true });
    await pause(200);
    assert.equal(await t.evaluate('L.D.plainText(L.D.doc.main.blocks[0])'), 'React-host editing');
    await t.key('y', { ctrl: true });
    await pause(200);
    assert.match(await t.evaluate('L.D.plainText(L.D.doc.main.blocks[0])'), /React-host editing works/);
    await t.evaluate(`L.app.newDocument({ build: (doc) => { doc.main.blocks = [L.D.para([L.D.text('Recovery placeholder')])]; } })`);
    await pause(200);
    await t.evaluate(`window.postMessage({ type: 'vibeoffice-react-restore', id: 'wordx:test-recovery', name: 'Recovered.txt', blob: new Blob(['Recovered text']), lossState: null }, location.origin)`);
    await new Promise((resolve) => setTimeout(resolve, 500));
    assert.equal(await t.evaluate('L.D.plainText(L.D.doc.main.blocks[0])'), 'Recovered text');
    assert.equal(await t.evaluate('VO.changeCount(L.D.doc) > 0'), true, 'restored work is dirty and autosave is re-armed');
    const openRequested = await t.evaluate(`new Promise((resolve) => {
      const listener = (event) => { if (event.data?.type === 'vibeoffice-react-request-open') { removeEventListener('message', listener); resolve(true); } };
      addEventListener('message', listener); L.app.openDialog(); setTimeout(() => resolve(false), 500);
    })`);
    assert.equal(openRequested, true, 'Quire Open requests use the React host');
    const saveResult = await t.evaluate(`(async () => {
      const response = (event) => { if (event.data?.type === 'vibeoffice-react-request-save') postMessage({ type: 'vibeoffice-react-save-result', requestId: event.data.requestId, result: 'saved' }, location.origin); };
      addEventListener('message', response); const result = await VO.reactFileSave('Recovered.docx', new Blob(['bytes'])); removeEventListener('message', response); return result;
    })()`);
    assert.equal(saveResult, 'saved', 'save requests complete through the host response');
    const appSaveResult = await t.evaluate(`(async () => {
      const response = (event) => { if (event.data?.type === 'vibeoffice-react-request-save') postMessage({ type: 'vibeoffice-react-save-result', requestId: event.data.requestId, result: 'saved' }, location.origin); };
      addEventListener('message', response); const result = await L.app.saveAs(); removeEventListener('message', response); return result;
    })()`);
    assert.equal(appSaveResult, 'saved', 'Quire Save As writes through the React host adapter');
    t.clean();
  } finally { await t.close(); }
});
