import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { openPage } from '../../tools/ooxml/cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const trackLegacyBusSubscriptions = `(() => {
  const counts = Object.create(null), tracked = new WeakMap();
  const wrapBus = (bus) => new Proxy(bus, { get(target, method, receiver) {
      if (method === 'on') {
        const on = Reflect.get(target, method, receiver);
        return function (event, callback) {
          counts[event] = (counts[event] || 0) + 1;
          tracked.set(callback, event);
          return on.call(this, event, callback);
        };
      }
      if (method === 'off') {
        const off = Reflect.get(target, method, receiver);
        return function (event, callback) {
          if (tracked.get(callback) === event) { counts[event]--; tracked.delete(callback); }
          return off.call(this, event, callback);
        };
      }
      return Reflect.get(target, method, receiver);
  } });
  let current;
  Object.defineProperty(window, 'L', { configurable: true, get: () => current, set: (value) => {
    current = value;
    if (!value || Object.prototype.hasOwnProperty.call(value, '__busProbe')) return;
    Object.defineProperty(value, '__busProbe', { value: true });
    let bus;
    Object.defineProperty(value, 'bus', { configurable: true, enumerable: true, get: () => bus, set: (next) => { bus = wrapBus(next); } });
  } });
  window.__legacyBusSubscriptionCount = (event) => counts[event] || 0;
})();`;

test('Excelx React task-pane search inputs have accessible names', { timeout: 30000 }, async () => {
  const probe = net.createServer();
  await new Promise((resolve) => probe.listen(0, '127.0.0.1', resolve));
  const port = probe.address().port;
  await new Promise((resolve) => probe.close(resolve));
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'suite/dist')], { stdio: 'ignore' });
  const previousBase = process.env.BASE;
  process.env.BASE = `http://127.0.0.1:${port}`;
  let page;
  try {
    for (let attempt = 0; attempt < 40; attempt++) {
      if (await fetch(process.env.BASE).then((response) => response.ok).catch(() => false)) break;
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    page = await openPage(undefined, { timeout: 15000, bypassServiceWorker: true, preload: trackLegacyBusSubscriptions });
    const started = await page.evaluate("new Promise((resolve) => { const end = Date.now() + 5000; const poll = () => document.querySelector('.suite-start') ? resolve(true) : Date.now() > end ? resolve(false) : setTimeout(poll, 50); poll(); })");
    assert.equal(started, true, 'the React Start Center rendered');
    await page.evaluate("(() => { window.__preinitBusCalls = 0; document.addEventListener('vibeoffice-engine-runtime-ready', (event) => { if (event.detail?.editor !== 'excelx') return; window.L.bus.on('__preinit-mount-scope-probe', () => window.__preinitBusCalls++); }); })()");
    await page.evaluate("document.querySelectorAll('.app')[1]?.querySelector('button')?.click()");
    const editorReady = await page.evaluate("new Promise((resolve) => { const end = Date.now() + 10000; const poll = () => window.L?.app?.books?.[window.L.app.cur]?.wb ? resolve(true) : Date.now() > end ? resolve(false) : setTimeout(poll, 50); poll(); })");
    assert.equal(editorReady, true, `Excelx engine mounted: ${page.errors.join('\n')}`);
    const preinitSubscription = await page.evaluate("(() => { L.bus.emit('__preinit-mount-scope-probe'); return window.__preinitBusCalls; })()");
    assert.equal(preinitSubscription, 1, 'a bus subscription installed after the runtime-ready event survives A.mount adopting the runtime DOM root');
    const directHostAppReadyListeners = await page.evaluate("({ directHost: !!VO.reactDirectHost, gridLoaded: window.__VIBE_OFFICE_DIRECT_HOST__?.loadedLegacyScripts?.some((path) => path.endsWith('/js/grid.js')), listeners: window.__legacyBusSubscriptionCount('app-ready') })");
    assert.deepEqual(directHostAppReadyListeners, { directHost: true, gridLoaded: true, listeners: 0 }, 'Excelx direct-host does not retain a script-evaluation app-ready subscription for a classic-only pointer listener');
    const helpCatalog = await page.evaluate("(() => { const catalog = L.reactModules?.excelxHelpContent; return { packageOwned: !!catalog, topics: catalog?.topics.length, keys: catalog?.keys.length, legacyAdapter: typeof L.panes?.reactHelpData }; })()");
    assert.deepEqual(helpCatalog, { packageOwned: true, topics: 15, keys: 54, legacyAdapter: 'undefined' }, 'the Excelx package owns the Help and keyboard-shortcut catalog without a legacy React bridge');
    await page.evaluate("L.app.toggleTask(true); L.panes.task.show('getting-started')");
    const ready = await page.evaluate("new Promise((resolve) => { const end = Date.now() + 10000; const poll = () => { const input = document.querySelector('#react-excelx-task-body-root input'); input ? resolve({ label: input.getAttribute('aria-label'), placeholder: input.placeholder }) : Date.now() > end ? resolve({ body: document.body.innerText.slice(-500), errors: window.__errors || [], apps: [...document.querySelectorAll('.app')].map((app) => app.textContent.trim()) }) : setTimeout(poll, 50); }; poll(); })");
    assert.deepEqual(ready, { label: 'Search help', placeholder: 'Search help' }, `Getting Started search has a translated accessible name: ${JSON.stringify(ready)}`);
    await page.evaluate("[...document.querySelectorAll('#react-excelx-task-body-root .tp-link')].find((button) => button.textContent.includes('What Excelx can do'))?.click()");
    const help = await page.evaluate("new Promise((resolve) => { const end = Date.now() + 3000; const poll = () => { const input = document.querySelector('#react-excelx-task-body-root #hp-q'); const topics = document.querySelectorAll('#react-excelx-task-body-root .help-topic').length; input && topics === 15 ? resolve({ label: input.getAttribute('aria-label'), placeholder: input.placeholder, topics }) : Date.now() > end ? resolve(false) : setTimeout(poll, 25); }; poll(); })");
    assert.deepEqual(help, { label: 'Search', placeholder: 'Search', topics: 15 }, 'React renders the package-owned Help catalog and an accessible search field');
    await page.evaluate("L.panes.task.show('keys')");
    const shortcuts = await page.evaluate("new Promise((resolve) => { const end = Date.now() + 3000; const poll = () => { const rows = document.querySelectorAll('#react-excelx-task-body-root tr').length; rows === 54 ? resolve(rows) : Date.now() > end ? resolve(false) : setTimeout(poll, 25); }; poll(); })");
    assert.equal(shortcuts, 54, 'React renders all package-owned keyboard shortcuts');
    const templateCatalog = await page.evaluate("new Promise((resolve)=>{L.panes.task.show('new');const end=Date.now()+2000;const poll=()=>{const section=[...document.querySelectorAll('#react-excelx-task-body-root .tp-sec')].find((entry)=>entry.querySelector('.tp-h')?.textContent==='Templates');if(section)resolve({expected:Object.entries(L.panes.TEMPLATES).map(([id,item])=>({id,name:item.name})),rendered:[...section.querySelectorAll('[data-template-id]')].map((button)=>({id:button.getAttribute('data-template-id'),name:button.textContent.trim()}))});else if(Date.now()>end)resolve(false);else setTimeout(poll,20)};poll()})");
    assert.deepEqual(templateCatalog.rendered, templateCatalog.expected, 'React renders the package-owned template catalog with every engine ID and display name in order');
    const templateCreations = await page.evaluate("(()=>{const previous=VO.oneDocument;VO.oneDocument=false;const cases=[['loan','LoanAmortization1','Loan','Loan Amortization Schedule'],['expense','ExpenseStatement1','Expenses','Expense Statement'],['invoice','SalesInvoice1','Invoice','INVOICE'],['timecard','Timecard1','Timecard','Weekly Timecard'],['budget','PersonalMonthlyBudget1','Budget','Personal Monthly Budget'],['tasks','TaskTracker1','Tasks','Task Tracker'],['shopping','ShoppingList1','Shopping','Shopping List'],['schedule','WeeklySchedule1','Schedule','Weekly Schedule']];try{return cases.map(([id,name,sheet,title])=>{[...document.querySelectorAll('#react-excelx-task-body-root [data-template-id]')].find(button=>button.dataset.templateId===id)?.click();const book=L.app.books[L.app.cur];return {id,name:book?.name,sheet:book?.wb?.sheets?.[0]?.name,title:book?.wb?.sheets?.[0]?.cell(0,0)?.v,expected:{name,sheet,title}}})}finally{VO.oneDocument=previous}})()");
    assert.deepEqual(templateCreations.map(({ id, name, sheet, title }) => ({ id, name, sheet, title })), templateCreations.map(({ id, expected }) => ({ id, ...expected })), 'each React-owned template choice creates the expected workbook and first-sheet content through the engine service');
    await page.evaluate("new Promise((resolve)=>{L.panes.task.show('getting-started');const end=Date.now()+2000;const poll=()=>{const input=document.querySelector('#react-excelx-task-body-root input[aria-label=\"Search help\"]');if(input){input.focus();resolve(true)}else if(Date.now()>end)resolve(false);else setTimeout(poll,20)};poll()})");
    await page.send('Input.insertText', { text: 'formulas' });
    await page.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13 });
    await page.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13 });
    const keyboardSearch = await page.evaluate("new Promise((resolve) => { const end=Date.now()+3000; const poll=()=>{const body=document.querySelector('#react-excelx-task-body-root'); const title=body?.querySelector('.help-topic h4')?.textContent; title==='Formulas'?resolve({title,query:body.querySelector('#hp-q')?.value,current:L.panes.task.current}):Date.now()>end?resolve({title,query:body?.querySelector('input')?.value,current:L.panes.task.current,helpQuery:L.panes.helpQuery,text:body?.innerText.slice(0,400)}):setTimeout(poll,25)};poll()})");
    assert.equal(keyboardSearch.title, 'Formulas', `Enter submits the React-controlled Getting Started search and filters the package-owned catalog: ${JSON.stringify(keyboardSearch)}`);
    await page.evaluate("new Promise((resolve)=>{L.panes.task.show('getting-started');const end=Date.now()+2000;const poll=()=>{const input=document.querySelector('#react-excelx-task-body-root input[aria-label=\"Search help\"]');if(input){input.focus();resolve(true)}else if(Date.now()>end)resolve(false);else setTimeout(poll,20)};poll()})");
    await page.send('Input.insertText', { text: 'AutoFill' });
    const typedSearch = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+1000;const poll=()=>{const input=document.querySelector('#react-excelx-task-body-root input[aria-label=\"Search help\"]');input?.value==='AutoFill'?resolve(true):Date.now()>end?resolve(input?.value||false):setTimeout(poll,20)};poll()})");
    assert.equal(typedSearch, true, `the Getting Started input receives the query before Go: ${typedSearch}`);
    await page.evaluate("[...document.querySelectorAll('#react-excelx-task-body-root button')].find((button) => button.textContent.trim() === 'Go')?.click()");
    const buttonSearch = await page.evaluate("new Promise((resolve) => { const end=Date.now()+3000; const poll=()=>{const body=document.querySelector('#react-excelx-task-body-root');const title=body?.querySelector('.help-topic h4')?.textContent;title==='AutoFill and series'?resolve({title,query:body.querySelector('#hp-q')?.value,current:L.panes.task.current}):Date.now()>end?resolve({title,query:body?.querySelector('input')?.value,current:L.panes.task.current,helpQuery:L.panes.helpQuery,text:body?.innerText.slice(0,400)}):setTimeout(poll,25)};poll()})");
    assert.equal(buttonSearch.title, 'AutoFill and series', `Go submits the same React-controlled query without reading the DOM: ${JSON.stringify(buttonSearch)}`);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE;
    else process.env.BASE = previousBase;
  }
});
