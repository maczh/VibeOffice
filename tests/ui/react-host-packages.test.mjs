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

test('an external React host wires its remote provider into direct editor Open and Save', { timeout: 60000 }, async () => {
  const previousBase = process.env.BASE;
  const previousCdp = process.env.CDP;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, [path.join(root, 'node_modules/vite/bin/vite.js'), '--config', 'tests/vite/host.vite.config.ts', '--host', '127.0.0.1', '--port', String(port), '--strictPort'], { cwd: root, stdio: 'ignore' });
  process.env.BASE = base;
  process.env.CDP ||= 'http://127.0.0.1:9223';
  let page, mountPage, errorPage;
  try {
    for (let attempt = 0; attempt < 80 && !await fetch(`${base}/tests/vite/host-fixture.html`).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: `${base}/tests/vite/host-fixture.html?lang=qps`, ready: "document.readyState === 'complete'", timeout: 20000, bypassServiceWorker: true });
    const ready = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+20000; const poll=()=>{const open=[...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Open host-provider sample')); open && !open.disabled ? resolve(true) : Date.now()>end ? resolve(false) : setTimeout(poll,50);}; poll(); })`);
    assert.equal(ready, true, `host fixture readiness: ${JSON.stringify({ body: await page.evaluate('document.body.innerText.slice(0,400)'), errors: page.errors })}`);
    assert.equal(await page.evaluate('document.title'), 'VibeOffice host fixture', 'mounting Wordx does not overwrite the embedding page title');
    const wordxShell = await page.evaluate("(() => ({roots:[...document.querySelectorAll('.vibe-editor__engine > .vibeoffice-legacy-root')].map((root)=>({owner:root.dataset.vibeofficeReactShell,html:root.innerHTML.slice(0,100)})),app:document.querySelector('.vibe-editor__engine > .vibeoffice-legacy-root > #app')?.outerHTML.slice(0,300)}))()");
    assert.equal(wordxShell.roots.length, 1, `Wordx host has one shell root: ${JSON.stringify(wordxShell)}`);
    assert.equal(wordxShell.roots[0].owner, 'wordx', `Wordx shell root is React-owned: ${JSON.stringify(wordxShell)}`);
    assert.match(wordxShell.app || '', /id="app" class="app"/, 'the React shell contains the engine app mount node');
    const runtimeAssets = await page.evaluate("performance.getEntriesByType('resource').map(({name})=>new URL(name).pathname).filter((pathname)=>pathname==='/wordx/runtime.json'||pathname==='/wordx/'||pathname==='/wordx/index.html')");
    assert.deepEqual(runtimeAssets, ['/wordx/runtime.json'], 'React-mounted editors fetch build-generated runtime metadata instead of parsing the legacy HTML shell');
    assert.deepEqual(await page.evaluate('({ id: L.APP_ID, name: L.APP_NAME, app: L.APP })'), { id: 'wordx', name: 'Wordx', app: 'Wordx' }, 'Wordx engine identity is independent of the embedding document metadata');
    assert.deepEqual(await page.evaluate("({ language: VO.i18n.lang, editorLanguage: VO.reactDirectHost.root.lang, hostLanguage: document.documentElement.lang, hostText: document.querySelector('#host-language-sentinel').textContent })"), { language: 'qps', editorLanguage: 'qps-ploc', hostLanguage: 'en', hostText: 'Create File' }, 'pseudo-localization is scoped to the mounted editor and does not rewrite the embedding host DOM');
    const serviceTeardown = await page.evaluate(`new Promise((resolve)=>{const engine=window.L,active={toolbar:engine.ui.toolbar,button:engine.ui.button,combo:engine.ui.combo,openMenu:engine.ui.openMenu},selector=document.querySelector('select[aria-label="VibeOffice editors"]');selector.value='none';selector.dispatchEvent(new Event('change',{bubbles:true}));const end=Date.now()+2500;const poll=()=>{if(!document.querySelector('.vibe-editor'))resolve({toolbar:engine.ui.toolbar!==active.toolbar,button:engine.ui.button!==active.button,combo:engine.ui.combo!==active.combo,openMenu:engine.ui.openMenu!==active.openMenu});else if(Date.now()>end)resolve(false);else setTimeout(poll,10)};poll()})`);
    assert.deepEqual(serviceTeardown, { toolbar: true, button: true, combo: true, openMenu: true }, 'unmount restores the engine namespace shared UI methods before another editor package mounts');
    const remountedWordx = await page.evaluate(`new Promise((resolve)=>{const selector=document.querySelector('select[aria-label="VibeOffice editors"]');selector.value='wordx';selector.dispatchEvent(new Event('change',{bubbles:true}));const end=Date.now()+10000;const poll=()=>document.querySelector('.vibe-editor__engine > .vibeoffice-legacy-root[data-vibeoffice-react-shell="wordx"]')&&![...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Open host-provider sample'))?.disabled?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,20);poll()})`);
    assert.equal(remountedWordx, true, 'the same cached engine namespace reinstalls its React UI services on remount');
    assert.match(await page.evaluate("document.querySelector('[aria-label=\"Editor title\"]')?.textContent || ''"), /Wordx/, 'Wordx reports its document title through the host callback');
    await page.evaluate("[...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Open host-provider sample'))?.click()");
    const opened = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+15000; const poll=()=>{const blocks=window.L?.D?.doc?.main?.blocks; const text=blocks?.length ? L.D.plainText(blocks[0]) : ''; text.includes('Opened through the host-provided storage adapter') ? resolve(text) : Date.now()>end ? resolve(false) : setTimeout(poll,50);}; poll(); })`);
    assert.ok(typeof opened === 'string', 'the editor opened the provider-selected sample');
    assert.equal(await page.evaluate('document.title'), 'VibeOffice host fixture', 'opening through Wordx leaves the host page title unchanged');
    const tabLifecycle = await page.evaluate(`(async()=>{
      const callbacks=[];
      const bodies=[document.createElement('div'),document.createElement('div')];
      bodies[0].textContent='First panel'; bodies[1].textContent='Second panel';
      const tabs=L.ui.tabs([{label:'First',body:bodies[0],onShow:()=>callbacks.push({index:0,hidden:bodies.map((body)=>body.hidden)})},{label:'Second',body:bodies[1],onShow:()=>callbacks.push({index:1,hidden:bodies.map((body)=>body.hidden)})}]);
      VO.reactDirectHost.root.append(tabs); window.__tabLifecycle={tabs,bodies,callbacks};
      const ready=await new Promise((resolve)=>{const end=Date.now()+1500;const poll=()=>tabs.querySelectorAll('[role=tabpanel]').length===2?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,10);poll()});
      if(!ready)return {ready:false}; callbacks.length=0;
      tabs.querySelectorAll('[role=tab]')[1].click();
      const afterClick={selected:tabs.querySelectorAll('[role=tab]')[1].getAttribute('aria-selected'),hidden:bodies.map((body)=>body.hidden),callback:callbacks.at(-1)};
      tabs.select(0);
      const afterApi={selected:tabs.querySelectorAll('[role=tab]')[0].getAttribute('aria-selected'),hidden:bodies.map((body)=>body.hidden),callback:callbacks.at(-1)};
      tabs.remove();
      const released=await new Promise((resolve)=>{const end=Date.now()+1000;const poll=()=>bodies.every((body)=>!body.hasAttribute('role')&&!body.hasAttribute('id')&&!body.hasAttribute('aria-labelledby')&&!body.hasAttribute('hidden')&&!body.classList.contains('tab-pane'))?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,10);poll()});
      const accessible=[...tabs.querySelectorAll('[role=tab]')].every((tab)=>{const panel=document.getElementById(tab.getAttribute('aria-controls'));return panel?.getAttribute('aria-labelledby')===tab.id;});
      return {ready,afterClick,afterApi,released,accessible};
    })()`);
    assert.deepEqual(tabLifecycle, {
      ready: true,
      afterClick: { selected: 'true', hidden: [true, false], callback: { index: 1, hidden: [true, false] } },
      afterApi: { selected: 'true', hidden: [false, true], callback: { index: 0, hidden: [false, true] } },
      released: true,
      accessible: true,
    }, 'shared React tabs own panel visibility before onShow callbacks and restore reusable panel DOM on unmount');
    await page.evaluate("[...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Save through host provider'))?.click()");
    const saved = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+10000; const poll=()=>{const text=document.querySelector('[aria-label="Host provider status"]')?.textContent||''; text==='Saved sample.txt at revision 2' ? resolve(text) : Date.now()>end ? resolve(false) : setTimeout(poll,50);}; poll(); })`);
    assert.equal(saved, 'Saved sample.txt at revision 2', 'the opened revision reaches the host provider on Save');
    await page.evaluate('performance.setResourceTimingBufferSize(5000)');
    for (const [app, title, revision] of [['excelx', 'Excelx editor', 3], ['pptx', 'PPTX editor', 4]]) {
      await page.evaluate(`(() => {const select=document.querySelectorAll('select')[1];select.value=${JSON.stringify(app)};select.dispatchEvent(new Event('change',{bubbles:true}))})()`);
      const mounted = await page.evaluate(`new Promise((resolve) => {const end=Date.now()+15000;const poll=()=>{const ready=document.querySelector('.vibe-editor')?.getAttribute('aria-label')===${JSON.stringify(title)}&&![...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Open host-provider sample'))?.disabled;ready?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,50)};poll()})`);
      assert.equal(mounted, true, `${title} mounts independently in the external host`);
      assert.equal(await page.evaluate("document.querySelector('.vibe-editor__engine > .vibeoffice-legacy-root')?.dataset.vibeofficeReactShell || ''"), app, `${title} engine shell is owned by its React package`);
      const appRuntimeAssets = await page.evaluate(`performance.getEntriesByType('resource').map(({name})=>new URL(name).pathname).filter((pathname)=>pathname==='/${app}/runtime.json'||pathname==='/${app}/'||pathname==='/${app}/index.html')`);
      assert.deepEqual(appRuntimeAssets, [`/${app}/runtime.json`], `${title} loads its generated runtime manifest and not its legacy HTML entry`);
      const expectedName = app === 'excelx' ? 'Excelx' : 'PPTX';
      assert.deepEqual(await page.evaluate('({ id: L.APP_ID, name: L.APP_NAME, app: L.APP })'), { id: app, name: expectedName, app: expectedName }, `${title} refreshes the engine identity when the host switches packages`);
      assert.equal(await page.evaluate('document.title'), 'VibeOffice host fixture', `${title} leaves the external host title unchanged`);
      assert.match(await page.evaluate("document.querySelector('[aria-label=\"Editor title\"]')?.textContent || ''"), new RegExp(app === 'excelx' ? 'Excelx' : 'PPTX'), `${title} reports its title through the optional callback`);
      await page.evaluate("[...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Open host-provider sample'))?.click()");
      const openedAppFile = await page.evaluate(`new Promise((resolve) => {const end=Date.now()+15000;const poll=()=>{const name=window.L?.app?.fileName||'';name.includes('sample')?resolve(name):Date.now()>end?resolve(false):setTimeout(poll,50)};poll()})`);
      assert.ok(typeof openedAppFile === 'string', `${title} opens the host provider's selected document`);
      await page.evaluate("[...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Save through host provider'))?.click()");
      const savedAppFile = await page.evaluate(`new Promise((resolve) => {const expected=${JSON.stringify(`Saved sample.txt at revision ${revision}`)},end=Date.now()+10000;const poll=()=>{const text=document.querySelector('[aria-label="Host provider status"]')?.textContent||'';text===expected?resolve(text):Date.now()>end?resolve(false):setTimeout(poll,50)};poll()})`);
      assert.equal(savedAppFile, `Saved sample.txt at revision ${revision}`, `${title} preserves the provider revision through Save`);
    }
    assert.deepEqual(page.errors, []);

    errorPage = await openPage(undefined, { url: `${base}/tests/vite/host-fixture.html?error-boundary=1`, ready: "document.readyState === 'complete'", timeout: 20000, bypassServiceWorker: true });
    const renderFailure = await errorPage.evaluate(`new Promise((resolve) => { const end=Date.now()+2500; const poll=()=>{const fallback=document.querySelector('[data-host-render-fallback]');const error=document.querySelector('[aria-label="Host provider status"]')?.textContent;const editor=document.querySelector('.vibe-editor');fallback&&error==='Host-provided React layer failed'&&editor?resolve({fallback:fallback.textContent,role:fallback.getAttribute('role'),editor:editor.getAttribute('aria-label'),error}):Date.now()>end?resolve(false):setTimeout(poll,50)};poll(); })`);
    assert.deepEqual(renderFailure, { fallback: 'Editor layer unavailable', role: 'alert', editor: 'Wordx editor', error: 'Host-provided React layer failed' }, 'a host-provided React layer failure is isolated, reported with its original error, and rendered accessibly without removing the editor');

    mountPage = await openPage(undefined, { url: `${base}/tests/vite/engine-mount-fixture.html`, ready: "typeof window.mountEngineFixture === 'function'", timeout: 10000, bypassServiceWorker: true });
    const mountOwnership = await mountPage.evaluate("(async()=>{const host=document.querySelector('#engine-host'),dispose=await window.mountEngineFixture(host);const mounted={before:!!host.querySelector('#host-before'),shell:host.textContent.includes('Mounted engine shell'),root:!!host.querySelector('[data-vibeoffice-engine-root]')};const after=document.createElement('span');after.id='host-after';host.append(after);dispose();return {...mounted,after:host.contains(after),before:!!host.querySelector('#host-before'),rootRemoved:!host.querySelector('[data-vibeoffice-engine-root]'),hostConnected:host.isConnected}})()");
    assert.deepEqual(mountOwnership, { before: true, shell: true, root: true, after: true, rootRemoved: true, hostConnected: true }, 'mount and disposal preserve host-owned content and remove only the editor root');
    const cancelledScriptHook = await mountPage.evaluate("(async()=>{const host=document.querySelector('#engine-host'),controller=new AbortController(),calls=[],previous=Object.getOwnPropertyDescriptor(window,'L'),hostApp={name:'embedding app'};Object.defineProperty(window,'L',{configurable:true,enumerable:true,writable:true,value:hostApp});try{await window.mountEngineFixture(host,{signal:controller.signal,beforeLegacyScript:(_runtime,script)=>{calls.push(script.pathname);if(calls.length===1)controller.abort()}});return {resolved:true,calls,root:!!host.querySelector('[data-vibeoffice-engine-root]')}}catch(error){return {name:error.name,calls,root:!!host.querySelector('[data-vibeoffice-engine-root]'),hostRestored:window.L===hostApp,hostUntouched:hostApp.name==='embedding app'}}finally{if(previous)Object.defineProperty(window,'L',previous);else delete window.L}})()");
    assert.deepEqual(cancelledScriptHook, { name: 'AbortError', calls: ['/fixture/pre-xml.js'], root: false, hostRestored: true, hostUntouched: true }, 'an aborted engine mount stops before subsequent compatibility hooks, removes its temporary root, and restores the embedding window.L namespace');
    const abortedBridgeMount = await mountPage.evaluate("(async()=>{const host=document.querySelector('#engine-host'),controller=new AbortController(),previous=Object.getOwnPropertyDescriptor(window,'L'),hostApp={name:'embedding app'};Object.defineProperty(window,'L',{configurable:true,enumerable:true,writable:true,value:hostApp});let startedResolve,stopped=0,started=false;const startedPromise=new Promise(resolve=>startedResolve=resolve);try{const mounting=window.mountBridgeFixture(host,controller.signal,()=>{started=true;startedResolve()},()=>stopped++);const start=await Promise.race([startedPromise.then(()=>true),new Promise(resolve=>setTimeout(()=>resolve(false),5000))]);if(!start){const settled=await Promise.race([mounting.then(()=>({ok:true}),error=>({error:String(error)})),new Promise(resolve=>setTimeout(()=>resolve(null),100))]);return {started,settled,root:!!host.querySelector('[data-vibeoffice-engine-root]'),errors:window.__vibeofficeDraftRecoveries};}const mountedRoot=!!host.querySelector('[data-vibeoffice-engine-root]');controller.abort();const removedOnAbort=!host.querySelector('[data-vibeoffice-engine-root]'),hostRestoredOnAbort=window.L===hostApp;window.finishBridgeFixtureMount();const cleanup=await mounting;cleanup();return {mountedRoot,removedOnAbort,hostRestoredOnAbort,stopped,hostConnected:host.isConnected};}finally{delete window.finishBridgeFixtureMount;if(previous)Object.defineProperty(window,'L',previous);else delete window.L}})()");
    assert.deepEqual(abortedBridgeMount, { mountedRoot: true, removedOnAbort: true, hostRestoredOnAbort: true, stopped: 1, hostConnected: true }, 'aborting during an asynchronous app mount immediately releases the engine document and later invokes its returned disposer exactly once');
    const abortedBridgePreparation = await mountPage.evaluate("(async()=>{const host=document.querySelector('#engine-host'),controller=new AbortController(),previous=Object.getOwnPropertyDescriptor(window,'L'),hostApp={name:'embedding app'};Object.defineProperty(window,'L',{configurable:true,enumerable:true,writable:true,value:hostApp});let started=false,stopped=0;try{await window.mountBridgeFixture(host,controller.signal,()=>{started=true;window.finishBridgeFixtureMount?.()},()=>stopped++,()=>controller.abort());return {name:'resolved',started,root:!!host.querySelector('[data-vibeoffice-engine-root]'),hostRestored:window.L===hostApp,stopped};}catch(error){return {name:error.name,started,root:!!host.querySelector('[data-vibeoffice-engine-root]'),hostRestored:window.L===hostApp,stopped};}finally{delete window.finishBridgeFixtureMount;if(previous)Object.defineProperty(window,'L',previous);else delete window.L}})()");
    assert.deepEqual(abortedBridgePreparation, { name: 'AbortError', started: false, root: false, hostRestored: true, stopped: 0 }, 'cancellation during runtime configuration is observed before any async app mount starts');
    const failedBridgeConfiguration = await mountPage.evaluate("(async()=>{const host=document.querySelector('#engine-host'),previous=Object.getOwnPropertyDescriptor(window,'L'),hostApp={name:'embedding app'};Object.defineProperty(window,'L',{configurable:true,enumerable:true,writable:true,value:hostApp});let started=false,stopped=0;try{await window.mountBridgeFixture(host,new AbortController().signal,()=>{started=true},()=>stopped++,()=>{throw new Error('host runtime configuration failed')});return {message:'resolved',started,root:!!host.querySelector('[data-vibeoffice-engine-root]'),hostRestored:window.L===hostApp,stopped};}catch(error){return {message:error.message,started,root:!!host.querySelector('[data-vibeoffice-engine-root]'),hostRestored:window.L===hostApp,stopped};}finally{if(previous)Object.defineProperty(window,'L',previous);else delete window.L}})()");
    assert.deepEqual(failedBridgeConfiguration, { message: 'host runtime configuration failed', started: false, root: false, hostRestored: true, stopped: 0 }, 'a rejected host runtime configuration releases the mounted engine document before returning its error');
    const hostNamespace = await mountPage.evaluate("(async()=>{const host=document.querySelector('#engine-host'),previous=Object.getOwnPropertyDescriptor(window,'L'),hostApp={name:'embedding app'};Object.defineProperty(window,'L',{configurable:true,enumerable:true,writable:true,value:hostApp});let stop,cached=null,checks=0;const hook=(runtime)=>{if(!cached){runtime.L={engine:true};cached=runtime.L}else if(runtime.L!==cached)throw new Error('the cached engine namespace was not restored');checks++};try{stop=await window.mountEngineFixture(host,{beforeLegacyScript:hook});const firstOwned=window.L!==hostApp&&window.L===cached;stop();stop=null;const firstRestored=window.L===hostApp;stop=await window.mountEngineFixture(host,{beforeLegacyScript:hook});const remountOwned=window.L!==hostApp&&window.L===cached;stop();stop=null;return {firstOwned,firstRestored,remountOwned,restored:window.L===hostApp,hostUntouched:hostApp.name==='embedding app',hookChecks:checks};}finally{stop?.();if(previous)Object.defineProperty(window,'L',previous);else delete window.L}})()");
    assert.deepEqual(hostNamespace, { firstOwned: true, firstRestored: true, remountOwned: true, restored: true, hostUntouched: true, hookChecks: 4 }, 'direct engine remounts reuse their isolated legacy namespace while restoring the embedding window.L after every disposal');
    const hostAccessor = await mountPage.evaluate("(async()=>{const host=document.querySelector('#engine-host'),previous=Object.getOwnPropertyDescriptor(window,'L'),hostApp={name:'getter-backed embedding app'},getHostNamespace=()=>hostApp;Object.defineProperty(window,'L',{configurable:true,enumerable:true,get:getHostNamespace});let stop;try{stop=await window.mountEngineFixture(host,{beforeLegacyScript(runtime){runtime.L={engine:true}}});const engineOwned=window.L!==hostApp&&window.L.engine===true;stop();stop=null;const restored=Object.getOwnPropertyDescriptor(window,'L');return {engineOwned,getterRestored:restored.get===getHostNamespace,hostRestored:window.L===hostApp,descriptorRestored:restored.configurable&&restored.enumerable&&!('value'in restored)};}finally{stop?.();if(previous)Object.defineProperty(window,'L',previous);else delete window.L}})()");
    assert.deepEqual(hostAccessor, { engineOwned: true, getterRestored: true, hostRestored: true, descriptorRestored: true }, 'mount and disposal preserve a getter-backed embedding window.L property descriptor');
    const concurrentMounts = await mountPage.evaluate("(async()=>{const host=document.querySelector('#engine-host');const results=await Promise.allSettled([window.mountEngineFixture(host),window.mountEngineFixture(host)]);const mounted=results.filter(result=>result.status==='fulfilled');const rejected=results.filter(result=>result.status==='rejected');mounted.forEach(result=>result.value());return {fulfilled:mounted.length,rejected:rejected.length,error:rejected[0]?.reason?.message,roots:host.querySelectorAll('[data-vibeoffice-engine-root]').length}})()");
    assert.deepEqual(concurrentMounts, { fulfilled: 1, rejected: 1, error: 'Only one VibeOffice editor can be mounted in a host document at a time', roots: 0 }, 'concurrent editor mounts reserve the host window before asynchronous stylesheet loading');
    assert.deepEqual(mountPage.errors, []);
  } finally {
    if (mountPage) await mountPage.close();
    if (errorPage) await errorPage.close();
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
    if (previousCdp === undefined) delete process.env.CDP; else process.env.CDP = previousCdp;
  }
});
