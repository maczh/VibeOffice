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
  let page;
  try {
    for (let attempt = 0; attempt < 80 && !await fetch(`${base}/tests/vite/host-fixture.html`).then((response) => response.ok).catch(() => false); attempt++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: `${base}/tests/vite/host-fixture.html?lang=qps`, ready: "document.readyState === 'complete'", timeout: 20000, bypassServiceWorker: true });
    const ready = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+20000; const poll=()=>{const open=[...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Open host-provider sample')); open && !open.disabled ? resolve(true) : Date.now()>end ? resolve(false) : setTimeout(poll,50);}; poll(); })`);
    assert.equal(ready, true, `host fixture readiness: ${JSON.stringify({ body: await page.evaluate('document.body.innerText.slice(0,400)'), errors: page.errors })}`);
    assert.equal(await page.evaluate('document.title'), 'VibeOffice host fixture', 'mounting Wordx does not overwrite the embedding page title');
    assert.deepEqual(await page.evaluate('({ id: L.APP_ID, name: L.APP_NAME, app: L.APP })'), { id: 'wordx', name: 'Wordx', app: 'Wordx' }, 'Wordx engine identity is independent of the embedding document metadata');
    assert.deepEqual(await page.evaluate("({ language: VO.i18n.lang, editorLanguage: VO.reactDirectHost.root.lang, hostLanguage: document.documentElement.lang, hostText: document.querySelector('#host-language-sentinel').textContent })"), { language: 'qps', editorLanguage: 'qps-ploc', hostLanguage: 'en', hostText: 'Create File' }, 'pseudo-localization is scoped to the mounted editor and does not rewrite the embedding host DOM');
    assert.match(await page.evaluate("document.querySelector('[aria-label=\"Editor title\"]')?.textContent || ''"), /Wordx/, 'Wordx reports its document title through the host callback');
    await page.evaluate("[...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Open host-provider sample'))?.click()");
    const opened = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+15000; const poll=()=>{const blocks=window.L?.D?.doc?.main?.blocks; const text=blocks?.length ? L.D.plainText(blocks[0]) : ''; text.includes('Opened through the host-provided storage adapter') ? resolve(text) : Date.now()>end ? resolve(false) : setTimeout(poll,50);}; poll(); })`);
    assert.ok(typeof opened === 'string', 'the editor opened the provider-selected sample');
    assert.equal(await page.evaluate('document.title'), 'VibeOffice host fixture', 'opening through Wordx leaves the host page title unchanged');
    await page.evaluate("[...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Save through host provider'))?.click()");
    const saved = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+10000; const poll=()=>{const text=document.querySelector('[aria-label="Host provider status"]')?.textContent||''; text==='Saved sample.txt at revision 2' ? resolve(text) : Date.now()>end ? resolve(false) : setTimeout(poll,50);}; poll(); })`);
    assert.equal(saved, 'Saved sample.txt at revision 2', 'the opened revision reaches the host provider on Save');
    for (const [app, title, revision] of [['excelx', 'Excelx editor', 3], ['pptx', 'PPTX editor', 4]]) {
      await page.evaluate(`(() => {const select=document.querySelectorAll('select')[1];select.value=${JSON.stringify(app)};select.dispatchEvent(new Event('change',{bubbles:true}))})()`);
      const mounted = await page.evaluate(`new Promise((resolve) => {const end=Date.now()+15000;const poll=()=>{const ready=document.querySelector('.vibe-editor')?.getAttribute('aria-label')===${JSON.stringify(title)}&&![...document.querySelectorAll('button')].find((button)=>button.textContent.includes('Open host-provider sample'))?.disabled;ready?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,50)};poll()})`);
      assert.equal(mounted, true, `${title} mounts independently in the external host`);
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
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
    if (previousCdp === undefined) delete process.env.CDP; else process.env.CDP = previousCdp;
  }
});
