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

test('React-owned Wordx SVG artwork drops active and external markup while retaining local artwork', { timeout: 60000 }, async () => {
  const previousBase = process.env.BASE;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'wordx/standalone')], { stdio: 'ignore' });
  process.env.BASE = base;
  let page;
  try {
    for (let n = 0; n < 40 && !await fetch(base).then((response) => response.ok).catch(() => false); n++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { timeout: 15000, bypassServiceWorker: true });
    const mounted = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+10000;const poll=()=>window.L?.app?.menus?.borders?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,50);poll()})");
    const state = mounted ? null : await page.evaluate("({url:location.href,title:document.title,editor:!!document.querySelector('.vibe-editor__engine #app'),app:!!L?.app,menus:Object.keys(L?.app?.menus||{}),catalog:typeof L?.app?.reactBorderMenuCatalog,loaded:window.__VIBE_OFFICE_DIRECT_HOST__?.loadedLegacyScripts||[]})");
    assert.equal(mounted, true, `Wordx direct host is ready: ${JSON.stringify({ state, errors: page.errors, requests: page.requests })}`);
    await page.evaluate("L.panes.task.show('getting-started')");
    const taskArtwork = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+3000;const poll=()=>{const svg=document.querySelector('svg[data-react-owned=wordx-task-content]');if(svg)resolve({owner:svg.dataset.reactOwned,scripts:svg.querySelectorAll('script').length,handlers:[...svg.querySelectorAll('*')].filter((node)=>[...node.attributes].some((attribute)=>/^on/i.test(attribute.name))).length});else if(Date.now()>end)resolve(false);else setTimeout(poll,20)};poll()})");
    assert.deepEqual(taskArtwork, { owner: 'wordx-task-content', scripts: 0, handlers: 0 }, 'Wordx React task-pane icons are reconciled as sanitized React SVG nodes rather than injected markup');

    await page.evaluate(`(() => {
      window.__svgExecuted = 0;
      const originalIcon = L.icons.get.bind(L.icons);
      const hostileSvg = '<svg width="18" height="18" viewBox="0 0 18 18" onload="window.__svgExecuted=1"><defs><linearGradient id="safe-gradient"><stop offset="0" stop-color="red"/></linearGradient></defs><script>window.__svgExecuted=2</script><foreignObject><img src="https://example.invalid/image.png"/></foreignObject><a href="javascript:window.__svgExecuted=4"><text>link</text></a><animate attributeName="x" onbegin="window.__svgExecuted=5"/><animateColor attributeName="fill" from="red" to="blue"/><discard begin="1s" href="#target"/><path d="M0 0h18v18" fill="url(#safe-gradient)" onmouseover="window.__svgExecuted=3"/><path d="M1 1h16" fill="url(https://example.invalid/palette.svg#remote)"/><use href="https://example.invalid/art.svg#remote"/></svg>';
      L.icons.get = (name, ...args) => name === 'hostileFixture' ? hostileSvg : originalIcon(name, ...args);
      L.app.BORDER_ITEMS.splice(0, L.app.BORDER_ITEMS.length, ['fixture', 'hostileFixture', 'SVG security fixture']);
      L.app.menus.borders({ left: 20, top: 20 });
    })()`);
    const rendered = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+3000;const poll=()=>{const svg=document.querySelector('#react-wordx-border-menu svg[data-react-owned=wordx-border]');if(svg){const paths=svg.querySelectorAll('path');resolve({owner:svg.dataset.reactOwned,scripts:svg.querySelectorAll('script').length,foreignObjects:svg.querySelectorAll('foreignObject').length,links:svg.querySelectorAll('a').length,animations:svg.querySelectorAll('animate,animateColor,discard').length,onload:svg.hasAttribute('onload'),onmouseover:paths[0]?.hasAttribute('onmouseover'),localFill:paths[0]?.getAttribute('fill'),remoteFill:paths[1]?.getAttribute('fill'),externalHref:svg.querySelector('use')?.getAttribute('href'),executed:window.__svgExecuted});}else if(Date.now()>end)resolve(false);else setTimeout(poll,20)};poll()})");
    assert.deepEqual(rendered, { owner: 'wordx-border', scripts: 0, foreignObjects: 0, links: 0, animations: 0, onload: false, onmouseover: false, localFill: 'url(#safe-gradient)', remoteFill: null, externalHref: null, executed: 0 });
    assert.deepEqual(page.errors, []);
    assert.deepEqual(page.requests, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});

test('Excelx React AutoShapes renders catalog previews as owned SVG nodes', { timeout: 60000 }, async () => {
  const previousBase = process.env.BASE;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'excelx/standalone')], { stdio: 'ignore' });
  process.env.BASE = base;
  let page;
  try {
    for (let n = 0; n < 40 && !await fetch(base).then((response) => response.ok).catch(() => false); n++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: base + '/', timeout: 15000, bypassServiceWorker: true });
    const ready = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+10000;const poll=()=>window.L?.drawing?.autoShapesMenu?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,50);poll()})");
    const state = ready ? null : await page.evaluate("({url:location.href,title:document.title,body:document.body.innerText.slice(0,400),app:window.L?.APP_ID,drawing:typeof window.L?.drawing?.autoShapesMenu,editor:!!document.querySelector('.vibe-editor__engine #app')})");
    assert.equal(ready, true, `Excelx direct host is ready: ${JSON.stringify({ state, errors: page.errors, requests: page.requests })}`);
    await page.evaluate("L.drawing.autoShapesMenu({left:20,top:20})");
    await page.evaluate("document.querySelectorAll('.menu .m-item.has-sub')[1]?.click()");
    const rendered = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+3000;const poll=()=>{const root=document.querySelector('#react-excelx-shape-menu-root'),svg=root?.querySelector('svg[data-react-owned=excelx-shape-menu]');if(svg)resolve({count:root.querySelectorAll('button').length,owner:svg.dataset.reactOwned,paths:svg.querySelectorAll('path').length,handlers:[...svg.querySelectorAll('*')].filter((node)=>[...node.attributes].some((attribute)=>/^on/i.test(attribute.name))).length});else if(Date.now()>end)resolve(false);else setTimeout(poll,20)};poll()})");
    const menuState = rendered ? null : await page.evaluate("({text:document.body.innerText.slice(-700),menus:[...document.querySelectorAll('.menu,[role=menu]')].map((node)=>({text:node.innerText,html:node.outerHTML.slice(0,1200)})),roots:[...document.querySelectorAll('[id*=shape-menu]')].map((node)=>node.outerHTML.slice(0,300))})");
    assert.ok(rendered && rendered.count > 0 && rendered.owner === 'excelx-shape-menu' && rendered.paths > 0 && rendered.handlers === 0, `Excelx AutoShapes are React-owned vector nodes: ${JSON.stringify({ rendered, menuState })}`);
    assert.deepEqual(page.errors, []);
    assert.deepEqual(page.requests, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});

test('PPTX React SVG menus sanitize generated previews in the production editor', { timeout: 60000 }, async () => {
  const previousBase = process.env.BASE;
  const port = await unusedPort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'pptx/standalone')], { stdio: 'ignore' });
  process.env.BASE = base;
  let page;
  try {
    for (let n = 0; n < 40 && !await fetch(base).then((response) => response.ok).catch(() => false); n++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, { url: base + '/', timeout: 15000, bypassServiceWorker: true });
    const ready = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+10000;const poll=()=>window.L?.app?.reactShapeMenuCatalog?resolve(true):Date.now()>end?resolve(false):setTimeout(poll,50);poll()})");
    const state = ready ? null : await page.evaluate("({url:location.href,title:document.title,editor:!!document.querySelector('.vibe-editor__engine #app'),app:window.L?.APP_ID,loaded:window.__VIBE_OFFICE_DIRECT_HOST__?.loadedLegacyScripts||[]})");
    assert.equal(ready, true, `PPTX standalone is ready: ${JSON.stringify({ state, errors: page.errors, requests: page.requests })}`);
    await page.evaluate(`(() => {
      window.__pptxSvgExecuted = 0;
      const hostile = '<svg width="18" height="18" viewBox="0 0 18 18" onload="window.__pptxSvgExecuted=1"><script>window.__pptxSvgExecuted=2</script><foreignObject><img src="https://example.invalid/image.png"/></foreignObject><path d="M1 1h16v16H1z" fill="#123456" onmouseover="window.__pptxSvgExecuted=3"/><use href="https://example.invalid/art.svg#remote"/></svg>';
      const host = document.querySelector('.vibe-editor__engine');
      for (const [name, detail] of [
        ['vibeoffice-react-diagram-menu', { left: 20, bottom: 20, kind: 'layout', groups: [{ id: 'test', label: '', choices: [{ id: 'fixture', label: 'Fixture', preview: hostile, selected: false }] }], choose() {} }],
        ['vibeoffice-react-line-menu', { left: 20, bottom: 20, kind: 'weight', choices: [{ id: 'fixture', label: 'Fixture', svg: hostile, value: 1 }], choose() {} }],
      ]) host.dispatchEvent(new CustomEvent(name, { detail, bubbles: true }));
      L.app.reactShapeMenuCatalog = () => [{ kind: 'fixture', label: 'Fixture', svg: hostile }];
      const slot = document.createElement('div');
      slot.id = 'react-pptx-shape-grid-root';
      document.querySelector('.vibeoffice-legacy-root[data-vibeoffice-engine-root]').append(slot);
    })()`);
    const result = await page.evaluate("new Promise((resolve)=>{const end=Date.now()+3000;const poll=()=>{const diagram=document.querySelector('.vibeoffice-diagram-menu svg[data-react-owned=pptx-diagram-menu]'),line=document.querySelector('.vibeoffice-line-style-menu svg[data-react-owned=pptx-line-style-menu]'),shape=document.querySelector('#react-pptx-shape-grid-root svg[data-react-owned=pptx-shape-grid]');if(diagram&&line&&shape){const inspect=(svg)=>({scripts:svg.querySelectorAll('script').length,foreignObjects:svg.querySelectorAll('foreignObject').length,handlers:[...svg.querySelectorAll('*')].filter((node)=>[...node.attributes].some((attribute)=>/^on/i.test(attribute.name))).length,path:!!svg.querySelector('path'),externalHref:svg.querySelector('[href]')?.getAttribute('href')??null});resolve({diagram:inspect(diagram),line:inspect(line),shape:inspect(shape),executed:window.__pptxSvgExecuted})}else if(Date.now()>end)resolve(false);else setTimeout(poll,20)};poll()})");
    const clean = { scripts: 0, foreignObjects: 0, handlers: 0, path: true, externalHref: null };
    assert.deepEqual(result, {
      diagram: clean,
      line: clean,
      shape: clean,
      executed: 0,
    });
    assert.deepEqual(page.errors, []);
    assert.deepEqual(page.requests, []);
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});
