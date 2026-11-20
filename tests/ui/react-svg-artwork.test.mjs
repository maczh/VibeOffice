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
