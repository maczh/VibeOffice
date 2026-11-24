import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage } from '../../tools/ooxml/cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

test('React Start Center retains the classic app navigation and unframed template artwork', { timeout: 60000 }, async () => {
  const probe = net.createServer();
  await new Promise((resolve) => probe.listen(0, '127.0.0.1', resolve));
  const port = probe.address().port;
  await new Promise((resolve) => probe.close(resolve));
  const base = `http://127.0.0.1:${port}`;
  const server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port), path.join(root, 'suite/dist')], { stdio: 'ignore' });
  const previousBase = process.env.BASE;
  process.env.BASE = base;
  let page;
  try {
    for (let n = 0; n < 40 && !await fetch(base).then((response) => response.ok).catch(() => false); n++) await new Promise((resolve) => setTimeout(resolve, 100));
    page = await openPage(undefined, {
      timeout: 15000,
      bypassServiceWorker: true,
      preload: "(() => { const open=IDBFactory.prototype.open,remove=IDBFactory.prototype.deleteDatabase; window.__legacyRecentDbActions=[]; IDBFactory.prototype.open=function(name,...args){if(String(name)==='vibeoffice')window.__legacyRecentDbActions.push('open');return open.call(this,name,...args)}; IDBFactory.prototype.deleteDatabase=function(name){if(String(name)==='vibeoffice')window.__legacyRecentDbActions.push('delete');return remove.call(this,name)} })()",
    });
    const visual = await page.evaluate(`new Promise((resolve) => {
      const end=Date.now()+5000;
      const poll=()=>{
        const heading=document.querySelector('.suite-app-links h2');
        const cards=[...document.querySelectorAll('.suite-template-thumb')];
        const imageCard=cards.find((card)=>card.querySelector('img'));
        const image=imageCard?.querySelector('img');
        if(heading&&cards.length&&image?.complete&&image.naturalWidth){
          resolve({ready:true,appsHeading:heading.textContent.trim(),apps:[...document.querySelectorAll('.suite-app-links button b')].map((node)=>node.textContent.trim()),icons:[...document.querySelectorAll('.suite-app-links .suite-app-icon')].map((node)=>({tag:node.tagName,width:node.getAttribute('width'),height:node.getAttribute('height')})),actions:[...document.querySelectorAll('.suite-actions .suite-action')].map((button)=>({text:button.textContent.trim(),icon:button.querySelector('svg')?.getAttribute('width')})),footerIcons:document.querySelectorAll('.suite-side-links svg').length,version:document.querySelector('.suite-version')?.textContent.trim(),feedbackPadding:getComputedStyle(document.querySelector('.suite-side-links a')).padding,templateBorder:getComputedStyle(cards[0]).borderWidth,imageNatural:[image.naturalWidth,image.naturalHeight],imageRendered:[Math.round(image.getBoundingClientRect().width),Math.round(image.getBoundingClientRect().height)],legacyRecentDbActions:window.__legacyRecentDbActions.slice()});
        } else if(Date.now()>end) resolve({ready:false,body:document.body.innerText.slice(0,700),errors:window.__capturedErrors||[]});
        else setTimeout(poll,30);
      };
      poll();
    })`);
    assert.equal(visual.ready, true, `Start Center gallery and assets are ready: ${JSON.stringify(visual)}`);
    assert.equal(visual.appsHeading, 'Apps');
    assert.deepEqual(visual.apps, ['Wordx', 'Excelx', 'PPTX']);
    assert.deepEqual(visual.icons, [{ tag: 'svg', width: '24', height: '30' }, { tag: 'svg', width: '24', height: '30' }, { tag: 'svg', width: '24', height: '30' }]);
    assert.deepEqual(visual.actions, [{ text: 'Open File…', icon: '20' }, { text: 'Create File', icon: '20' }], 'the primary sidebar actions retain the legacy icons and labels');
    assert.ok(visual.footerIcons >= 4 && visual.version, 'the sidebar footer retains its app icons and version label');
    assert.equal(visual.feedbackPadding, '6px 8px', 'footer link hit area matches the classic Start Center spacing');
    assert.equal(visual.templateBorder, '0px', 'the React gallery does not add a card frame absent from the classic gallery');
    assert.deepEqual(visual.legacyRecentDbActions, [], 'the React Start Center neither opens nor deletes the legacy Recent Files database');
    assert.ok(visual.imageNatural[0] > 0 && visual.imageNatural[1] > 0 && visual.imageRendered[1] > visual.imageRendered[0], `template previews keep their portrait artwork proportions: ${JSON.stringify(visual)}`);
    assert.deepEqual(page.errors, []);
  } finally {
    await page?.close().catch(() => {});
    server.kill('SIGTERM');
    if (previousBase === undefined) delete process.env.BASE;
    else process.env.BASE = previousBase;
  }
});
