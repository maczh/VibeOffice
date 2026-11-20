import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import net from 'node:net';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage } from '../../tools/ooxml/cdp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

test('PPTX selection controls are React-owned and retain their resize handles', { timeout: 60000 }, async () => {
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
    page = await openPage(undefined, { timeout: 30000, bypassServiceWorker: true });
    await page.evaluate("document.querySelector('.suite-app-links .suite-app-mark.pptx')?.closest('button')?.click()");
    let result = { ready: false, shape: false };
    for (let i = 0; i < 80; i++) {
      result = await page.evaluate(`(() => {
        if (!document.querySelector('.suite-editor') || !window.L?.ed || !window.L?.pres?.slides?.length) return { ready: false, shape: false };
        const shape = L.ed.slide().shapes[0];
        if (!shape) return { ready: true, shape: false };
        if (!window.__selectionOverlayTestSelected) { window.__selectionOverlayTestSelected = shape.id; L.ed.select(shape.id); }
        const frame = document.querySelector('.ed-overlay .sel-frame[data-sid="' + CSS.escape(shape.id) + '"]');
        return { ready: true, shape: true, owned: frame?.dataset.reactOwned === 'pptx-selection-control', handles: frame?.querySelectorAll('.hd').length || 0, frame: !!frame };
      })()`);
      if (result.frame) break;
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
    assert.deepEqual(result, { ready: true, shape: true, owned: true, handles: 9, frame: true }, `React-owned PPTX selection overlay: ${JSON.stringify({ result, errors: page.errors })}`);
    const reactThumbnails = await page.evaluate(`new Promise((resolve) => {
      L.app.toggleLeft(true);
      L.panes.left.setTab('slides');
      L.panes.left.render?.();
      const end=Date.now()+3000;
      const poll=()=>{const nodes=[...document.querySelectorAll('.th-frame > [data-react-owned="pptx-slide-artwork"]')];if(nodes.length)resolve({count:nodes.length,tag:nodes[0].tagName,className:nodes[0].className,children:nodes[0].childElementCount});else if(Date.now()>end)resolve({count:0});else setTimeout(poll,25)};
      poll();
    })`);
    assert.ok(reactThumbnails.count > 0, `PPTX slide thumbnails render as React-owned element trees: ${JSON.stringify(reactThumbnails)}`);
    assert.equal(reactThumbnails.tag, 'DIV');
    assert.equal(reactThumbnails.className, 'thumb-box');
    assert.ok(reactThumbnails.children > 0, 'the React-owned thumbnail retains rendered slide artwork');
    const sorterArtwork = await page.evaluate(`new Promise((resolve) => {
      L.app.setView('sorter');
      const end=Date.now()+3000;
      const poll=()=>{const nodes=[...document.querySelectorAll('.sorter-grid .so-frame > [data-react-owned="pptx-slide-artwork"]')];if(nodes.length){L.app.setView('normal');resolve({count:nodes.length,className:nodes[0].className})}else if(Date.now()>end){L.app.setView('normal');resolve({count:0})}else setTimeout(poll,25)};
      poll();
    })`);
    assert.ok(sorterArtwork.count > 0, `PPTX Slide Sorter artwork uses React-owned element trees: ${JSON.stringify(sorterArtwork)}`);
    assert.equal(sorterArtwork.className, 'thumb-box');
    const masterArtwork = await page.evaluate(`new Promise((resolve) => {
      L.app.setView('master');
      const end=Date.now()+3000;
      const poll=()=>{const nodes=[...document.querySelectorAll('.thumbs[aria-label="Slide Masters"] .th-frame > [data-react-owned="pptx-slide-artwork"]')];if(nodes.length){L.app.setView('normal');resolve({count:nodes.length,className:nodes[0].className})}else if(Date.now()>end){L.app.setView('normal');resolve({count:0})}else setTimeout(poll,25)};
      poll();
    })`);
    assert.ok(masterArtwork.count > 0, `PPTX Master thumbnails use React-owned element trees: ${JSON.stringify(masterArtwork)}`);
    assert.equal(masterArtwork.className, 'thumb-box');
    const taskPreviews = await page.evaluate(`new Promise((resolve) => {
      const waitFor=(selector,done)=>{const end=Date.now()+3000;const poll=()=>{const node=document.querySelector(selector);if(node)done(node);else if(Date.now()>end)done(null);else setTimeout(poll,25)};poll()};
      L.app.toggleTask(true);
      const navigationIcons=document.querySelectorAll('.tp-nav > [data-react-owned="pptx-svg-preview"]').length;
      L.panes.task.show('getting-started');
      waitFor('.tp-link svg[data-react-owned="pptx-svg-preview"]',icon=>{
        const iconResult=icon?{tag:icon.tagName,hasArtwork:icon.childElementCount>0}:null;
        L.panes.task.show('design');
      waitFor('.dsn-preview > [data-react-owned="pptx-svg-preview"]',design=>{
        const designResult=design?{tag:design.tagName,className:design.className,hasArtwork:design.childElementCount>0}:null;
        L.panes.task.show('layout');
        waitFor('.lay-it > [data-react-owned="pptx-svg-preview"]',layout=>{
          const layoutResult=layout?{tag:layout.tagName,hasArtwork:layout.childElementCount>0}:null;
          L.app.toggleTask(false);
          resolve({navigationIcons,icon:iconResult,design:designResult,layout:layoutResult});
        });
      });
      });
    })`);
    assert.deepEqual(taskPreviews, { navigationIcons: 4, icon: { tag: 'svg', hasArtwork: true }, design: { tag: 'DIV', className: 'thumb-box', hasArtwork: true }, layout: { tag: 'svg', hasArtwork: true } }, 'PPTX pane chrome and task-pane icons, design and layout previews render as React-owned element trees');
    const notesArtwork = await page.evaluate(`new Promise((resolve) => {
      L.app.setView('notes');
      const end=Date.now()+3000;
      const poll=()=>{const node=document.querySelector('.notespage .np-slide > [data-react-owned="pptx-slide-artwork"]');if(node){const result={tag:node.tagName,className:node.className,children:node.childElementCount};L.app.setView('normal');resolve(result)}else if(Date.now()>end){L.app.setView('normal');resolve(null)}else setTimeout(poll,25)};
      poll();
    })`);
    assert.ok(notesArtwork, 'Notes Page slide artwork mounts');
    assert.equal(notesArtwork.tag, 'DIV');
    assert.equal(notesArtwork.className, 'thumb-box');
    assert.ok(notesArtwork.children > 0, 'Notes Page retains the slide preview artwork');
    const thumbnailImageFallback = await page.evaluate(`new Promise((resolve) => {
      const slide=L.ed.slide(),originalShapes=[...slide.shapes],blob=new Blob(['not an image'],{type:'image/png'}),mediaId=L.media.add(blob,'react-thumbnail-fallback.png');
      const cleanup=()=>{slide.shapes=originalShapes;L.ed.render();L.panes.left.render?.();const media=L.media.get(mediaId);if(media?.url)URL.revokeObjectURL(media.url);L.media.all().delete(mediaId)};
      slide.shapes.push(L.model.newImage(slide,mediaId,30,30,120,80));
      L.ed.render();L.panes.left.render?.();
      const end=Date.now()+3000;
      const poll=()=>{const missing=document.querySelector('.th-frame .pic.missing .pic-missing');if(missing){const result={text:missing.textContent,image:!!missing.parentElement?.querySelector('img')};cleanup();resolve(result)}else if(Date.now()>end){cleanup();resolve(null)}else setTimeout(poll,25)};
      poll();
    })`);
    assert.deepEqual(thumbnailImageFallback, { text: 'Picture', image: false }, 'React thumbnails retain the broken-image fallback and its localized label');
    const backgroundOwnership = await page.evaluate(`new Promise((resolve) => {
      const slide=L.ed.slide(), had=Object.prototype.hasOwnProperty.call(slide,'bg'), original=slide.bg;
      slide.bg={t:'solid',c:'#123456'};
      const restore=()=>{if(had)slide.bg=original;else delete slide.bg;L.ed.render()};
      L.ed.render();
      const end=Date.now()+2000;
      const poll=()=>{const target=document.querySelector('.ed-host > .sl.sl-edit > .sl-bg'),painted=target?.querySelector('[data-react-owned="pptx-slide-background"]');if(painted?.style.backgroundColor==='rgb(18, 52, 86)'){const state={painted:true,color:painted.style.backgroundColor,engineImage:target.style.backgroundImage,engineColor:target.style.backgroundColor};restore();resolve(state)}else if(Date.now()>end){const color=painted?.style.backgroundColor;restore();resolve({painted:!!painted,color})}else setTimeout(poll,20)};
      poll();
    })`);
    assert.equal(backgroundOwnership.painted, true, `React renders the live slide background and refreshes it after the engine rerenders the slide: ${JSON.stringify(backgroundOwnership)}`);
    assert.equal(backgroundOwnership.color, 'rgb(18, 52, 86)');
    assert.deepEqual({ image: backgroundOwnership.engineImage, color: backgroundOwnership.engineColor }, { image: 'none', color: 'transparent' }, 'React takes over the painted background instead of duplicating the engine layer');
    const handle = await page.evaluate(`new Promise((resolve) => {
      const shapeId=L.ed.slide().shapes[0].id;
      L.ed.select(shapeId);
      const end=Date.now()+1000;
      const poll=()=>{const node=document.querySelector('.ed-overlay .sel-frame[data-sid="'+CSS.escape(shapeId)+'"] .hd[data-h="se"]'),box=node?.getBoundingClientRect();if(box)resolve({pointerEvents:getComputedStyle(node).pointerEvents});else if(Date.now()>end)resolve(null);else setTimeout(poll,20)};
      poll();
    })`);
    assert.deepEqual(handle, { pointerEvents: 'auto' }, 'React preserves pointer handling on the selected shape resize control');
    const rulerReady = await page.evaluate(`new Promise((resolve) => {
      window.__pptxRulerPreference = L.app.opts.ruler;
      L.app.opts.ruler = true;
      L.ed.drawRulers();
      const end=Date.now()+2000;
      const poll=()=>{const horizontal=document.querySelector('.ruler-h canvas[data-react-owned="pptx-ruler"]'),vertical=document.querySelector('.ruler-v canvas[data-react-owned="pptx-ruler"]');horizontal&&vertical?resolve({horizontal:horizontal.width>0,vertical:vertical.height>0,visible:!document.querySelector('.ruler-h').hidden&&!document.querySelector('.ruler-v').hidden}):Date.now()>end?resolve({timeout:true,show:L.ed.reactRulerCatalog?.().show,bus:typeof L.bus?.on,horizontal:!!document.querySelector('.ruler-h'),vertical:!!document.querySelector('.ruler-v'),errors:window.__rulerErrors}):setTimeout(poll,20)};
      poll();
    })`);
    assert.deepEqual(rulerReady, { horizontal: true, vertical: true, visible: true }, 'React renders both ruler canvases from the engine geometry catalog');
    const rulerRefresh = await page.evaluate(`new Promise((resolve) => {
      const original=L.ed.reactRulerCatalog;
      const before=document.querySelector('.ruler-h canvas[data-react-owned="pptx-ruler"]').style.width;
      L.ed.reactRulerCatalog=()=>({...original(),width:original().width+37});
      L.bus.emit('pptx-rulers');
      const expected=(parseFloat(before)+37)+'px';
      const end=Date.now()+2000;
      const poll=()=>{const canvas=document.querySelector('.ruler-h canvas[data-react-owned="pptx-ruler"]');if(canvas?.style.width===expected)resolve({before,after:canvas.style.width});else if(Date.now()>end)resolve({before,after:canvas?.style.width,expected});else setTimeout(poll,20)};
      poll();
      setTimeout(()=>{L.ed.reactRulerCatalog=original},0);
    })`);
    assert.equal(rulerRefresh.after, `${parseFloat(rulerRefresh.before) + 37}px`, 'React redraws the ruler canvas when the engine publishes updated geometry');
    await page.evaluate('L.app.opts.ruler=window.__pptxRulerPreference;L.ed.drawRulers();delete window.__pptxRulerPreference');
    const gridReady = await page.evaluate(`new Promise((resolve) => {
      window.__pptxGridBefore = { ...L.ed.grid };
      Object.assign(L.ed.grid, { show: true, guides: true, gx: null, gy: null });
      L.ed.layout();
      const end=Date.now()+2000;
      const poll=()=>{const grid=document.querySelector('.ed-grid [data-react-owned="pptx-grid"]'),guides=document.querySelectorAll('.ed-grid [data-react-owned="pptx-guide"]');grid&&guides.length===2?resolve({texture:grid.style.backgroundImage.includes('radial-gradient'),guides:guides.length,axes:[...guides].map((guide)=>guide.classList.contains('gv')?'x':'y')}):Date.now()>end?resolve(false):setTimeout(poll,20)};
      poll();
    })`);
    assert.deepEqual(gridReady, { texture: true, guides: 2, axes: ['x', 'y'] }, 'React renders the grid texture and both reference guides from engine state');
    await page.evaluate('Object.assign(L.ed.grid,window.__pptxGridBefore);delete window.__pptxGridBefore;L.ed.layout()');
    assert.deepEqual(page.errors, []);
    await page.evaluate("document.querySelector('.suite-editor .back')?.click()");
    const layersDisposed = await page.evaluate(`new Promise((resolve) => { const end=Date.now()+5000; const poll=()=>{if(document.querySelector('.suite-start')) resolve(!document.querySelector('[data-react-owned="pptx-selection-overlay"]')&&!document.querySelector('[data-react-owned="pptx-slide-background"]')); else if(Date.now()>end) resolve(false); else setTimeout(poll,25)};poll(); })`);
    assert.equal(layersDisposed, true, 'unmounting the PPTX editor removes React-owned selection and background layers');
  } finally {
    if (page) await page.close();
    server.kill();
    if (previousBase === undefined) delete process.env.BASE; else process.env.BASE = previousBase;
  }
});
