/* Wordx — dialogs, part 3: Format Object (pictures, AutoShapes, text boxes, WordArt), Fill Effects,
 * Tools (Word Count, AutoSummarize, Hyphenation, Envelopes and Labels, Letter Wizard, Macros, Templates,
 * AutoCorrect, Customize, Options) and File (Properties, Print, Save As, New from template, Zoom,
 * Paste Special, About).
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const { h } = L;
  const ui = L.ui;
  const G = L.dlg;
  const doc = () => D.doc;
  const A = () => L.app;

  /* ================= Format Object ================= */
  /* the line dashes (drawing.js DASH) and fills as shown */
  const DASH_NAMES = { solid: L.N('Solid'), sysDot: L.N('Round Dot'), sysDash: L.N('Square Dot'), dot: L.N('Dot'), dash: L.N('Dash'), dashDot: L.N('Dash Dot'), lgDash: L.N('Long Dash'), lgDashDot: L.N('Long Dash Dot'), lgDashDotDot: L.N('Long Dash Dot Dot'), sysDashDot: L.N('Short Dash Dot'), sysDashDotDot: L.N('Short Dash Dot Dot') };
  const FILL_NAMES = { none: L.N('No Fill'), grad: L.N('Gradient'), patt: L.N('Pattern'), img: L.N('Picture'), solid: L.N('Color') };
  G.formatObject = function (tab) {
    const it0 = L.drawing.selectedItem();
    if (!it0) return;
    const direct = !!(window.VO && VO.reactDirectHost);
    const it = L.clone(it0);
    const isImg = it.t === 'img', isShape = it.t === 'shape', isTb = !!it.tb, isWA = !!it.wordart, isLine = isShape && (it.geom === 'line' || L.geom.isLine(it.geom));
    const title = isImg ? 'Format Picture' : isTb ? 'Format Text Box' : isWA ? 'Format LettersArt' : it.t === 'chart' ? 'Format Object' : 'Format AutoShape';
    /* Colors and Lines */
    let fill = it.fill ? L.clone(it.fill) : { t: 'none' };
    const fillBtn = G.color(fill.t === 'solid' ? String(fill.c || '#FFFFFF').replace('#', '') : null, { id: 'fo-fill', mode: 'fill', automatic: false, noneText: L.t(FILL_NAMES[fill.t] && fill.t !== 'solid' ? FILL_NAMES[fill.t] : 'No Fill'), onChange: (v) => { fill = v ? { t: 'solid', c: '#' + v, a: 1 - transp.get() / 100 } : { t: 'none' }; } });
    const fx = ui.button('Fill &Effects...', async () => { const f = await G.fillEffects(fill); if (f) { fill = f; fillBtn.set(null); fillBtn.querySelector('.cpk-l').textContent = L.t(f.t !== 'none' && FILL_NAMES[f.t] || L.N('Fill')); } }, { class: 'btn small' });
    const transp = ui.spin({ value: Math.round((1 - (fill.a == null ? 1 : fill.a)) * 100), min: 0, max: 100, step: 5, dec: 0, unit: ' %', id: 'fo-tr', onChange: (v) => { if (fill.t === 'solid') fill.a = 1 - v / 100; } });
    const ln = isImg ? (it.border ? { c: '#' + (it.border.color || '000000'), w: it.border.sz || 0.75, dash: 'solid' } : { t: 'none' }) : it.line ? L.clone(it.line) : { t: 'none' };
    const lineBtn = G.color(ln.t === 'none' || !ln.c ? null : String(ln.c).replace('#', ''), { id: 'fo-line', mode: 'line', automatic: false, noneText: L.t('No Line') });
    const dash = ui.select(Object.keys(L.drawing.DASH).map((k) => [k, DASH_NAMES[k] || k]), ln.dash || 'solid', null, { id: 'fo-dash' });
    const weight = ui.spin({ value: ln.w || 0.75, min: 0, max: 1584, step: 0.25, dec: 2, unit: ' pt', id: 'fo-wt' });
    const arrows = [['none', L.N('(none)')], ['triangle', L.N('Triangle')], ['stealth', L.N('Stealth')], ['diamond', L.N('Diamond')], ['oval', L.N('Oval')], ['arrow', L.N('Arrow')]];
    const head = ui.select(arrows, (ln.head && ln.head.type) || 'none', null, { id: 'fo-ah' });
    const tail = ui.select(arrows, (ln.tail && ln.tail.type) || 'none', null, { id: 'fo-at' });
    const colorsRoot = direct ? h('div', { id: 'react-wordx-format-object-colors-root' }) : null;
    const colorsBody = direct ? h('div', { class: 'col' }, colorsRoot) : null;
    let directColors = { fill: fill.t === 'solid' ? String(fill.c || '#FFFFFF') : null, line: ln.t === 'none' || !ln.c ? null : String(ln.c), transparency: Math.round((1 - (fill.a == null ? 1 : fill.a)) * 100), dash: ln.dash || 'solid', weight: ln.w || 0.75, head: (ln.head && ln.head.type) || 'none', tail: (ln.tail && ln.tail.type) || 'none' };
    const colorCatalog = () => ({ fill: directColors.fill, fillType: fill.t || 'none', line: directColors.line, transparency: directColors.transparency, dash: directColors.dash, dashes: Object.keys(L.drawing.DASH).map((value) => ({ value, label: DASH_NAMES[value] || value })), weight: directColors.weight, head: directColors.head, tail: directColors.tail, arrows: arrows.map(([value, label]) => ({ value, label })), showFill: !isImg, showArrows: isLine || (isShape && !isTb && !isWA) });
    G.reactFormatObjectColorsCatalog = colorCatalog;
    G.reactFormatObjectColorsSet = (key, value) => {
      if (key === 'transparency') { directColors.transparency = Math.max(0, Math.min(100, Number(value) || 0)); if (fill.t === 'solid') fill.a = 1 - directColors.transparency / 100; }
      else if (key === 'dash' && L.drawing.DASH[value]) directColors.dash = value;
      else if (key === 'weight') directColors.weight = Math.max(0, Math.min(1584, Number(value) || 0));
      else if (key === 'head' || key === 'tail') { if (arrows.some(([item]) => item === value)) directColors[key] = value; }
      return colorCatalog();
    };
    G.reactFormatObjectColorsAction = async (action, anchor) => {
      if (action === 'fill') ui.colorMenu(anchor, { mode: 'fill', effects: true }, async (value) => { if (value && value.effects) { const result = await G.fillEffects(fill); if (result) fill = result; } else fill = value && value.none ? { t: 'none' } : typeof value === 'string' ? { t: 'solid', c: value.toUpperCase(), a: 1 - directColors.transparency / 100 } : fill; directColors.fill = fill.t === 'solid' ? fill.c : null; L.bus.emit('wordx-format-object-colors-refresh'); });
      else if (action === 'line') ui.colorMenu(anchor, { mode: 'line', automatic: false }, (value) => { directColors.line = typeof value === 'string' ? value : null; L.bus.emit('wordx-format-object-colors-refresh'); });
      else if (action === 'effects') { const result = await G.fillEffects(fill); if (result) { fill = result; directColors.fill = fill.t === 'solid' ? fill.c : null; L.bus.emit('wordx-format-object-colors-refresh'); } }
      return colorCatalog();
    };
    /* Size */
    const hgt = G.len(it.h || 72, { id: 'fo-h', min: 1, onChange: (v) => { if (lock.input.checked && it.h) { wid.set((it.w * v) / it.h); scW.set(Math.round((wid.get() / (it.natW || it.w)) * 100)); } scH.set(Math.round((v / (it.natH || it.h)) * 100)); } });
    const wid = G.len(it.w || 72, { id: 'fo-w', min: 1, onChange: (v) => { if (lock.input.checked && it.w) { hgt.set((it.h * v) / it.w); scH.set(Math.round((hgt.get() / (it.natH || it.h)) * 100)); } scW.set(Math.round((v / (it.natW || it.w)) * 100)); } });
    const rot = ui.spin({ value: it.rot || 0, min: -3600, max: 3600, step: 1, dec: 1, unit: '°', id: 'fo-rot' });
    const scH = ui.spin({ value: 100, min: 1, max: 5000, step: 1, dec: 0, unit: ' %', id: 'fo-sh', onChange: (v) => { hgt.set(((it.natH || it.h) * v) / 100); if (lock.input.checked) { wid.set(((it.natW || it.w) * v) / 100); scW.set(v); } } });
    const scW = ui.spin({ value: 100, min: 1, max: 5000, step: 1, dec: 0, unit: ' %', id: 'fo-sw', onChange: (v) => { wid.set(((it.natW || it.w) * v) / 100); if (lock.input.checked) { hgt.set(((it.natH || it.h) * v) / 100); scH.set(v); } } });
    if (it.natW) { scW.set(Math.round((it.w / it.natW) * 100)); scH.set(Math.round((it.h / it.natH) * 100)); }
    let aspectLocked = isImg || it.t === 'chart' || !!it.lockAspect;
    const lock = ui.check('Lock &aspect ratio', aspectLocked, (value) => { aspectLocked = value; });
    const reset = ui.button('Re&set', () => { if (it.natW) { wid.set(it.natW); hgt.set(it.natH); scW.set(100); scH.set(100); } rot.set(0); }, { class: 'btn small' });
    const sizeRoot = direct ? h('div', { id: 'react-wordx-format-object-size-root' }) : null;
    const sizeBody = direct ? h('div', { class: 'col' }, sizeRoot) : null;
    let directSize = null;
    G.reactFormatObjectSizeCatalog = () => ({ width:it.w||72,height:it.h||72,rotation:it.rot||0,locked:aspectLocked,naturalWidth:it.natW||it.w||72,naturalHeight:it.natH||it.h||72,hasNatural:!!it.natW });
    /* Layout */
    let wrap = it.float ? it.float.wrap || 'none' : 'inline';
    const wrapBox = h('div', { class: 'preset-grid' });
    const wrapSvg = { inline: '<svg width="40" height="36"><path d="M2 6h36M2 12h8M30 12h8M2 18h36M2 24h36M2 30h36" stroke="#999"/><rect x="12" y="8" width="16" height="10" fill="#9cf" stroke="#333"/></svg>', square: '<svg width="40" height="36"><path d="M2 6h36M2 12h10M28 12h10M2 18h10M28 18h10M2 24h10M28 24h10M2 30h36" stroke="#999"/><rect x="13" y="10" width="14" height="16" fill="#9cf" stroke="#333"/></svg>', tight: '<svg width="40" height="36"><path d="M2 6h36M2 12h12M26 12h12M2 18h9M29 18h9M2 24h12M26 24h12M2 30h36" stroke="#999"/><circle cx="20" cy="18" r="8" fill="#9cf" stroke="#333"/></svg>', behind: '<svg width="40" height="36"><rect x="10" y="8" width="20" height="20" fill="#9cf" stroke="#333"/><path d="M2 6h36M2 12h36M2 18h36M2 24h36M2 30h36" stroke="#555"/></svg>', none: '<svg width="40" height="36"><path d="M2 6h36M2 12h36M2 18h36M2 24h36M2 30h36" stroke="#999"/><rect x="10" y="8" width="20" height="20" fill="#9cf" stroke="#333"/></svg>' };
    for (const [k, l] of [['inline', L.N('In l&ine with text')], ['square', L.N('&Square')], ['tight', L.N('&Tight')], ['behind', L.N('&Behind text')], ['none', L.N('I&n front of text')]]) {
      const b = h('button', { type: 'button', class: wrap === k ? 'on' : '' }, h('span', { html: wrapSvg[k] }), h('span', { html: ui.ampHTML(l) }));
      b.addEventListener('click', () => { wrap = k; L.$$('button', wrapBox).forEach((x) => x.classList.toggle('on', x === b)); });
      wrapBox.appendChild(b);
    }
    const hAlign = G.radios('fo-ha', [['left', '&Left'], ['center', '&Center'], ['right', '&Right'], ['other', '&Other']], it.float && it.float.posH && it.float.posH.align ? it.float.posH.align : 'other');
    let chosenHAlign = it.float && it.float.posH && it.float.posH.align || 'other';
    const layoutRoot = direct ? h('div', { id: 'react-wordx-format-object-layout-root' }) : null;
    const layoutBody = direct ? h('div', { class: 'col' }, layoutRoot) : null;
    const layoutChoices = [['inline','In line with text'],['square','Square'],['tight','Tight'],['behind','Behind text'],['none','In front of text']];
    G.reactFormatObjectLayoutCatalog = () => ({ wrap, alignment:chosenHAlign, wraps:layoutChoices.map(([id,label])=>({id,label})) });
    G.reactFormatObjectLayoutSet = (key,value) => { if(key==='wrap'&&layoutChoices.some(([id])=>id===value))wrap=value;else if(key==='alignment'&&['left','center','right','other'].includes(value))chosenHAlign=value;else if(key==='advanced')advanced();return G.reactFormatObjectLayoutCatalog(); };
    const adv = ui.button('Ad&vanced...', () => advanced(), { class: 'btn small' });
    let posH = it.float ? L.clone(it.float.posH) : { rel: 'column', off: 0 }, posV = it.float ? L.clone(it.float.posV) : { rel: 'paragraph', off: 0 };
    let dist = it.float ? L.clone(it.float.dist || { t: 0, b: 0, l: 9, r: 9 }) : { t: 0, b: 0, l: 9, r: 9 };
    let moveWith = it.float ? it.float.posV.rel === 'paragraph' : true, overlap = it.float ? it.float.allowOverlap !== false : true;
    function advanced() {
      if (direct) {
        const root=h('div',{id:'react-wordx-format-object-advanced-root'});
        const values={tab:'position',hMode:posH.align?'align':'abs',hAlign:posH.align||'left',hRel:posH.rel||'column',hOff:posH.off||0,vMode:posV.align?'align':'abs',vAlign:posV.align||'top',vRel:posV.rel||'paragraph',vOff:posV.off||0,moveWith,overlap,top:dist.t||0,bottom:dist.b||0,left:dist.l||0,right:dist.r||0};
        const catalog=()=>({...values,hRels:[['column','Column'],['margin','Margin'],['page','Page'],['character','Character']],vRels:[['paragraph','Paragraph'],['margin','Margin'],['page','Page'],['line','Line']]});
        G.reactFormatObjectAdvancedCatalog=catalog;
        G.reactFormatObjectAdvancedSet=(key,value)=>{if(['tab','hMode','hAlign','hRel','vMode','vAlign','vRel'].includes(key))values[key]=value;else if(['moveWith','overlap'].includes(key))values[key]=!!value;else if(['hOff','vOff','top','bottom','left','right'].includes(key))values[key]=Math.max(0,Number(value)||0);return catalog();};
        G.reactFormatObjectAdvancedAction=()=>{
          posH=values.hMode==='align'?{rel:values.hRel,align:values.hAlign}:{rel:values.hRel,off:values.hOff};
          posV=values.vMode==='align'?{rel:values.vRel==='paragraph'?'margin':values.vRel,align:values.vAlign}:{rel:values.vRel,off:values.vOff};
          moveWith=values.moveWith;overlap=values.overlap;dist={t:values.top,b:values.bottom,l:values.left,r:values.right};chosenHAlign=posH.align||'other';
          if(wrap==='inline')wrap='square';
          L.bus.emit('wordx-format-object-layout-refresh');
          return catalog();
        };
        ui.dialog({title:'Advanced Layout',width:460,body:root,buttons:[{label:'OK',primary:true,onClick:()=>G.reactFormatObjectAdvancedAction()},{label:'Cancel'}]});
        return;
      }
      const rels = [['column', L.N('Column|layout')], ['margin', L.N('Margin')], ['page', L.N('Page')], ['character', L.N('Character')]];
      const vrels = [['paragraph', L.N('Paragraph')], ['margin', L.N('Margin')], ['page', L.N('Page')], ['line', L.N('Line')]];
      const hMode = G.radios('fa-hm', [['align', 'Ali&gnment'], ['abs', 'Absolute &position']], posH.align ? 'align' : 'abs');
      const hAl = ui.select([['left', 'Left'], ['center', 'Centered'], ['right', 'Right']], posH.align || 'left', null, { id: 'fa-ha' });
      const hRel = ui.select(rels, posH.rel || 'column', null, { id: 'fa-hr' });
      const hOff = G.len(posH.off || 0, { id: 'fa-ho' });
      const vMode = G.radios('fa-vm', [['align', 'Alig&nment'], ['abs', 'Absolute po&sition']], posV.align ? 'align' : 'abs');
      const vAl = ui.select([['top', 'Top'], ['center', 'Centered'], ['bottom', 'Bottom']], posV.align || 'top', null, { id: 'fa-va' });
      const vRel = ui.select(vrels, posV.rel || 'paragraph', null, { id: 'fa-vr' });
      const vOff = G.len(posV.off || 0, { id: 'fa-vo' });
      const mv = ui.check('&Move object with text', moveWith, null), ov = ui.check('Allo&w overlap', overlap, null);
      const dT = G.len(dist.t || 0, { id: 'fa-dt', min: 0 }), dB = G.len(dist.b || 0, { id: 'fa-db', min: 0 }), dL = G.len(dist.l || 0, { id: 'fa-dl', min: 0 }), dR = G.len(dist.r || 0, { id: 'fa-dr', min: 0 });
      ui.dialog({ title: 'Advanced Layout', width: 460, body: ui.tabs([{ label: 'Picture &Position', body: h('div', { class: 'col' }, ui.group('Horizontal', hMode, G.row(hAl, h('span', { text: 'relative to' }), hRel), G.row(hOff, h('span', { text: 'to the right of' }))), ui.group('Vertical', vMode, G.row(vAl, h('span', { text: 'relative to' }), vRel), G.row(vOff, h('span', { text: 'below' }))), ui.group('Options', mv, ov)) }, { label: '&Text Wrapping', body: ui.group('Distance from text', G.row(G.f('To&p:', dT, 'narrow'), G.f('&Left:', dL, 'narrow')), G.row(G.f('B&ottom:', dB, 'narrow'), G.f('&Right:', dR, 'narrow'))) }]), buttons: [{ label: 'OK', primary: true, onClick: () => {
        posH = hMode.get() === 'align' ? { rel: hRel.value, align: hAl.value } : { rel: hRel.value, off: hOff.get() };
        posV = vMode.get() === 'align' ? { rel: vRel.value === 'paragraph' ? 'margin' : vRel.value, align: vAl.value } : { rel: vRel.value, off: vOff.get() };
        moveWith = mv.input.checked; overlap = ov.input.checked;
        dist = { t: dT.get(), b: dB.get(), l: dL.get(), r: dR.get() };
        chosenHAlign = posH.align || 'other';
        if (!direct) hAlign.set(chosenHAlign);
        if (wrap === 'inline') wrap = 'square';
        if (direct) L.bus.emit('wordx-format-object-layout-refresh');
      } }, { label: 'Cancel' }] });
    }
    /* Picture */
    const crop = Object.assign({ l: 0, t: 0, r: 0, b: 0 }, it.crop || {});
    const cropL = G.len((crop.l || 0) * (it.natW || it.w), { id: 'fo-cl' }), cropR = G.len((crop.r || 0) * (it.natW || it.w), { id: 'fo-cr' }), cropT = G.len((crop.t || 0) * (it.natH || it.h), { id: 'fo-ct' }), cropB = G.len((crop.b || 0) * (it.natH || it.h), { id: 'fo-cb' });
    const color = ui.select([['', 'Automatic'], ['gray', 'Grayscale'], ['bw', 'Black & White'], ['washout', 'Washout']], it.gray ? 'gray' : it.bw ? 'bw' : it.washout ? 'washout' : '', null, { id: 'fo-col' });
    const bright = ui.spin({ value: Math.round(((it.bright || 0) + 1) * 50), min: 0, max: 100, step: 3, dec: 0, unit: ' %', id: 'fo-br' });
    const contrast = ui.spin({ value: Math.round(((it.contrast || 0) + 1) * 50), min: 0, max: 100, step: 3, dec: 0, unit: ' %', id: 'fo-co' });
    const pictureRoot = direct && isImg ? h('div', { id: 'react-wordx-format-object-picture-root' }) : null;
    const pictureBody = direct && isImg ? h('div', { class: 'col' }, pictureRoot) : null;
    let directPicture = null;
    G.reactFormatObjectPictureCatalog = () => ({ crop:{left:(crop.l||0)*(it.natW||it.w),right:(crop.r||0)*(it.natW||it.w),top:(crop.t||0)*(it.natH||it.h),bottom:(crop.b||0)*(it.natH||it.h)},mode:it.gray?'gray':it.bw?'bw':it.washout?'washout':'',brightness:Math.round(((it.bright||0)+1)*50),contrast:Math.round(((it.contrast||0)+1)*50),naturalWidth:it.natW||it.w,naturalHeight:it.natH||it.h });
    /* Text Box */
    const ins = Object.assign({ l: 7.2, t: 3.6, r: 7.2, b: 3.6 }, it.ins || {});
    const iL = G.len(ins.l, { id: 'fo-il', min: 0 }), iR = G.len(ins.r, { id: 'fo-ir', min: 0 }), iT = G.len(ins.t, { id: 'fo-it', min: 0 }), iB = G.len(ins.b, { id: 'fo-ib', min: 0 });
    const anchor = ui.select([['t', 'Top'], ['ctr', 'Center'], ['b', 'Bottom']], it.anchor || 't', null, { id: 'fo-anc' });
    const autofit = ui.check('Resize AutoShape to &fit text', !!it.autofit, null);
    const textBoxRoot = direct ? h('div', { id: 'react-wordx-format-object-textbox-root' }) : null;
    const textBoxBody = direct ? h('div', { class: 'col' }, textBoxRoot) : null;
    let directTextBox = null;
    G.reactFormatObjectTextBoxCatalog = () => ({ margins:{left:ins.l,right:ins.r,top:ins.t,bottom:ins.b},anchor:it.anchor||'t',autofit:!!it.autofit });
    /* Web */
    const alt = h('textarea', { rows: 4, id: 'fo-alt', style: 'width:100%' });
    alt.value = it.alt || it.descr || '';
    const webRoot = direct ? h('div', { id: 'react-wordx-format-object-web-root' }) : null;
    const webBody = direct ? h('div', { class: 'col' }, webRoot) : null;
    G.reactFormatObjectWebCatalog = () => ({ alt: alt.value });
    const tabsDef = [
      { label: 'Colors and Lines', body: direct ? colorsBody : h('div', { class: 'col' }, isImg ? null : ui.group('Fill', G.row(G.f('C&olor:', fillBtn), fx), G.f('&Transparency:', transp)), ui.group('Line', G.row(G.f('Co&lor:', lineBtn), G.f('&Dashed:', dash)), G.f('&Weight:', weight)), isLine || (isShape && !isTb && !isWA) ? ui.group('Arrows', G.row(G.f('&Begin style:', head), G.f('&End style:', tail))) : null) },
      { label: 'Size', body: direct ? sizeBody : h('div', { class: 'col' }, ui.group('Size and rotate', G.row(G.f('H&eight:', hgt, 'narrow'), G.f('Wi&dth:', wid, 'narrow')), G.f('Ro&tation:', rot)), ui.group('Scale', G.row(G.f('&Height:', scH, 'narrow'), G.f('&Width:', scW, 'narrow')), lock), it.natW ? h('div', { class: 'tp-note', text: L.t('Original size — Height: {0}, Width: {1}', D.fmtLen(it.natH), D.fmtLen(it.natW)) }) : null, h('div', { class: 'tp-btns' }, reset)) },
      { label: 'Layout', body: direct ? layoutBody : h('div', { class: 'col' }, ui.group('Wrapping style', wrapBox), ui.group('Horizontal alignment', h('div', { class: 'row' }, ...hAlign.querySelectorAll('.ck'))), h('div', { class: 'tp-btns' }, adv)) },
    ];
    if (isImg) tabsDef.push({ label: 'Picture', body: direct ? pictureBody : h('div', { class: 'col' }, ui.group('Crop from', G.row(G.f('&Left:', cropL, 'narrow'), G.f('T&op:', cropT, 'narrow')), G.row(G.f('Ri&ght:', cropR, 'narrow'), G.f('&Bottom:', cropB, 'narrow'))), ui.group('Image control', G.f('&Color:', color), G.f('Brig&htness:', bright), G.f('Co&ntrast:', contrast))) });
    if (isTb) tabsDef.push({ label: 'Text Box', body: direct ? textBoxBody : h('div', { class: 'col' }, ui.group('Internal margin', G.row(G.f('&Left:', iL, 'narrow'), G.f('&Top:', iT, 'narrow')), G.row(G.f('&Right:', iR, 'narrow'), G.f('&Bottom:', iB, 'narrow'))), G.f('&Vertical alignment:', anchor), autofit) });
    tabsDef.push({ label: 'Web', body: direct ? webBody : h('div', { class: 'col' }, h('label', { for: 'fo-alt', html: ui.ampHTML('&Alternative text:') }), alt, h('div', { class: 'tp-note', text: 'Web browsers display alternative text while pictures are loading or if they are missing. Screen readers read it aloud.' })) });
    const tabIdx = tab === 'picture' ? tabsDef.findIndex((t) => t.label === 'Picture') : tab === 'layout' ? 2 : tab === 'size' ? 1 : 0;
    ui.dialog({
      title, width: 470, body: ui.tabs(tabsDef, Math.max(0, tabIdx)),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        if (direct) {
          const size = sizeRoot.querySelector('#react-wordx-format-object-size-body');
          if (size) directSize = { height:Number(size.querySelector('[data-size-field=height]').value)||1,width:Number(size.querySelector('[data-size-field=width]').value)||1,rotation:Number(size.querySelector('[data-size-field=rotation]').value)||0,locked:size.querySelector('[data-size-field=locked]').checked };
          if (isTb) { const box=textBoxRoot.querySelector('#react-wordx-format-object-textbox-body'); if(box)directTextBox={left:Number(box.querySelector('[data-box-margin=left]').value),right:Number(box.querySelector('[data-box-margin=right]').value),top:Number(box.querySelector('[data-box-margin=top]').value),bottom:Number(box.querySelector('[data-box-margin=bottom]').value),anchor:box.querySelector('#react-wordx-format-object-anchor').value,autofit:box.querySelector('[data-box-autofit]').checked}; }
            if (isImg) { const picture=pictureRoot.querySelector('#react-wordx-format-object-picture-body');if(picture)directPicture={left:Number(picture.querySelector('[data-picture-crop=left]').value)||0,right:Number(picture.querySelector('[data-picture-crop=right]').value)||0,top:Number(picture.querySelector('[data-picture-crop=top]').value)||0,bottom:Number(picture.querySelector('[data-picture-crop=bottom]').value)||0,mode:picture.querySelector('#react-wordx-format-object-picture-mode').value,brightness:Number(picture.querySelector('[data-picture-tone=brightness]').value),contrast:Number(picture.querySelector('[data-picture-tone=contrast]').value)}; }
        }
        L.drawing.modifySelected(title, (x) => {
          if (!isImg) { if (fill.t === 'solid') fill.a = 1 - (direct ? directColors.transparency : transp.get()) / 100; x.fill = fill; }
          const lc = direct ? directColors.line && String(directColors.line).replace('#', '') : lineBtn.get();
          const lineDash = direct ? directColors.dash : dash.value, lineWeight = direct ? directColors.weight : weight.get(), lineHead = direct ? directColors.head : head.value, lineTail = direct ? directColors.tail : tail.value;
          if (isImg) x.border = lc ? { val: lineDash === 'solid' ? 'single' : 'dashed', sz: lineWeight, color: lc } : undefined;
          else x.line = lc ? Object.assign({}, x.line && x.line.t !== 'none' ? x.line : {}, { c: '#' + lc, w: lineWeight, dash: lineDash, head: lineHead === 'none' ? undefined : { type: lineHead, w: 'med', len: 'med' }, tail: lineTail === 'none' ? undefined : { type: lineTail, w: 'med', len: 'med' } }) : { t: 'none' };
          if (x.line && x.line.t !== 'none') delete x.line.t;
          const width = directSize ? directSize.width : wid.get(), height = directSize ? directSize.height : hgt.get(), rotation = directSize ? directSize.rotation : rot.get();
          if (x.t === 'group' && (width !== x.w || height !== x.h)) { const kx = width / x.w, ky = height / x.h; const sc = (g) => { for (const k of g.kids || []) { k.x *= kx; k.y *= ky; k.w *= kx; k.h *= ky; sc(k); } }; sc(x); }
          x.w = L.round(width, 2); x.h = L.round(height, 2);
          x.rot = rotation ? ((rotation % 360) + 360) % 360 : undefined;
          aspectLocked = directSize ? directSize.locked : aspectLocked;
          if (aspectLocked && !isImg) x.lockAspect = true;
          L.drawing.setWrap(x, wrap);
          if (x.float) {
            const alignment = direct ? chosenHAlign : hAlign.get();
            x.float.posH = alignment !== 'other' ? { rel: posH.rel === 'character' ? 'column' : posH.rel || 'column', align: alignment } : posH;
            x.float.posV = posV;
            x.float.dist = dist;
            x.float.allowOverlap = overlap;
          }
          if (isImg) {
            const nw = x.natW || it.w, nh = x.natH || it.h;
            const c = directPicture ? { l:directPicture.left/nw,r:directPicture.right/nw,t:directPicture.top/nh,b:directPicture.bottom/nh } : { l: (cropL.get() || 0) / nw, r: (cropR.get() || 0) / nw, t: (cropT.get() || 0) / nh, b: (cropB.get() || 0) / nh };
            x.crop = c.l || c.r || c.t || c.b ? c : undefined;
            const colorMode=directPicture?directPicture.mode:color.value,brightness=directPicture?directPicture.brightness:bright.get(),contrastValue=directPicture?directPicture.contrast:contrast.get();
            x.gray = colorMode === 'gray' || undefined; x.bw = colorMode === 'bw' || undefined; x.washout = colorMode === 'washout' || undefined;
            x.bright = L.round(brightness / 50 - 1, 2) || undefined; x.contrast = L.round(contrastValue / 50 - 1, 2) || undefined;
          }
          if (isTb) { x.ins = directTextBox ? { l:directTextBox.left,r:directTextBox.right,t:directTextBox.top,b:directTextBox.bottom } : { l: iL.get(), r: iR.get(), t: iT.get(), b: iB.get() }; const vertical=directTextBox?directTextBox.anchor:anchor.value; x.anchor = vertical === 't' ? undefined : vertical; x.autofit = (directTextBox?directTextBox.autofit:autofit.input.checked) || undefined; }
          x.alt = (direct ? webRoot.querySelector('#react-wordx-format-object-alt')?.value : alt.value) || undefined;
        });
      } }, { label: 'Cancel' }],
    });
  };
  /** Fill Effects: resolves with a fill object or null */
  G.fillEffects = function (cur) {
    if (window.VO && VO.reactDirectHost) {
      return new Promise((resolve) => {
        cur = cur || {};
        const presets = { 'Early Sunset': ['#000082', '#66008F', '#BA0066', '#FF0000', '#FF8200'], Horizon: ['#DCEBF5', '#83A7C3', '#768FB9', '#83A7C3', '#FFFFFF'], Ocean: ['#03D4A8', '#21D6E0', '#0087E6', '#005CBF'], Fire: ['#FFF200', '#FF7A00', '#FF0300', '#4D0808'], Silver: ['#FFFFFF', '#E6E6E6', '#7D8496', '#E6E6E6'], Gold: ['#E6DCAC', '#E6D78A', '#C7AC4C', '#E6D78A', '#E6DCAC'], Rainbow: ['#A603AB', '#0819FB', '#1A8D48', '#FFFF00', '#EE3F17', '#E81766'] };
        const state = { tab: cur.t === 'patt' ? 'pattern' : cur.t === 'img' ? 'picture' : 'gradient', colors: 'two', color1: cur.t === 'grad' && cur.stops ? String(cur.stops[0].c) : '#3366FF', color2: cur.t === 'grad' && cur.stops ? String(cur.stops[cur.stops.length - 1].c) : '#FFFFFF', preset: 'Early Sunset', style: 'h', pattern: cur.t === 'patt' ? cur.prst : 'pct50', foreground: cur.t === 'patt' ? cur.fg : '#000000', background: cur.t === 'patt' ? cur.bg : '#FFFFFF', picture: cur.t === 'img' ? cur.media : null, pictureName: cur.t === 'img' ? 'Picture selected' : 'No picture selected' };
        let result = null;
        const catalog = () => ({ ...state, presets: Object.keys(presets), styles: [['h','Horizontal'],['v','Vertical'],['du','Diagonal up'],['dd','Diagonal down'],['corner','From corner'],['center','From center']], patterns: Object.keys(L.drawing.PATTERNS).map((key) => ({ key, name:(L.geom.PATTERN_NAMES||{})[key]||key, svg:L.drawing.patternSVG(key,'#000','#fff',22) })) });
        G.reactFillEffectsCatalog = catalog;
        G.reactFillEffectsSet = (key, value) => { if (key === 'tab' && ['gradient','texture','pattern','picture'].includes(value)) state.tab = value; else if (['colors','color1','color2','preset','style','pattern','foreground','background'].includes(key)) state[key] = value; return catalog(); };
        G.reactFillEffectsAction = async (action, anchor) => {
          if (['color1','color2','foreground','background'].includes(action)) ui.colorMenu(anchor,{mode:'plain',scheme:false},(value)=>{if(typeof value==='string')state[action]=value;L.bus.emit('wordx-fill-effects-refresh');});
          else if (action === 'picture') { const file=(await L.pickFiles('image/*'))[0]; if(file){state.picture=L.media.add(new Blob([await L.readAsArrayBuffer(file)],{type:file.type}),file.name);state.pictureName=file.name;L.bus.emit('wordx-fill-effects-refresh');} }
          else if (action === 'accept') {
            if (state.tab === 'gradient') { const colors=state.colors==='preset'?presets[state.preset]:state.colors==='one'?[state.color1,'#FFFFFF']:[state.color1,state.color2]; const stops=colors.map((color,index)=>({p:index/(colors.length-1),c:color,a:1})); result={t:'grad',ang:{h:90,v:0,du:45,dd:135,corner:45,center:0}[state.style],stops,path:state.style==='center'?'circle':state.style==='corner'?'rect':undefined}; }
            else if (state.tab === 'pattern' || state.tab === 'texture') result={t:'patt',prst:state.pattern,fg:state.foreground,bg:state.background};
            else if (state.tab === 'picture' && state.picture) result={t:'img',media:state.picture};
          }
          return catalog();
        };
        const root = h('div',{id:'react-wordx-fill-effects-root'});
        ui.dialog({title:'Fill Effects',width:420,body:root,buttons:[{label:'OK',primary:true,onClick:()=>G.reactFillEffectsAction('accept')},{label:'Cancel'}],onClose:()=>resolve(result)});
      });
    }
    return new Promise((resolve) => {
      cur = cur || {};
      let res = null;
      const c1 = G.color(cur.t === 'grad' && cur.stops ? String(cur.stops[0].c).replace('#', '') : '3366FF', { id: 'fe-c1', automatic: false });
      const c2 = G.color(cur.t === 'grad' && cur.stops ? String(cur.stops[cur.stops.length - 1].c).replace('#', '') : 'FFFFFF', { id: 'fe-c2', automatic: false });
      const colors = G.radios('fe-cols', [['one', 'O&ne color'], ['two', '&Two colors'], ['preset', 'Pr&eset']], 'two');
      const PRESETS = { 'Early Sunset': ['#000082', '#66008F', '#BA0066', '#FF0000', '#FF8200'], Horizon: ['#DCEBF5', '#83A7C3', '#768FB9', '#83A7C3', '#FFFFFF'], Ocean: ['#03D4A8', '#21D6E0', '#0087E6', '#005CBF'], Fire: ['#FFF200', '#FF7A00', '#FF0300', '#4D0808'], Silver: ['#FFFFFF', '#E6E6E6', '#7D8496', '#E6E6E6'], Gold: ['#E6DCAC', '#E6D78A', '#C7AC4C', '#E6D78A', '#E6DCAC'], Rainbow: ['#A603AB', '#0819FB', '#1A8D48', '#FFFF00', '#EE3F17', '#E81766'] };
      const preset = ui.select([L.N('Early Sunset'), L.N('Horizon'), L.N('Ocean'), L.N('Fire'), L.N('Silver'), L.N('Gold'), L.N('Rainbow')].map((n) => [n, n]), 'Early Sunset', null, { id: 'fe-pre' });   // PRESETS' names
      const style = G.radios('fe-st', [['h', 'Hori&zontal'], ['v', '&Vertical'], ['du', 'Diagonal &up'], ['dd', 'Dia&gonal down'], ['corner', '&From corner'], ['center', 'From ce&nter']], 'h');
      const pattGrid = h('div', { class: 'shape-grid', style: 'grid-template-columns:repeat(8,26px)' });
      let patt = cur.t === 'patt' ? cur.prst : 'pct50';
      const pattName = (k) => (L.geom.PATTERN_NAMES || {})[k] || k;   // the pattern's name as shown (common/geometry.js), else its OOXML name
      for (const k of Object.keys(L.drawing.PATTERNS)) { const b = h('button', { type: 'button', class: k === patt ? 'on' : '', 'data-tip': pattName(k), 'aria-label': pattName(k), html: L.drawing.patternSVG(k, '#000', '#fff', 22) }); b.addEventListener('click', () => { patt = k; L.$$('button', pattGrid).forEach((x) => x.classList.toggle('on', x === b)); }); pattGrid.appendChild(b); }
      const fg = G.color('000000', { id: 'fe-fg', automatic: false }), bg = G.color('FFFFFF', { id: 'fe-bg', automatic: false });
      let pic = cur.t === 'img' ? cur.media : null;
      const picName = h('div', { class: 'tp-note', text: pic ? 'Picture selected' : 'No picture selected' });
      let which = cur.t === 'patt' ? 2 : cur.t === 'img' ? 3 : 0;
      const tabs = ui.tabs([
        { label: '&Gradient', body: h('div', { class: 'col' }, ui.group('Colors', colors, G.row(G.f('Color &1:', c1), G.f('Color &2:', c2)), G.f('Preset colors:', preset)), ui.group('Shading styles', style)), onShow: () => { which = 0; } },
        { label: 'Te&xture', body: h('div', { class: 'tp-note', text: 'Textures are drawn as patterns; choose a pattern or a picture.' }), onShow: () => { which = 1; } },
        { label: '&Pattern', body: h('div', { class: 'col' }, pattGrid, G.row(G.f('&Foreground:', fg), G.f('&Background:', bg))), onShow: () => { which = 2; } },
        { label: 'P&icture', body: h('div', { class: 'col' }, ui.button('Select &Picture...', async () => { const f = (await L.pickFiles('image/*'))[0]; if (!f) return; pic = L.media.add(new Blob([await L.readAsArrayBuffer(f)], { type: f.type }), f.name); picName.textContent = f.name; }), picName), onShow: () => { which = 3; } },
      ], which);
      ui.dialog({ title: 'Fill Effects', width: 420, body: tabs, buttons: [{ label: 'OK', primary: true, onClick: () => {
        if (which === 0) {
          const ang = { h: 90, v: 0, du: 45, dd: 135, corner: 45, center: 0 }[style.get()];
          let stops;
          if (colors.get() === 'preset') { const cs = PRESETS[preset.value]; stops = cs.map((c, i) => ({ p: i / (cs.length - 1), c, a: 1 })); }
          else if (colors.get() === 'one') stops = [{ p: 0, c: '#' + c1.get(), a: 1 }, { p: 1, c: '#FFFFFF', a: 1 }];
          else stops = [{ p: 0, c: '#' + c1.get(), a: 1 }, { p: 1, c: '#' + c2.get(), a: 1 }];
          res = { t: 'grad', ang, stops, path: style.get() === 'center' ? 'circle' : style.get() === 'corner' ? 'rect' : undefined };
        } else if (which === 2 || which === 1) res = { t: 'patt', prst: patt, fg: '#' + fg.get(), bg: '#' + bg.get() };
        else if (which === 3 && pic) res = { t: 'img', media: pic };
      } }, { label: 'Cancel' }], onClose: () => resolve(res) });
    });
  };
  G.drawingGrid = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    if (direct) {
      const root = h('div', { id: 'react-wordx-drawing-grid-root' });
      G.reactDrawingGridCatalog = () => ({ snap: !!A().opts.snapGrid, horizontal: A().opts.gridH || 9, vertical: A().opts.gridV || 9 });
      G.reactDrawingGridAction = (value) => { const opts = A().opts; opts.snapGrid = !!value.snap; opts.gridH = Math.max(1, Number(value.horizontal) || 1); opts.gridV = Math.max(1, Number(value.vertical) || 1); A().saveOpts(); return G.reactDrawingGridCatalog(); };
      ui.dialog({ title: 'Drawing Grid', width: 300, body: root, buttons: [{ label: 'OK', primary: true, onClick: () => { const body = root.querySelector('#react-wordx-drawing-grid-body'); if (body) G.reactDrawingGridAction({ snap: body.querySelector('[name=react-wordx-grid-snap]').checked, horizontal: body.querySelector('#react-wordx-grid-horizontal').value, vertical: body.querySelector('#react-wordx-grid-vertical').value }); } }, { label: 'Cancel' }] });
      return;
    }
    const snap = ui.check('&Snap objects to grid', !!A().opts.snapGrid, null);
    const hs = G.len(A().opts.gridH || 9, { id: 'dg-h', min: 1 }), vs = G.len(A().opts.gridV || 9, { id: 'dg-v', min: 1 });
    ui.dialog({ title: 'Drawing Grid', width: 300, body: h('div', { class: 'col' }, snap, ui.group('Grid settings', G.f('&Horizontal spacing:', hs, 'wide'), G.f('&Vertical spacing:', vs, 'wide'))), buttons: [{ label: 'OK', primary: true, onClick: () => { A().opts.snapGrid = snap.input.checked; A().opts.gridH = hs.get(); A().opts.gridV = vs.get(); A().saveOpts(); } }, { label: 'Cancel' }] });
  };


  /* ================= Format ▸ AutoFormat ================= */
  G.autoFormatDoc = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    if (direct) {
      const root = h('div', { id: 'react-wordx-autoformat-root' });
      G.reactAutoFormatCatalog = () => ({ mode: 'now', kind: 'general', options: { headings: true, lists: true, quotes: A().opts.smartQuotes !== false, dashes: A().opts.dashes !== false, ordinals: A().opts.ordinals !== false, fractions: A().opts.fractions !== false, urls: A().opts.autoUrl !== false, empty: true } });
      G.reactAutoFormatAction = (value) => {
        const cfg = { ...value.options };
        if (value.kind === 'letter') cfg.headings = false;
        const n = autoFormat(cfg);
        if (value.mode === 'review') {
          const reviewRoot=h('div',{id:'react-wordx-autoformat-review-root'});
          G.reactAutoFormatReviewCatalog=()=>({changes:n});
          ui.dialog({ title: 'AutoFormat', width: 380, body: reviewRoot, buttons: [{ label: '&Accept All', primary: true }, { label: '&Reject All', onClick: () => { D.undo(); LY.update(); E.restoreDom(true); } }] });
        }
        else A().status(n === 1 ? L.t('AutoFormat made {0} change.', n) : L.t('AutoFormat made {0} changes.', n), 6000);
        return n;
      };
      ui.dialog({ title: 'AutoFormat', width: 400, body: root, buttons: [{ label: 'OK', primary: true, onClick: () => {
        const body = root.querySelector('#react-wordx-autoformat-body'); if (!body) return;
        const options = Object.fromEntries(Array.from(body.querySelectorAll('[data-autoformat-option]'), (input) => [input.dataset.autoformatOption, input.checked]));
        G.reactAutoFormatAction({ mode: body.querySelector('[name=react-wordx-autoformat-mode]:checked')?.value || 'now', kind: body.querySelector('#react-wordx-autoformat-kind').value, options });
      } }, { label: 'Cancel' }] });
      return;
    }
    const mode = G.radios('af-m', [['now', '&AutoFormat now'], ['review', 'AutoFormat and &review each change']], 'now');
    const kind = ui.select([['general', 'General document'], ['letter', 'Letter'], ['email', 'E-mail']], 'general', null, { id: 'af-k' });
    const o = A().opts;
    const opts = { headings: ui.check('Built-in Heading &styles', true, null), lists: ui.check('&List styles', true, null), quotes: ui.check('"Straight &quotes" with “smart quotes”', o.smartQuotes !== false, null), dashes: ui.check('&Hyphens (--) with dash (—)', o.dashes !== false, null), ordinals: ui.check('&Ordinals (1st) with superscript', o.ordinals !== false, null), fractions: ui.check('&Fractions (1/2) with fraction character', o.fractions !== false, null), urls: ui.check('&Internet paths with hyperlinks', o.autoUrl !== false, null), empty: ui.check('Remove &extra empty paragraphs', true, null) };
    ui.dialog({ title: 'AutoFormat', width: 400, body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'Wordx will automatically format the document.' }), mode, G.f('Please select a document type to help improve the formatting process:', kind, 'wide'), ui.group('Apply / replace', ...Object.values(opts))), buttons: [{ label: 'OK', primary: true, onClick: () => {
      const cfg = {}; for (const k in opts) cfg[k] = opts[k].input.checked;
      if (kind.value === 'letter') cfg.headings = false;
      const n = autoFormat(cfg);
      if (mode.get() === 'review') ui.dialog({ title: 'AutoFormat', width: 380, body: h('div', { class: 'tp-note', text: n === 1 ? L.t('Formatting completed: {0} change. You can accept all of them, or reject all and return the document to its previous state.', n) : L.t('Formatting completed: {0} changes. You can accept all of them, or reject all and return the document to its previous state.', n) }), buttons: [{ label: '&Accept All', primary: true }, { label: '&Reject All', onClick: () => { D.undo(); LY.update(); E.restoreDom(true); } }] });
      else A().status(n === 1 ? L.t('AutoFormat made {0} change.', n) : L.t('AutoFormat made {0} changes.', n), 6000);
    } }, { label: 'Cancel' }] });
  };
  function autoFormat(cfg) {
    const d = doc();
    let n = 0;
    E.edit('AutoFormat', () => {
      D.reindex(d);
      const paras = d.main.paras.slice();
      const textReplace = (p, re, fn) => {
        const t = D.ptext(p);
        const hits = [];
        let m;
        re.lastIndex = 0;
        while ((m = re.exec(t))) { hits.push([m.index, m[0], fn(m, t)]); if (!re.global) break; }
        for (let i = hits.length - 1; i >= 0; i--) {
          const [at, src, rep] = hits[i];
          if (rep == null || rep === src) continue;
          const rp = O.rPrAt(D.pos(p, at + 1));
          O.deleteRange(D.pos(p, at), D.pos(p, at + src.length));
          if (typeof rep === 'string') O.insertText(D.pos(p, at), rep, rp);
          else { O.insertText(D.pos(p, at), rep.text, Object.assign({}, rp, rep.rPr)); }
          n++;
        }
      };
      for (const p of paras) {
        if (cfg.quotes) {
          textReplace(p, /"/g, (m, t) => { const pr = t[m.index - 1]; return !pr || /[\s([{—–\-]/.test(pr) ? '“' : '”'; });
          textReplace(p, /'/g, (m, t) => { const pr = t[m.index - 1]; return !pr || /[\s([{—–\-“]/.test(pr) ? '‘' : '’'; });
        }
        if (cfg.dashes) { textReplace(p, /(\S) ?-- ?(\S)/g, (m) => m[1] + '—' + m[2]); textReplace(p, /(\w) - (\w)/g, (m) => m[1] + ' – ' + m[2]); }
        if (cfg.fractions) textReplace(p, /\b(1\/2|1\/4|3\/4)\b/g, (m) => ({ '1/2': '½', '1/4': '¼', '3/4': '¾' })[m[1]]);
        if (cfg.ordinals) {
          const t = D.ptext(p);
          const re = /\b(\d+)(st|nd|rd|th)\b/g;
          let m;
          const hits = [];
          while ((m = re.exec(t))) hits.push([m.index + m[1].length, m[2].length]);
          for (const [a, l] of hits) { O.setRunProps(D.pos(p, a), D.pos(p, a + l), { vert: 'superscript' }); n++; }
        }
        if (cfg.urls) {
          const t = D.ptext(p);
          const re = /\b((?:https?:\/\/|www\.)[^\s<>"“”]+[^\s<>"“”.,;:!?)\]])/g;
          let m;
          const hits = [];
          while ((m = re.exec(t))) hits.push([m.index, m[1]]);
          for (const [a, u] of hits) { const r = O.rPrAt(D.pos(p, a + 1)); if (r.link) continue; O.setRunProps(D.pos(p, a), D.pos(p, a + u.length), { link: { url: /^www\./.test(u) ? 'http://' + u : u }, style: 'Hyperlink' }); n++; }
        }
      }
      D.reindex(d);
      const blocks = d.main.blocks;
      const plain = (p) => D.ptext(p).replace(/[-]/g, '');
      if (cfg.lists) {
        let bulletNum = null, numNum = null, lastKind = null, touchedNum = false;
        for (let i = 0; i < blocks.length; i++) {
          const p = blocks[i];
          if (p.t !== 'p' || (p.pPr.num && p.pPr.num.id)) { lastKind = null; continue; }
          const t = plain(p);
          const mb = /^\s*([*•\-–>])\s+\S/.exec(t), mn = /^\s*(\d+|[a-zA-Z])[.)]\s+\S/.exec(t);
          if (!mb && !mn) { if (t.trim()) lastKind = null; continue; }
          const kind = mb ? 'b' : 'n';
          const cut = (mb || mn)[0].length - 1;
          O.deleteRange(D.pos(p, 0), D.pos(p, cut));
          D.touch(p);
          if (!touchedNum) { D.touchKey(d, 'numbering'); touchedNum = true; }
          if (kind === 'b') { if (!bulletNum || lastKind !== 'b') bulletNum = D.addNum(d, D.makeBulletAbs('•', 'Symbol', '')); p.pPr.num = { id: bulletNum, lvl: 0 }; }
          else { if (!numNum || lastKind !== 'n') numNum = D.addNum(d, D.makeNumberAbs(/[a-z]/.test(mn[1]) ? 'lowerLetter' : /[A-Z]/.test(mn[1]) ? 'upperLetter' : 'decimal', '%1.')); p.pPr.num = { id: numNum, lvl: 0 }; }
          if (!p.pPr.style || p.pPr.style === 'Normal') p.pPr.style = 'ListParagraph' in d.styles ? D.useStyle(d, 'ListParagraph') : p.pPr.style;
          lastKind = kind;
          n++;
        }
      }
      if (cfg.headings) {
        for (let i = 0; i < blocks.length; i++) {
          const p = blocks[i];
          if (p.t !== 'p' || D.headingLevel(d, p) || (p.pPr.num && p.pPr.num.id) || (p.pPr.style && p.pPr.style !== 'Normal')) continue;
          const t = plain(p).trim();
          const next = blocks[i + 1];
          if (!t || t.length > 70 || /[.,;:!?]$/.test(t) || t.split(/\s+/).length > 10) continue;
          if (!next || next.t !== 'p' || !plain(next).trim() || plain(next).trim().length < t.length) continue;
          const prev = blocks[i - 1];
          if (prev && prev.t === 'p' && plain(prev).trim() && !D.headingLevel(d, prev)) continue;
          D.touch(p);
          p.pPr.style = D.useStyle(d, /^[A-Z0-9 \-:&]+$/.test(t) || i < 3 ? 'Heading1' : 'Heading2');
          n++;
        }
      }
      if (cfg.empty) {
        const cont = d.main;
        let removed = false;
        for (let i = blocks.length - 2; i > 0; i--) {
          const a = blocks[i], b = blocks[i - 1];
          if (a.t === 'p' && b.t === 'p' && !a.sect && !b.sect && !plain(a).trim() && !plain(b).trim() && a.runs.every((r) => r.t === 'text') && b.runs.every((r) => r.t === 'text')) { if (!removed) { D.touchList(cont); removed = true; } blocks.splice(i, 1); n++; }
        }
        if (removed) d._idxDirty = true;
      }
      D.reindex(d);
      return E.sel && D.byId(d, E.sel.f.p.id) ? E.sel : D.pos(D.firstPara(d.main.blocks), 0);
    });
    return n;
  }
  /* File ▸ New ▸ From existing document: open a copy that is not tied to the original file name */
  L.app.newFromExisting = async function () {
    const f = (await L.pickFiles('.docx,.dotx,.docm,.dotm,.htm,.html,.txt,.rtf'))[0];
    if (!f) return;
    await A().openFile(f);
    const nm = 'Document' + (++A().untitled);
    A().fileName = nm;
    A().fileType = 'docx';
    doc().dirty = true;
    A().updateTitle();
  };

  /* ================= Word Count ================= */
  G.wordCount = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const d = doc();
    const selOnly = E.sel && !E.collapsed();
    let includeNotes = !!L.store.get('wcNotes', false);
    const incl = direct ? null : ui.check('Include &footnotes and endnotes', includeNotes, () => { includeNotes = incl.input.checked; L.store.set('wcNotes', includeNotes); draw(); });
    const tbl = h('table', { class: 'wc-table' });
    const draw = () => {
      L.clear(tbl);
      let paras;
      if (selOnly) { const [a, b] = E.range(); paras = D.parasBetween(d, a.p, b.p).map((p) => (p === a.p || p === b.p ? D.para(D.sliceRuns(p, p === a.p ? a.o : 0, p === b.p ? b.o : D.plen(p))) : p)); }
      else { paras = D.allParas(d); if (includeNotes) for (const k of ['fn', 'en']) for (const id in d[k]) D.walk(d[k][id], (b) => { if (b.t === 'p') paras.push(b); }); }
      const st = D.stats(paras);
      let lines = 0;
      if (LY.pages.length) for (const pg of LY.pages) for (const f of pg.el.querySelectorAll('.pg-body .p')) { if (f.closest('.balloons')) continue; lines += LY.lineCount(f); }
      const rows = [[L.N('Pages'), selOnly ? '' : LY.pages.length || 1], [L.N('Words'), st.words], ['Characters (no spaces)', st.charsNoSp], ['Characters (with spaces)', st.chars], [L.N('Paragraphs'), st.paras], [L.N('Lines'), selOnly ? '' : lines || st.paras]];
      for (const [k, v] of rows) if (v !== '') tbl.appendChild(h('tr', null, h('td', { text: k }), h('td', { text: Number(v).toLocaleString() })));
    };
    G.reactWordCountCatalog = () => ({ selection: !!selOnly, includeNotes, rows: Array.from(tbl.querySelectorAll('tr'), (row) => Array.from(row.cells, (cell) => cell.textContent)) });
    G.reactWordCountAction = (value) => { includeNotes = !!value; L.store.set('wcNotes', includeNotes); draw(); return G.reactWordCountCatalog(); };
    const reactRoot = h('div', { id: 'react-wordx-word-count-root' });
    draw();
    ui.dialog({ title: 'Word Count', width: 300, body: direct ? reactRoot : h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: selOnly ? L.t('Statistics: (selection)') : 'Statistics:' }), tbl, incl), buttons: [{ label: 'Close', primary: true }] });
  };

  /* ================= AutoSummarize ================= */
  G.autoSummarize = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const kind = G.radios('as-k', [['highlight', '&Highlight key points'], ['top', '&Insert an executive summary or abstract at the top of the document'], ['new', '&Create a new document and put the summary there'], ['hide', 'Hide &everything but the summary without leaving the original document']], 'highlight');
    const pct = ui.select([[10, '10 sentences'], [20, '20 sentences'], [100, '100 words or less'], [500, '500 words or less'], ['10%', '10%'], ['25%', '25%'], ['50%', '50%'], ['75%', '75%']], '25%', null, { id: 'as-pct' });
    const upd = ui.check('&Update document statistics', true, null);
    const reactRoot = h('div', { id: 'react-wordx-summarize-root' });
    ui.dialog({ title: 'AutoSummarize', width: 460, body: direct ? reactRoot : h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'Wordx has examined the document and picked the sentences most relevant to the main theme.' }), ui.group('Type of summary', kind), ui.group('Length of summary', G.f('&Percent of original:', pct, 'wide')), upd), buttons: [{ label: 'OK', primary: true, onClick: () => direct ? summarize(reactRoot.querySelector('input[name=react-wordx-summary-kind]:checked')?.value || 'highlight', reactRoot.querySelector('#react-wordx-summary-length')?.value || '25%', !!reactRoot.querySelector('#react-wordx-summary-update')?.checked) : summarize(kind.get(), pct.value, upd.input.checked) }, { label: 'Cancel' }] });
  };
  function summarize(kind, len, updStats) {
    const d = doc();
    D.reindex(d);
    const sents = [];
    const freq = new Map();
    const STOP = new Set('a an and are as at be but by for from has have he her his i in is it its of on or our she that the their them there they this to was we were what when which who will with you your not no so if than then also can may'.split(' '));
    for (const p of d.main.paras) {
      const t = D.ptext(p);
      const re = /[^.!?]+[.!?]*\s*/g;
      let m;
      while ((m = re.exec(t))) {
        const s = m[0];
        if (!s.trim()) continue;
        const words = (s.toLowerCase().match(/[\p{L}']+/gu) || []).filter((w) => !STOP.has(w) && w.length > 2);
        for (const w of words) freq.set(w, (freq.get(w) || 0) + 1);
        sents.push({ p, a: m.index, b: m.index + s.trimEnd().length, text: s.trim(), words, idx: sents.length });
      }
    }
    if (!sents.length) return;
    for (const s of sents) s.score = s.words.reduce((a, w) => a + (freq.get(w) || 0), 0) / Math.max(4, s.words.length) + (D.headingLevel(d, s.p) ? 2 : 0) + (s.idx === 0 ? 1 : 0);
    let pick;
    const totalWords = sents.reduce((a, s) => a + s.text.split(/\s+/).length, 0);
    const ranked = sents.slice().sort((a, b) => b.score - a.score);
    if (/%$/.test(len)) { const target = (totalWords * parseFloat(len)) / 100; pick = []; let w = 0; for (const s of ranked) { if (w >= target && pick.length) break; pick.push(s); w += s.text.split(/\s+/).length; } }
    else if (+len >= 100) { pick = []; let w = 0; for (const s of ranked) { const n = s.text.split(/\s+/).length; if (w + n > +len && pick.length) break; pick.push(s); w += n; } }
    else pick = ranked.slice(0, +len);
    pick.sort((a, b) => a.idx - b.idx);
    if (kind === 'highlight' || kind === 'hide') {
      E.edit('AutoSummarize', () => {
        if (kind === 'highlight') for (const s of pick) O.setRunProps(D.pos(s.p, s.a), D.pos(s.p, s.b), { hl: 'yellow' });
        else { const set = new Set(pick); for (const s of sents) if (!set.has(s)) O.setRunProps(D.pos(s.p, s.a), D.pos(s.p, s.b), { hidden: true }); }
        return E.sel;
      });
      const n = pick.length;
      A().status(kind === 'highlight' ? (n === 1 ? L.t('{0} key sentence highlighted. Undo to restore.', n) : L.t('{0} key sentences highlighted. Undo to restore.', n))
        : n === 1 ? L.t('{0} key sentence kept visible; the rest is hidden text. Undo to restore.', n) : L.t('{0} key sentences kept visible; the rest is hidden text. Undo to restore.', n), 8000);
    } else if (kind === 'top') {
      E.edit('AutoSummarize', () => {
        const cont = D.touchList(d.main);
        const title = D.para([D.text('Summary', { b: true })], {});
        const body = D.para([D.text(pick.map((s) => s.text).join(' '))], {});
        cont.blocks.unshift(title, body, D.para());
        d._idxDirty = true;
        return D.pos(title, 0);
      });
    } else {
      const text = pick.map((s) => s.text).join(' ');
      A().newDocument({ build: (nd) => { nd.main.blocks = [D.para([D.text('Summary of ' + A().fileName, { b: true, sz: 14 })]), D.para([D.text(text)])]; } });
    }
    if (updStats) { d.props.description = (pick.map((s) => s.text).join(' ')).slice(0, 255); }
  }

  /* ================= Hyphenation ================= */
  G.hyphenation = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const d = doc();
    const auto = direct ? null : ui.check('&Automatically hyphenate document', !!d.settings.autoHyphen, null);
    const caps = direct ? null : ui.check('Hyphenate words in &CAPS', d.settings.hyphenCaps !== false, null);
    const zone = direct ? null : G.len(d.settings.hyphenZone || 18, { id: 'hy-z', min: 0 });
    const limit = direct ? null : G.num(d.settings.consecHyphen || null, { id: 'hy-l', min: 0, autoLabel: L.t('No limit') });
    const reactRoot = h('div', { id: 'react-wordx-hyphenation-root' });
    G.reactHyphenationCatalog = () => ({ auto: !!d.settings.autoHyphen, caps: d.settings.hyphenCaps !== false, zone: d.settings.hyphenZone || 18, limit: d.settings.consecHyphen || '' });
    ui.dialog({ title: 'Hyphenation', width: 320, body: direct ? reactRoot : h('div', { class: 'col' }, auto, caps, G.f('Hyphenation &zone:', zone, 'wide'), G.f('&Limit consecutive hyphens to:', limit, 'wide')), buttons: [{ label: 'OK', primary: true, onClick: () => { const value = direct ? { auto: !!reactRoot.querySelector('#react-wordx-hyphen-auto')?.checked, caps: !!reactRoot.querySelector('#react-wordx-hyphen-caps')?.checked, zone: Number(reactRoot.querySelector('#react-wordx-hyphen-zone')?.value || 18), limit: Number(reactRoot.querySelector('#react-wordx-hyphen-limit')?.value) || undefined } : { auto: auto.input.checked, caps: caps.input.checked, zone: zone.get(), limit: limit.get() || undefined }; E.edit('Hyphenation', () => { D.touchKey(d, 'settings'); d.settings.autoHyphen = value.auto; d.settings.hyphenCaps = value.caps; d.settings.hyphenZone = value.zone; d.settings.consecHyphen = value.limit; return E.sel; }); LY.root.classList.toggle('hyphens', value.auto); } }, { label: '&Manual...', onClick: () => { ui.msg('Manual hyphenation: place the insertion point inside a long word and press Ctrl+- to insert an optional hyphen.', { icon: 'info' }); return false; } }, { label: 'Cancel' }] });
  };

  /* ================= Envelopes and Labels ================= */
  const LABELS = [['5160', 'Avery 5160 – Address', 3, 10, 189, 72, 13.5, 36, 9], ['5161', 'Avery 5161 – Address', 2, 10, 288, 72, 11.25, 36, 13.5], ['5162', 'Avery 5162 – Address', 2, 7, 288, 96, 11.25, 63, 13.5], ['5163', 'Avery 5163 – Shipping', 2, 5, 288, 144, 11.25, 36, 13.5], ['5164', 'Avery 5164 – Shipping', 2, 3, 288, 240, 11.25, 36, 13.5], ['5167', 'Avery 5167 – Return Address', 4, 20, 126, 36, 21.6, 36, 21.6], ['5395', 'Avery 5395 – Name Badge', 2, 4, 243, 168, 49.5, 41.4, 27], ['L7160', 'Avery L7160 (A4) – Address', 3, 7, 179.15, 107.7, 20.7, 43.9, 7.2]];
  G.envelopes = function (tab) {
    const direct = !!(window.VO && VO.reactDirectHost);
    const d = doc();
    let sel = E.sel && !E.collapsed() ? O.textRange(...E.range(), '\n') : '';
    const addr = h('textarea', { rows: 5, id: 'el-addr', style: 'width:100%' });
    addr.value = sel;
    const ret = h('textarea', { rows: 4, id: 'el-ret', style: 'width:100%' });
    ret.value = L.store.get('userAddress', '') || (L.store.get('userName', '') + '\n');
    const omit = ui.check('&Omit', false, null);
    const size = ui.select([['10', 'Size 10 (4 1/8 x 9 1/2 in)'], ['DL', 'DL (110 x 220 mm)'], ['C5', 'C5 (162 x 229 mm)'], ['Monarch', 'Monarch (3 7/8 x 7 1/2 in)']], '10', null, { id: 'el-size' });
    const laddr = h('textarea', { rows: 5, id: 'el-laddr', style: 'width:100%' });
    laddr.value = sel;
    const lprod = ui.select(LABELS.map((x) => [x[0], x[1]]), '5160', null, { id: 'el-prod' });
    const full = G.radios('el-full', [['full', '&Full page of the same label'], ['single', 'Single &label']], 'full');
    const row = G.num(1, { min: 1, max: 30, id: 'el-row' }), colN = G.num(1, { min: 1, max: 4, id: 'el-col' });
    let cur = tab === 'labels' ? 1 : 0;
    const reactRoot = h('div', { id: 'react-wordx-envelopes-root' });
    const reactEnvelopeState = { tab: cur, selectedText: sel, address: sel, returnAddress: ret.value, omit: false, size: '10', labelAddress: sel, product: '5160', printMode: 'full', row: 1, column: 1 };
    G.reactEnvelopeCatalog = () => ({ ...reactEnvelopeState, products: LABELS.map(([id, name]) => ({ id, name })) });
    G.reactEnvelopeTab = (value) => { cur = value === 'labels' ? 1 : 0; reactEnvelopeState.tab = cur; };
    G.reactEnvelopeAction = (action, value) => {
      cur = value.tab === 'labels' ? 1 : 0;
      Object.assign(reactEnvelopeState, value);
      if (action === 'add') { if (cur === 0) addEnvelope(value); else newLabels(value); }
      else if (action === 'print') { if (cur === 0) { addEnvelope(value); setTimeout(() => A().printPDF(), 100); } else newLabels(value); }
      return G.reactEnvelopeCatalog();
    };
    const tabs = ui.tabs([
      { label: '&Envelopes', body: h('div', { class: 'col' }, h('label', { for: 'el-addr', html: ui.ampHTML('&Delivery address:') }), addr, G.row(h('label', { for: 'el-ret', html: ui.ampHTML('&Return address:') }), omit), ret, G.f('Envelope si&ze:', size)), onShow: () => { cur = 0; } },
      { label: '&Labels', body: h('div', { class: 'col' }, h('label', { for: 'el-laddr', html: ui.ampHTML('&Address:') }), laddr, G.f('Label &product:', lprod), ui.group('Print', full, G.row(G.f('Ro&w:', row, 'narrow'), G.f('Colu&mn:', colN, 'narrow')))), onShow: () => { cur = 1; } },
    ], cur);
    ui.dialog({ title: 'Envelopes and Labels', width: 420, body: direct ? reactRoot : tabs, buttons: [{ label: '&Add to Document', primary: true, onClick: () => { if (direct) G.reactEnvelopeAction('add', envelopeValues()); else if (cur === 0) addEnvelope(); else newLabels(); } }, { label: cur === 0 ? '&Print (PDF)...' : 'New &Document', onClick: () => { if (direct) G.reactEnvelopeAction('print', envelopeValues()); else if (cur === 0) { addEnvelope(); setTimeout(() => A().printPDF(), 100); } else newLabels(); } }, { label: 'Cancel' }] });
    function envelopeValues() { const body = reactRoot.querySelector('#react-wordx-envelopes-body'); return { tab: reactEnvelopeState.tab === 1 ? 'labels' : 'envelopes', address: body?.querySelector('#react-wordx-envelope-address')?.value || '', returnAddress: body?.querySelector('#react-wordx-envelope-return')?.value || '', omit: !!body?.querySelector('#react-wordx-envelope-omit')?.checked, size: body?.querySelector('#react-wordx-envelope-size')?.value || '10', labelAddress: body?.querySelector('#react-wordx-label-address')?.value || '', product: body?.querySelector('#react-wordx-label-product')?.value || '5160', printMode: body?.querySelector('[name=react-wordx-label-mode]:checked')?.value || 'full', row: Number(body?.querySelector('#react-wordx-label-row')?.value) || 1, column: Number(body?.querySelector('#react-wordx-label-column')?.value) || 1 }; }
    function addEnvelope(values) {
      const returnAddress = values ? values.returnAddress : ret.value, omitReturn = values ? values.omit : omit.input.checked, delivery = values ? values.address : addr.value, envelopeSize = values ? values.size : size.value;
      L.store.set('userAddress', returnAddress);
      const sz = { 10: [684, 297], DL: [623.6, 311.8], C5: [649.15, 459.2], Monarch: [540, 279] }[envelopeSize];
      E.edit('Add Envelope', () => {
        const cont = D.touchList(d.main);
        const addressStyle = D.useStyle(d, 'EnvelopeAddress'), returnStyle = D.useStyle(d, 'EnvelopeReturn');
        const blocks = [];
        if (!omitReturn) for (const line of returnAddress.split('\n')) blocks.push(D.para(line ? [D.text(line)] : [], { style: returnStyle }));
        const lines = delivery.split('\n');
        lines.forEach((line, i) => blocks.push(D.para(line ? [D.text(line)] : [], { style: addressStyle, ind: { l: sz[0] * 0.42 }, sp: i === 0 ? { b: Math.max(0, sz[1] * 0.42 - (omitReturn ? 0 : 60)) } : undefined })));
        const last = blocks[blocks.length - 1];
        last.sect = Object.assign(D.defaultSect(), { pgW: sz[0], pgH: sz[1], orient: 'landscape', mt: 18, mb: 18, ml: 36, mr: 36, type: 'nextPage' });
        cont.blocks.unshift(...blocks);
        d._idxDirty = true;
        return D.pos(blocks[0], 0);
      });
    }
    function newLabels(values) {
      const spec = LABELS.find((x) => x[0] === (values ? values.product : lprod.value));
      const [, , nc, nr, lw, lh, ml, mt, gap] = spec;
      const text = values ? values.labelAddress : laddr.value;
      A().newDocument({ build: (nd) => {
        nd.sect.ml = ml; nd.sect.mr = 0; nd.sect.mt = mt; nd.sect.mb = 0; nd.sect.hdr = 0; nd.sect.ftr = 0;
        if (/A4/.test(spec[1])) { nd.sect.pgW = 595.3; nd.sect.pgH = 841.9; }
        const cols = [];
        for (let c = 0; c < nc; c++) { cols.push(lw); if (c < nc - 1 && gap) cols.push(gap); }
        const t = D.table([], cols, { style: 'TableNormal', w: { type: 'dxa', v: cols.reduce((a, b) => a + b, 0) }, layout: 'fixed', cellMar: { l: 5.4, r: 5.4, t: 0, b: 0 }, look: {} });
        for (let r = 0; r < nr; r++) {
          const cells = [];
          for (let c = 0; c < cols.length; c++) {
            const isGap = gap && c % 2 === 1;
            const lr = Math.floor(c / (gap ? 2 : 1));
            const fill = !isGap && ((values ? values.printMode : full.get()) === 'full' || (r === (values ? values.row : row.get()) - 1 && lr === (values ? values.column : colN.get()) - 1));
            const blocks = fill ? text.split('\n').map((ln) => D.para(ln ? [D.text(ln)] : [], { ind: { l: 5 }, sp: { b: 0, a: 0 } })) : [D.para()];
            cells.push(D.cell(blocks, { w: cols[c], vAlign: 'center' }));
          }
          t.rows.push(D.row(cells, { h: lh, hRule: 'exact', cantSplit: true }));
        }
        nd.main.blocks = [t, D.para()];
      } });
    }
  };

  /* ================= Letter Wizard ================= */
  G.letterWizard = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const reactRoot = h('div', { id: 'react-wordx-letter-wizard-root' });
    const date = ui.check('&Date line:', true, null);
    const datePic = ui.select(L.fields.DATE_PICTURES.slice(0, 6).map((p) => [p, L.fields.formatDate(new Date(), p)]), 'MMMM d, yyyy', null, { id: 'lw-date', translate: 'no' });
    const style = ui.select([['full', 'Full block'], ['modified', 'Modified block'], ['semi', 'Semi-block']], 'full', null, { id: 'lw-style' });
    const rname = h('input', { type: 'text', id: 'lw-rn', style: 'width:100%' });
    const raddr = h('textarea', { rows: 4, id: 'lw-ra', style: 'width:100%' });
    const sal = h('input', { type: 'text', id: 'lw-sal', value: 'Dear ', list: 'lw-sals', style: 'width:100%' });
    const sals = h('datalist', { id: 'lw-sals' }, ...['Dear ', 'Dear Sir or Madam:', 'Ladies and Gentlemen:', 'To Whom It May Concern:'].map((x) => h('option', { value: x })));
    const refLine = h('input', { type: 'text', id: 'lw-ref', style: 'width:100%', placeholder: 'In reply to:' });
    const subj = h('input', { type: 'text', id: 'lw-sub', style: 'width:100%' });
    const cc = h('input', { type: 'text', id: 'lw-cc', style: 'width:100%' });
    const sname = h('input', { type: 'text', id: 'lw-sn', value: L.store.get('userName', ''), style: 'width:100%' });
    const saddr = h('textarea', { rows: 3, id: 'lw-sa', style: 'width:100%' });
    saddr.value = L.store.get('userAddress', '');
    const closing = h('input', { type: 'text', id: 'lw-cl', value: 'Sincerely,', list: 'lw-cls', style: 'width:100%' });
    const cls = h('datalist', { id: 'lw-cls' }, ...L.autocorrect.AUTOTEXT.Closing.map((x) => h('option', { value: x })));
    const title = h('input', { type: 'text', id: 'lw-ti', style: 'width:100%' });
    const encl = G.num(0, { id: 'lw-enc', min: 0 });
    const tabs = ui.tabs([
      { label: 'Letter Format', body: h('div', { class: 'col' }, G.row(date, datePic), G.f('Page design / Letter style:', style, 'wide')) },
      { label: 'Recipient Info', body: h('div', { class: 'col' }, G.f('Recipient\'s name:', rname, 'wide'), h('label', { for: 'lw-ra', text: 'Delivery address:' }), raddr, sals, G.f('Salutation:', sal, 'wide')) },
      { label: 'Other Elements', body: h('div', { class: 'col' }, G.f('Reference line:', refLine, 'wide'), G.f('Subject:', subj, 'wide'), G.f('Courtesy copies (cc):', cc, 'wide')) },
      { label: 'Sender Info', body: h('div', { class: 'col' }, G.f('Sender\'s name:', sname, 'wide'), h('label', { for: 'lw-sa', text: 'Return address:' }), saddr, cls, G.f('Complimentary closing:', closing, 'wide'), G.f('Job title:', title, 'wide'), G.f('Enclosures:', encl, 'wide')) },
    ]);
    G.reactLetterWizardCatalog = () => ({ dates:L.fields.DATE_PICTURES.slice(0,6).map((value)=>({value,label:L.fields.formatDate(new Date(),value)})), styles:[['full','Full block'],['modified','Modified block'],['semi','Semi-block']], salutations:['Dear ','Dear Sir or Madam:','Ladies and Gentlemen:','To Whom It May Concern:'], closings:L.autocorrect.AUTOTEXT.Closing, name:L.store.get('userName',''), address:L.store.get('userAddress','') });
    G.reactLetterWizardAction = (v) => createLetter(v);
    function reactValues() { const body=reactRoot.querySelector('#react-wordx-letter-wizard-body'), read=(id)=>body?.querySelector('#'+id); return { date:!!read('react-wordx-letter-date')?.checked, datePic:read('react-wordx-letter-date-format')?.value||'MMMM d, yyyy', style:read('react-wordx-letter-style')?.value||'full', recipient:read('react-wordx-letter-recipient')?.value||'', delivery:read('react-wordx-letter-delivery')?.value||'', salutation:read('react-wordx-letter-salutation')?.value||'Dear ', reference:read('react-wordx-letter-reference')?.value||'', subject:read('react-wordx-letter-subject')?.value||'', cc:read('react-wordx-letter-cc')?.value||'', sender:read('react-wordx-letter-sender')?.value||'', returnAddress:read('react-wordx-letter-return')?.value||'', closing:read('react-wordx-letter-closing')?.value||'Sincerely,', title:read('react-wordx-letter-title')?.value||'', enclosures:Number(read('react-wordx-letter-enclosures')?.value)||0 }; }
    ui.dialog({ title: 'Letter Wizard', width: 460, body: direct ? reactRoot : tabs, buttons: [{ label: 'OK', primary: true, onClick: () => createLetter(direct ? reactValues() : null) }, { label: 'Cancel' }] });
    function createLetter(values) {
      const recipient=values?values.recipient:rname.value, delivery=values?values.delivery:raddr.value, salutation=values?values.salutation:sal.value, reference=values?values.reference:refLine.value, subject=values?values.subject:subj.value, copies=values?values.cc:cc.value, sender=values?values.sender:sname.value, returnAddress=values?values.returnAddress:saddr.value, close=values?values.closing:closing.value, jobTitle=values?values.title:title.value, enclosureCount=values?values.enclosures:encl.get(), dateEnabled=values?values.date:date.input.checked, dateFormat=values?values.datePic:datePic.value, letterStyle=values?values.style:style.value;
      L.store.set('userAddress', returnAddress);
      const indent = letterStyle === 'full' ? 0 : 252;
      const P = (t, o) => D.para(t ? [D.text(t)] : [], o || {});
      const bl = [];
      for (const line of returnAddress.split('\n').filter(Boolean)) bl.push(P(line, { ind: { l: indent } }));
      if (bl.length) bl.push(P(''));
      if (dateEnabled) bl.push(P(L.fields.formatDate(new Date(), dateFormat), { ind: { l: indent } }), P(''));
      if (reference) bl.push(P(reference), P(''));
      if (recipient) bl.push(P(recipient));
      for (const line of delivery.split('\n').filter(Boolean)) bl.push(P(line));
      bl.push(P(''));
      if (subject) bl.push(D.para([D.text('Subject: ', { b: true }), D.text(subject)]), P(''));
      bl.push(P(salutation.trim() === 'Dear' || salutation === 'Dear ' ? 'Dear ' + (recipient || '') + ':' : salutation), P(''));
      bl.push(P('Type your letter here.', { ind: { fl: letterStyle === 'semi' ? 36 : 0 } }), P(''));
      bl.push(P(close, { ind: { l: indent } }), P('', { ind: { l: indent } }), P('', { ind: { l: indent } }), P(sender, { ind: { l: indent } }));
      if (jobTitle) bl.push(P(jobTitle, { ind: { l: indent } }));
      if (enclosureCount) bl.push(P(''), P(`Enclosure${enclosureCount > 1 ? 's' : ''} (${enclosureCount})`));
      if (copies) bl.push(P(''), P('cc: ' + copies));
      E.pasteBlocks(bl.map((b) => (delete b.partial, b)));
    }
  };

  /* ================= Macros, Templates & Organizer ================= */
  G.macros = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const lb = G.listBox([], null, { height: 140 });
    const root = h('div', { id: 'react-wordx-macros-root' });
    ui.dialog({ title: 'Macros', width: 360, body: direct ? root : h('div', { class: 'col' }, h('label', { text: 'Macro name:' }), h('input', { type: 'text', style: 'width:100%', disabled: true }), lb, h('div', { class: 'tp-note', text: 'Macros are not run in Wordx. VBA projects in .docm files are kept when the document is saved as .docm by Word, but Wordx saves the document text only.' })), buttons: [{ label: 'Close', primary: true }] });
  };
  G.templates = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const d = doc();
    const auto = ui.check('Automatically &update document styles', !!d.settings.attachedTemplateUpdate, null);
    const name = h('input', { type: 'text', value: d.props.template || 'Normal', style: 'width:100%', readonly: true });
    const reactRoot = h('div', { id: 'react-wordx-templates-root' });
    G.reactTemplatesCatalog = () => ({ name: d.props.template || 'Normal', update: !!d.settings.attachedTemplateUpdate });
    G.reactTemplatesAction = async (action, value) => {
      if (action === 'attach') {
        const f = (await L.pickFiles('.dotx,.dotm,.docx'))[0];
        if (!f) return G.reactTemplatesCatalog();
        const res = await L.docx.read(await L.readAsArrayBuffer(f));
        d.props.template = f.name;
        if (value) A().mergeStylesFrom(res.doc, true);
        A().status(L.t('Template attached: {0}', f.name));
      } else if (action === 'organizer') G.organizer();
      else if (action === 'save') d.settings.attachedTemplateUpdate = value || undefined;
      return G.reactTemplatesCatalog();
    };
    ui.dialog({ title: 'Templates and Add-ins', width: 400, body: direct ? reactRoot : h('div', { class: 'col' }, ui.group('Document template', name, auto, h('div', { class: 'tp-btns' }, ui.button('&Attach...', async () => { const f = (await L.pickFiles('.dotx,.dotm,.docx'))[0]; if (!f) return; const res = await L.docx.read(await L.readAsArrayBuffer(f)); d.props.template = f.name; name.value = f.name; if (auto.input.checked) A().mergeStylesFrom(res.doc, true); A().status(L.t('Template attached: {0}', f.name)); }, { class: 'btn small' }), ui.button('&Organizer...', () => G.organizer(), { class: 'btn small' }))), h('div', { class: 'tp-note', text: 'Attaching a template with "Automatically update document styles" copies its styles into this document.' })), buttons: [{ label: 'OK', primary: true, onClick: () => { d.settings.attachedTemplateUpdate = direct ? !!reactRoot.querySelector('#react-wordx-template-update')?.checked : auto.input.checked; } }, { label: 'Cancel' }] });
  };
  G.organizer = function () {
    if (window.VO && VO.reactDirectHost) {
      const d = doc();
      let other = null, otherName = '';
      const catalog = () => ({ mineName: L.t('In {0}:', A().fileName), otherName: L.t('In {0}:', otherName || '(no file)'), mine: Object.keys(d.styles).sort().map((id) => ({ id, name: D.styleDisplayName(d.styles[id]), builtin: !!d.styles[id].builtin })), theirs: other ? Object.keys(other.styles).sort().map((id) => ({ id, name: D.styleDisplayName(other.styles[id]) })) : [] });
      G.reactOrganizerCatalog = catalog;
      G.reactOrganizerAction = async (action, value) => {
        const ids = Array.isArray(value?.ids) ? value.ids : [];
        if (action === 'open') {
          const f = (await L.pickFiles('.docx,.dotx,.docm,.dotm'))[0];
          if (f) { other = (await L.docx.read(await L.readAsArrayBuffer(f))).doc; otherName = f.name; }
        } else if (action === 'copy' && other) {
          for (const id of ids) if (other.styles[id]) E.edit('Copy Style', () => { D.touchKey(d, 'styles'); d.styles[id] = L.clone(other.styles[id]); D.stylesChanged(); return E.sel; });
        } else if (action === 'delete') {
          E.edit('Delete Styles', () => { D.touchKey(d, 'styles'); for (const id of ids) if (d.styles[id] && !d.styles[id].builtin && id !== 'Normal') delete d.styles[id]; D.stylesChanged(); return E.sel; });
        } else if (action === 'rename' && ids[0] && d.styles[ids[0]] && value?.name?.trim()) {
          E.edit('Rename Style', () => { D.touchKey(d, 'styles'); d.styles[ids[0]].name = value.name.trim(); return E.sel; });
        }
        return catalog();
      };
      const root = h('div', { id: 'react-wordx-organizer-root' });
      ui.dialog({ title: 'Organizer', width: 560, body: root, buttons: [{ label: 'Close', primary: true }] });
      return;
    }
    const d = doc();
    let other = null, otherName = '';
    const mine = h('select', { size: 12, multiple: true, style: 'width:100%', translate: 'no' });
    const theirs = h('select', { size: 12, multiple: true, style: 'width:100%', translate: 'no' });
    const fillSel = (sel, dd) => { L.clear(sel); if (!dd) return; for (const k of Object.keys(dd.styles).sort()) sel.appendChild(h('option', { value: k }, document.createTextNode(D.styleDisplayName(dd.styles[k])))); };
    fillSel(mine, d);
    const lblOther = h('div', { class: 'cd-cap', text: 'In (no file):' });
    ui.dialog({ title: 'Organizer', width: 560, body: G.row(h('div', { class: 'col', style: 'flex:1' }, h('div', { class: 'cd-cap', text: L.t('In {0}:', A().fileName) }), mine), h('div', { class: 'col', style: 'width:110px;justify-content:center' }, ui.button('<- &Copy', () => { if (!other) return; for (const o of theirs.selectedOptions) { E.edit('Copy Style', () => { D.touchKey(d, 'styles'); d.styles[o.value] = L.clone(other.styles[o.value]); D.stylesChanged(); return E.sel; }); } fillSel(mine, d); }), ui.button('&Delete', () => { E.edit('Delete Styles', () => { D.touchKey(d, 'styles'); for (const o of mine.selectedOptions) if (!d.styles[o.value].builtin && o.value !== 'Normal') delete d.styles[o.value]; D.stylesChanged(); return E.sel; }); fillSel(mine, d); }), ui.button('&Rename...', async () => { const o = mine.selectedOptions[0]; if (!o) return; const n = await ui.prompt('New name:', D.styleDisplayName(d.styles[o.value]), 'Rename'); if (!n) return; E.edit('Rename Style', () => { D.touchKey(d, 'styles'); d.styles[o.value].name = n; return E.sel; }); fillSel(mine, d); })), h('div', { class: 'col', style: 'flex:1' }, lblOther, theirs, ui.button('Open &File...', async () => { const f = (await L.pickFiles('.docx,.dotx,.docm,.dotm'))[0]; if (!f) return; other = (await L.docx.read(await L.readAsArrayBuffer(f))).doc; otherName = f.name; lblOther.textContent = L.t('In {0}:', otherName); fillSel(theirs, other); }))), buttons: [{ label: 'Close', primary: true }] });
  };

  /* ================= AutoCorrect ================= */
  G.autocorrect = function (tab) {
    const direct = !!(window.VO && VO.reactDirectHost);
    const o = A().opts;
    const ac = L.autocorrect;
    const chk = (k, l, def) => ui.check(l, o[k] !== undefined ? o[k] !== false : def !== false, (v) => { o[k] = v; A().saveOpts(); });
    const from = h('input', { type: 'text', id: 'ac-from', style: 'width:110px' }), to = h('input', { type: 'text', id: 'ac-to', style: 'flex:1;min-width:0' });
    const list = G.listBox([], null, { height: 150, translate: 'no', onChange: (v) => { from.value = v; to.value = ac.list().get(v) || ''; } });
    const fill = () => { const m = ac.list(); list.setItems(Array.from(m.keys()).sort((a, b) => a.localeCompare(b)).map((k) => ({ value: k, label: `${k}  →  ${m.get(k)}` }))); };
    fill();
    from.addEventListener('input', () => { const m = ac.list(); if (m.has(from.value)) to.value = m.get(from.value); });
    const mainRoot = h('div', { id: 'react-wordx-autocorrect-main-root' });
    const mainKeys = ['acButtons','twoCaps','capSentence','capCells','capDays','capsLock','autocorrect','autoComplete'];
    let selectedEntry = '';
    G.reactAutoCorrectMainCatalog = () => ({ options: Object.fromEntries(mainKeys.map((key) => [key, o[key] !== false])), entries: Array.from(ac.list().entries()).sort((a,b)=>a[0].localeCompare(b[0])), selected: selectedEntry });
    G.reactAutoCorrectMainAction = (action, value) => {
      if (action === 'option' && mainKeys.includes(value?.key)) { o[value.key] = !!value.enabled; A().saveOpts(); }
      else if (action === 'select') selectedEntry = value?.from || '';
      else if (action === 'entry' && value?.from) { ac.setEntry(value.from, value.to); selectedEntry = value.from; }
      else if (action === 'delete' && value?.from) { ac.setEntry(value.from, null); selectedEntry = ''; }
      else if (action === 'exceptions') exc();
      return G.reactAutoCorrectMainCatalog();
    };
    const exc = () => {
      const direct = !!(window.VO && VO.reactDirectHost);
      const ta = h('textarea', { rows: 10, style: 'width:100%' }); ta.value = ac.exceptions().join('\n');
      const root = h('div', { id: 'react-wordx-autocorrect-exceptions-root' });
      G.reactAutoCorrectExceptionsCatalog = () => ({ words: ac.exceptions().join('\n') });
      G.reactAutoCorrectExceptionsAction = (value) => { ac.setExceptions(String(value || '').split(/\s+/).filter(Boolean)); return G.reactAutoCorrectExceptionsCatalog(); };
      ui.dialog({ title: 'AutoCorrect Exceptions', width: 300, body: direct ? root : h('div', { class: 'col' }, h('div', { text: "Don't capitalize after:" }), ta), buttons: [{ label: 'OK', primary: true, onClick: () => { if (direct) G.reactAutoCorrectExceptionsAction(root.querySelector('textarea')?.value || ''); else ac.setExceptions(ta.value.split(/\s+/).filter(Boolean)); } }, { label: 'Cancel' }] });
    };
    const atName = h('input', { type: 'text', id: 'at-name', style: 'width:100%' });
    const atList = G.listBox(L.store.get('autoText', []).map((x) => x.name), null, { height: 120, translate: 'no', onChange: (v) => { const e = L.store.get('autoText', []).find((x) => x.name === v); atName.value = v; atPv.textContent = e ? e.text : ''; } });
    const atPv = h('div', { class: 'tp-box', style: 'min-height:40px;white-space:pre-wrap', translate: 'no' });
    const autoTextRoot = h('div', { id: 'react-wordx-autotext-root' });
    let selectedAutoText = '';
    G.reactAutoTextCatalog = () => ({ entries: L.store.get('autoText', []).map((entry) => ({ name: entry.name, text: entry.text })), selected: selectedAutoText, name: atName.value, preview: L.store.get('autoText', []).find((entry) => entry.name === selectedAutoText)?.text || '', autoComplete: o.autoComplete !== false });
    G.reactAutoTextAction = (action, value) => {
      if (action === 'select') { selectedAutoText = value || ''; atName.value = selectedAutoText; }
      else if (action === 'add') {
        atName.value = value || atName.value;
        const text = E.sel && !E.collapsed() ? O.textRange(...E.range(), '\n') : '';
        if (!text) ui.msg('Select the text you want to store as AutoText first.', { icon: 'info' });
        else { const name = atName.value.trim() || text.slice(0, 32); ac.addAutoText(name, text); selectedAutoText = name; atName.value = name; }
      } else if (action === 'delete' && (value || selectedAutoText)) { selectedAutoText = value || selectedAutoText; ac.deleteAutoText(selectedAutoText); selectedAutoText = ''; atName.value = ''; atPv.textContent = ''; }
      else if (action === 'insert') { const entry = L.store.get('autoText', []).find((item) => item.name === (value || selectedAutoText)); if (entry) E.insertTextAt(entry.text, L.N('AutoText')); }
      return G.reactAutoTextCatalog();
    };
    const formatRoot = h('div', { id: 'react-wordx-autocorrect-format-root' });
    const formatKeys = ['smartQuotes','ordinals','fractions','dashes','autoFormat','autoUrl','autoBullets','autoBorders','autoTables','autoNumbers','autoHeadings'];
    const formatOptions = () => Object.fromEntries(formatKeys.map((key) => [key, key === 'autoHeadings' ? !!o[key] : o[key] !== false]));
    G.reactAutoCorrectFormatCatalog = () => ({ options: formatOptions() });
    G.reactAutoCorrectFormatAction = (key, enabled) => { if (formatKeys.includes(key)) { o[key] = !!enabled; A().saveOpts(); } return G.reactAutoCorrectFormatCatalog(); };
    const tabs = ui.tabs([
      { label: '&AutoCorrect', body: direct ? h('div', { class: 'col' }, mainRoot) : h('div', { class: 'col' }, chk('acButtons', 'Show Auto&Correct Options buttons'), chk('twoCaps', 'Correct TWo INitial &CApitals'), chk('capSentence', 'Capitalize first letter of &sentences'), chk('capCells', 'Capitalize first letter of table &cells'), chk('capDays', 'Capitalize &names of days'), chk('capsLock', 'Correct accidental usage of cAPS &LOCK key'), ui.button('&Exceptions...', exc, { class: 'btn small' }), chk('autocorrect', 'Replace &text as you type'), G.row(G.f('&Replace:', from), G.f('&With:', to)), list, h('div', { class: 'tp-btns' }, ui.button('&Add', () => { if (!from.value) return; ac.setEntry(from.value, to.value); fill(); }, { class: 'btn small' }), ui.button('&Delete', () => { if (!from.value) return; ac.setEntry(from.value, null); fill(); from.value = ''; to.value = ''; }, { class: 'btn small' }))) },
      { label: 'AutoFormat As You T&ype', body: direct ? h('div', { class: 'col' }, formatRoot) : h('div', { class: 'col' }, ui.group('Replace as you type', chk('smartQuotes', L.N('"&Straight quotes" with “smart quotes”')), chk('ordinals', '&Ordinals (1st) with superscript'), chk('fractions', '&Fractions (1/2) with fraction character (½)'), chk('dashes', '&Hyphens (--) with dash (—)'), chk('autoFormat', L.N('*&Bold* and _italic_ with real formatting')), chk('autoUrl', '&Internet and network paths with hyperlinks')), ui.group('Apply as you type', chk('autoBullets', 'Automatic b&ulleted lists'), chk('autoBorders', 'Border &lines'), chk('autoTables', '&Tables'), chk('autoNumbers', 'Automatic &numbered lists'), chk('autoHeadings', 'Built-in Heading st&yles', false))) },
      { label: 'Auto&Text', body: direct ? h('div', { class: 'col' }, autoTextRoot) : h('div', { class: 'col' }, chk('autoComplete', 'Show &AutoComplete suggestions'), h('label', { for: 'at-name', text: 'Enter AutoText entries here:' }), atName, atList, h('div', { class: 'cd-cap', text: 'Preview' }), atPv, h('div', { class: 'tp-btns' }, ui.button('&Add', () => { const t = E.sel && !E.collapsed() ? O.textRange(...E.range(), '\n') : ''; if (!t) { ui.msg('Select the text you want to store as AutoText first.', { icon: 'info' }); return; } const n = atName.value.trim() || t.slice(0, 32); ac.addAutoText(n, t); atList.setItems(L.store.get('autoText', []).map((x) => x.name)); }, { class: 'btn small' }), ui.button('&Delete', () => { ac.deleteAutoText(atName.value); atList.setItems(L.store.get('autoText', []).map((x) => x.name)); atPv.textContent = ''; }, { class: 'btn small' }), ui.button('&Insert', () => { const e = L.store.get('autoText', []).find((x) => x.name === atName.value); if (e) E.insertTextAt(e.text, L.N('AutoText')); }, { class: 'btn small' }))) },
    ], tab === 'autotext' ? 2 : tab === 'format' ? 1 : 0);
    ui.dialog({ title: 'AutoCorrect: English (U.S.)', width: 460, body: tabs, buttons: [{ label: 'OK', primary: true }, { label: 'Cancel' }] });
  };

  /* ================= Customize ================= */
  G.customize = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    if (direct) {
      const root = h('div', { id: 'react-wordx-customize-root' });
      const toolbarIds = ['standard', 'formatting', 'drawing', 'tables', 'reviewing', 'picture', 'outlining', 'headerFooter', 'mailMerge', 'wordart'];
      const catalog = () => ({
        toolbars: toolbarIds.map((id) => ({ id, label: ui.cmds['tb_' + id]?.label || id, shown: A().tbShown(id) })),
        commands: Object.values(ui.cmds).filter((command) => command.label && !command.id.startsWith('tb_')).sort((a, b) => ui.stripAmp(a.menuLabel || a.label).localeCompare(ui.stripAmp(b.menuLabel || b.label))).map((command) => ({ id: command.id, label: ui.stripAmp(command.menuLabel || command.label), key: command.key ? ui.keyText(command.key) : '' })),
        shortcuts: Object.entries(A().SHORTCUTS).sort((a, b) => a[1].localeCompare(b[1])).flatMap(([key, id]) => { const command = ui.cmds[id]; return command ? [{ key: ui.keyText(key), label: ui.stripAmp(command.menuLabel || command.label) }] : []; }),
        screenTips: A().opts.screenTips !== false,
      });
      G.reactCustomizeCatalog = catalog;
      G.reactCustomizeAction = (action, value) => {
        if (action === 'toolbar' && value?.id) A().toggleToolbar(value.id);
        else if (action === 'screenTips') { A().opts.screenTips = !!value?.enabled; A().saveOpts(); L.domHost().classList.toggle('notips', !value?.enabled); }
        else if (action === 'command' && value?.id) ui.exec(value.id);
        return catalog();
      };
      ui.dialog({ title: 'Customize', width: 460, body: root, buttons: [{ label: 'Close', primary: true }] });
      return;
    }
    const tbs = ['standard', 'formatting', 'drawing', 'tables', 'reviewing', 'picture', 'outlining', 'headerFooter', 'mailMerge', 'wordart'].map((k) => ui.check(ui.cmds['tb_' + k].label, A().tbShown(k), () => A().toggleToolbar(k)));
    const cmds = Object.values(ui.cmds).filter((c) => c.label && !c.id.startsWith('tb_')).sort((a, b) => ui.stripAmp(a.label).localeCompare(ui.stripAmp(b.label)));
    /* the commands' names as shown (stripAmp translates them) and their keys */
    const cmdList = G.listBox(cmds.map((c) => ({ value: c.id, label: ui.stripAmp(c.menuLabel || c.label) + (c.key ? '   ' + ui.keyText(c.key) : '') })), null, { height: 220, translate: 'no' });
    const keys = h('table', { class: 'kbd-table', translate: 'no' });
    for (const [k, id] of Object.entries(A().SHORTCUTS).sort((a, b) => a[1].localeCompare(b[1]))) { const c = ui.cmds[id]; if (c) keys.appendChild(h('tr', null, h('td', null, document.createTextNode(ui.keyText(k))), h('td', null, document.createTextNode(ui.stripAmp(c.menuLabel || c.label))))); }
    const o = A().opts;
    ui.dialog({ title: 'Customize', width: 460, body: ui.tabs([
      { label: 'Tool&bars', body: h('div', { class: 'col' }, ...tbs) },
      { label: '&Commands', body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'All Wordx commands. Double-click one to run it.' }), cmdList), onShow: () => {} },
      { label: '&Options', body: h('div', { class: 'col' }, ui.check('Show Standard and Formatting toolbars on two rows', true, null), ui.check('Always show full menus', true, null), ui.check('Show Scree&nTips on toolbars', o.screenTips !== false, (v) => { o.screenTips = v; A().saveOpts(); L.domHost().classList.toggle('notips', !v); }), ui.check('Show shortcut &keys in ScreenTips', true, null)) },
      { label: '&Keyboard', body: h('div', { style: 'max-height:300px;overflow:auto' }, keys) },
    ]), buttons: [{ label: 'Close', primary: true }] });
    cmdList.querySelector('.listbox').addEventListener('dblclick', () => { const id = cmdList.get(); if (id) ui.exec(id); });
  };

  /* ================= Options ================= */
  G.options = function (tab) {
    const o = A().opts;
    const d = doc();
    const pwInp = h('input', { type: 'password', id: 'op-pw', value: d.openPassword || '', style: 'width:220px', autocomplete: 'new-password' });
    const chk = (label, get, set) => ui.check(label, get(), set);
    const opt = (k, l, def) => chk(l, () => (o[k] !== undefined ? !!o[k] : !!def), (v) => { o[k] = v; });
    const lyo = (k, l) => chk(l, () => !!LY.opts[k], (v) => { LY.opts[k] = v; rerender = true; });
    let rerender = false;
    const fieldShade = ui.select([['never', 'Never'], ['always', 'Always'], ['selected', 'When selected']], o.fieldShading || 'selected', (v) => { o.fieldShading = v; }, { id: 'op-fs' });
    const units = ui.select([['in', 'Inches'], ['cm', 'Centimeters'], ['mm', 'Millimeters'], ['pt', 'Points'], ['pi', 'Picas']], o.unit || 'in', (v) => { o.unit = v; }, { id: 'op-unit' });
    const uname = h('input', { type: 'text', id: 'op-un', value: o.userName || '', style: 'width:100%' });
    const uinit = h('input', { type: 'text', id: 'op-ui', value: o.userInitials || '', style: 'width:80px' });
    const uaddr = h('textarea', { rows: 3, id: 'op-ua', style: 'width:100%' });
    uaddr.value = L.store.get('userAddress', '');
    const insColor = ui.select([['author', 'By author'], ['red', 'Red'], ['blue', 'Blue'], ['green', 'Green']], o.insColor || 'author', (v) => { o.insColor = v; }, { id: 'op-ic' });
    const balloons = ui.select([['always', 'Always'], ['never', 'Never'], ['comments', 'Only for comments/formatting']], LY.opts.balloons === false ? 'never' : 'comments', (v) => { LY.opts.balloons = v !== 'never'; rerender = true; }, { id: 'op-bal' });
    const compat = ui.select([['11', 'Microsoft Word 2003'], ['12', 'Microsoft Word 2007'], ['14', 'Microsoft Word 2010'], ['15', 'Microsoft Word 2013 and later']], String(d.settings.compat || 11), (v) => { d.settings.compat = +v; d.dirty = true; rerender = true; }, { id: 'op-compat' });
    const autoRec = G.num(o.autoSaveMin || 10, { min: 1, max: 120, id: 'op-asr' });
    const look = ui.lookField();
    const direct = !!(window.VO && VO.reactDirectHost);
    const viewRoot = direct ? h('div', { id: 'react-wordx-options-view-root' }) : null;
    const viewBody = direct ? h('div', { class: 'col' }, viewRoot) : null;
    const generalRoot = direct ? h('div', { id: 'react-wordx-options-general-root' }) : null;
    const generalBody = direct ? h('div', { class: 'col' }, generalRoot) : null;
    const editRoot = direct ? h('div', { id: 'react-wordx-options-edit-root' }) : null;
    const editBody = direct ? h('div', { class: 'col' }, editRoot) : null;
    const printRoot = direct ? h('div', { id: 'react-wordx-options-print-root' }) : null;
    const printBody = direct ? h('div', { class: 'col' }, printRoot) : null;
    const saveRoot = direct ? h('div', { id: 'react-wordx-options-save-root' }) : null;
    const saveBody = direct ? h('div', { class: 'col' }, saveRoot) : null;
    const trackRoot = direct ? h('div', { id: 'react-wordx-options-track-root' }) : null;
    const trackBody = direct ? h('div', { class: 'col' }, trackRoot) : null;
    const userRoot = direct ? h('div', { id: 'react-wordx-options-user-root' }) : null;
    const userBody = direct ? h('div', { class: 'col' }, userRoot) : null;
    const compatRoot = direct ? h('div', { id: 'react-wordx-options-compat-root' }) : null;
    const compatBody = direct ? h('div', { class: 'col' }, compatRoot) : null;
    const securityRoot = direct ? h('div', { id: 'react-wordx-options-security-root' }) : null;
    const securityBody = direct ? h('div', { class: 'col' }, securityRoot) : null;
    const spellingRoot = direct ? h('div', { id: 'react-wordx-options-spelling-root' }) : null;
    const spellingBody = direct ? h('div', { class: 'col' }, spellingRoot) : null;
    const suitePrefs = window.VO;
    let generalLook = suitePrefs && suitePrefs.look && suitePrefs.look.get();
    let generalLanguage = suitePrefs && suitePrefs.lang && suitePrefs.lang.get();
    const viewOptions = [
      ['startupPane', 'Startup Task Pane', true], ['paneFollows', 'Task pane follows the selected object', true], ['highlightShow', L.N('Highlight'), true],
      ['showBookmarks', L.N('Bookmarks'), true, true], ['statusBar', 'Status bar', true], ['screenTips', L.N('ScreenTips'), true],
      ['fieldCodes', 'Field codes', true, true], ['smartTags', 'Smart tags', false], ['animText', 'Animated text', false],
      ['showMarks', L.N('All'), true, true], ['showHidden', 'Hidden text', false, true], ['textBoundaries', 'Text boundaries', false, true],
      ['vruler', 'Vertical ruler (Print view only)', true], ['whiteSpace', 'White space between pages (Print view only)', true],
    ];
    G.reactOptionsViewCatalog = () => ({
      options: viewOptions.map(([key, label, fallback, layout]) => ({ key, label, checked: layout ? !!LY.opts[key] : (o[key] !== undefined ? !!o[key] : !!fallback), layout: !!layout })),
      fieldShading: o.fieldShading || 'selected',
    });
    G.reactOptionsViewSet = (key, value) => {
      if (key === 'fieldShading') o.fieldShading = value;
      else {
        const item = viewOptions.find(([id]) => id === key);
        if (!item) return;
        if (item[3]) { LY.opts[key] = !!value; rerender = true; }
        else o[key] = !!value;
      }
      return G.reactOptionsViewCatalog();
    };
    G.reactOptionsGeneralCatalog = () => ({
      options: ['bgRepag', 'blueBg', 'sound'].map((key) => ({ key, checked: o[key] === undefined ? key === 'bgRepag' : !!o[key] })),
      unit: o.unit || 'in', units: [['in','Inches'],['cm','Centimeters'],['mm','Millimeters'],['pt','Points'],['pi','Picas']],
      look: generalLook || 'classic', looks: suitePrefs && suitePrefs.look ? suitePrefs.look.list().map(({ id, name }) => ({ id, name })) : [],
      language: generalLanguage || 'auto', languages: suitePrefs && suitePrefs.lang ? suitePrefs.lang.list().map(({ id, name }) => ({ id, name: id === 'auto' ? L.t('Auto (as the browser)') : name })) : [],
    });
    G.reactOptionsGeneralSet = (key, value) => {
      if (key === 'look') generalLook = value;
      else if (key === 'language') generalLanguage = value;
      else if (key === 'unit') o.unit = value;
      else if (['bgRepag','blueBg','sound'].includes(key)) o[key] = !!value;
      return G.reactOptionsGeneralCatalog();
    };
    const editOptions = [['typingReplaces','Typing replaces selection',true],['dragDrop','Drag-and-drop text editing',true],['ctrlClick','Use CTRL + Click to follow hyperlink',true],['smartPara','Use smart paragraph selection',true],['overtype','Overtype mode',true],['smartCut','Smart cut and paste',true],['pasteOptions','Show Paste Options buttons',true]];
    G.reactOptionsEditCatalog = () => editOptions.map(([key,label,fallback]) => ({ key, label, checked: key === 'overtype' ? !!E.overtype : (o[key] !== undefined ? !!o[key] : !!fallback) }));
    G.reactOptionsEditSet = (key, value) => {
      if (key === 'overtype') E.overtype = !!value;
      else if (editOptions.some(([id]) => id === key)) o[key] = !!value;
      return G.reactOptionsEditCatalog();
    };
    const printOptions = [['updateFieldsOnPrint','Update fields',false],['printBg','Background colors and images',true],['printProps','Document properties',false],['printCodes','Field codes',false],['printHidden','Hidden text',false],['printDrawings','Drawing objects',true]];
    G.reactOptionsPrintCatalog = () => ({ printing: printOptions.slice(0,2).map(([key,label,fallback])=>({key,label,checked:o[key]===undefined?fallback:!!o[key]})), include: printOptions.slice(2).map(([key,label,fallback])=>({key,label,checked:o[key]===undefined?fallback:!!o[key]})) });
    G.reactOptionsPrintSet = (key, value) => { if(printOptions.some(([id])=>id===key))o[key]=!!value;return G.reactOptionsPrintCatalog(); };
    G.reactOptionsSaveCatalog = () => ({ promptProps:o.promptProps!==undefined?!!o.promptProps:false,autoRecover:o.autoRecover!==false,autoSaveMin:o.autoSaveMin||10,formats:[['docx','Word Document (*.docx)'],['dotx','Word Template (*.dotx)']] });
    G.reactOptionsSaveSet = (key, value) => { if(key==='promptProps'||key==='autoRecover')o[key]=!!value;return G.reactOptionsSaveCatalog(); };
    G.reactOptionsTrackCatalog = () => ({ color:o.insColor||'author',colors:[['author','By author'],['red','Red'],['blue','Blue'],['green','Green']],balloons:LY.opts.balloons===false?'never':'comments',balloonOptions:[['always','Always'],['never','Never'],['comments','Only for comments/formatting']] });
    G.reactOptionsTrackSet = (key, value) => { if(key==='color')o.insColor=value;else if(key==='balloons'){LY.opts.balloons=value!=='never';rerender=true;}return G.reactOptionsTrackCatalog(); };
    G.reactOptionsUserCatalog = () => ({ name:o.userName||'',initials:o.userInitials||'',address:L.store.get('userAddress','') });
    G.reactOptionsCompatCatalog = () => ({ version:String(d.settings.compat||11),versions:[['11','Microsoft Word 2003'],['12','Microsoft Word 2007'],['14','Microsoft Word 2010'],['15','Microsoft Word 2013 and later']] });
    G.reactOptionsCompatSet = (value) => { d.settings.compat=+value;d.dirty=true;rerender=true;return G.reactOptionsCompatCatalog(); };
    G.reactOptionsSecurityCatalog = () => ({ removePersonal:o.removePersonal===true,warnTracked:o.warnTracked===true,password:d.openPassword||'' });
    G.reactOptionsSecuritySet = (key, value) => { if(key==='removePersonal'||key==='warnTracked')o[key]=!!value;return G.reactOptionsSecurityCatalog(); };
    const spellOptions = [['spell','Check spelling as you type',true],['ignoreUpper','Ignore words in UPPERCASE',true],['ignoreNum','Ignore words with numbers',true],['ignoreUrl','Ignore Internet and file addresses',true],['grammarAsType','Check grammar as you type',true],['grammar','Check grammar with spelling',true]];
    G.reactOptionsSpellingCatalog = () => ({ spelling:spellOptions.slice(0,4).map(([key,label,fallback])=>({key,label,checked:o[key]===undefined?fallback:!!o[key]})),grammar:spellOptions.slice(4).map(([key,label,fallback])=>({key,label,checked:o[key]===undefined?fallback:!!o[key]})),hidden:!!(L.spell&&L.spell.hidden) });
    G.reactOptionsSpellingSet = (key,value) => { if(key==='hidden'){if(L.spell)L.spell.hidden=!!value;}else if(spellOptions.some(([id])=>id===key))o[key]=!!value;return G.reactOptionsSpellingCatalog(); };
    G.reactOptionsSpellingAction = (action) => {
      if(action==='dictionary'){
        if(window.VO&&VO.reactDirectHost){
          const root=h('div',{id:'react-wordx-custom-dictionary-root'});
          G.reactCustomDictionaryCatalog=()=>({words:L.store.get('customDict',[]).join('\n')});
          G.reactCustomDictionaryAction=(words)=>{L.store.set('customDict',String(words).split(/\s+/).filter(Boolean).map((word)=>word.toLowerCase()));if(L.spell)L.spell.reloadCustom();return G.reactCustomDictionaryCatalog();};
          ui.dialog({title:'Custom Dictionary',width:340,body:root,buttons:[{label:'OK',primary:true,onClick:()=>G.reactCustomDictionaryAction(root.querySelector('textarea')?.value||'')},{label:'Cancel'}]});
        }else{const ta=h('textarea',{rows:12,style:'width:100%'});ta.value=L.store.get('customDict',[]).join('\n');ui.dialog({title:'CUSTOM.DIC',width:300,body:ta,buttons:[{label:'OK',primary:true,onClick:()=>{L.store.set('customDict',ta.value.split(/\s+/).filter(Boolean).map((x)=>x.toLowerCase()));if(L.spell)L.spell.reloadCustom();}},{label:'Cancel'}]});}
      }
      else if(action==='readability'){const r=L.spell.readability();ui.msg(L.t('Words: {0}\nCharacters: {1}\nParagraphs: {2}\nSentences: {3}\n\nSentences per Paragraph: {4}\nWords per Sentence: {5}\nCharacters per Word: {6}\n\nPassive Sentences: {7}%\nFlesch Reading Ease: {8}\nFlesch-Kincaid Grade Level: {9}',r.words,r.chars,r.paras,r.sentences,r.sPerPara,r.wPerS,r.cPerW,r.passive,r.ease,r.grade),{icon:'info',title:'Readability Statistics'});}
      else if(action==='recheck'&&L.spell){L.spell.ignoredOnce.clear();L.spell.refreshSoon(0);}
    };
    const tabs = ui.tabs([
      { label: 'View', body: direct ? viewBody : h('div', { class: 'col' }, ui.group('Show', h('div', { class: 'grid2' }, opt('startupPane', 'Startup Task Pane', true), opt('paneFollows', 'Task pane follows the selected object', true), opt('highlightShow', L.N('Highlight'), true), lyo('showBookmarks', L.N('Bookmarks')), opt('statusBar', 'Status bar', true), opt('screenTips', L.N('ScreenTips'), true), lyo('fieldCodes', 'Field codes'), opt('smartTags', 'Smart tags', false), opt('animText', 'Animated text', false)), G.f('Field shading:', fieldShade)), ui.group('Formatting marks', h('div', { class: 'grid2' }, lyo('showMarks', L.N('All')), lyo('showHidden', 'Hidden text'))), ui.group('Print and Web Layout options', h('div', { class: 'grid2' }, lyo('textBoundaries', 'Text boundaries'), opt('vruler', 'Vertical ruler (Print view only)', true), opt('whiteSpace', 'White space between pages (Print view only)', true)))) },
      { label: 'General', body: direct ? generalBody : h('div', { class: 'col' }, look && look.el, opt('bgRepag', 'Background repagination', true), opt('blueBg', 'Blue background, white text', false), opt('sound', 'Provide feedback with sound', false), G.f('Measurement units:', units, 'wide')) },
      { label: 'Edit', body: direct ? editBody : h('div', { class: 'col' }, ui.group('Editing options', opt('typingReplaces', 'Typing replaces selection', true), opt('dragDrop', 'Drag-and-drop text editing', true), opt('ctrlClick', 'Use CTRL + Click to follow hyperlink', true), opt('smartPara', 'Use smart paragraph selection', true), chk('Overtype mode', () => E.overtype, (v) => { E.overtype = v; }), opt('smartCut', 'Smart cut and paste', true)), ui.group('Cut and paste options', opt('pasteOptions', 'Show Paste Options buttons', true))) },
      { label: 'Print', body: direct ? printBody : h('div', { class: 'col' }, ui.group('Printing options', opt('updateFieldsOnPrint', 'Update fields', false), opt('printBg', 'Background colors and images', true)), ui.group('Include with document', opt('printProps', 'Document properties', false), opt('printCodes', 'Field codes', false), opt('printHidden', 'Hidden text', false), opt('printDrawings', 'Drawing objects', true))) },
      { label: 'Save', body: direct ? saveBody : h('div', { class: 'col' }, opt('promptProps', 'Prompt for document properties', false), G.row(opt('autoRecover', 'Save AutoRecover info every:', true), autoRec, h('span', { text: 'minutes' })), h('div', { class: 'tp-note', text: 'AutoRecover keeps a copy in this browser only; it is offered when Wordx opens after an unexpected close.' }), G.f('Save Word files as:', ui.select([['docx', 'Word Document (*.docx)'], ['dotx', 'Word Template (*.dotx)']], 'docx', null, { id: 'op-fmt' }), 'wide')) },
      { label: 'Security', body: direct ? securityBody : h('div', { class: 'col' }, ui.group('Privacy options', opt('removePersonal', 'Remove personal information from file properties on save', false), opt('warnTracked', 'Warn before printing, saving or sending a file that contains tracked changes or comments', false)), ui.group('File encryption options for this document', h('div', { class: 'col' }, h('label', { for: 'op-pw', text: 'Password to open:' }), pwInp, h('div', { class: 'tp-note', text: 'Saved .docx files are encrypted with AES-256, as Word 2010 and later do. Passwords are case-sensitive; if you forget it, the document cannot be recovered.' }))), h('div', { class: 'tp-note', text: 'To restrict editing instead, use Tools ▸ Protect Document.' })) },
      { label: 'Spelling & Grammar', body: direct ? spellingBody : h('div', { class: 'col' }, ui.group('Spelling', opt('spell', 'Check spelling as you type', true), chk('Hide spelling and grammar errors in this document', () => !!(L.spell && L.spell.hidden), (v) => { if (L.spell) L.spell.hidden = v; }), opt('ignoreUpper', 'Ignore words in UPPERCASE', true), opt('ignoreNum', 'Ignore words with numbers', true), opt('ignoreUrl', 'Ignore Internet and file addresses', true), ui.button('Custom Dictionaries...', () => G.reactOptionsSpellingAction('dictionary'), { class: 'btn small' })), ui.group('Grammar', opt('grammarAsType', 'Check grammar as you type', true), opt('grammar', 'Check grammar with spelling', true), ui.button('Readability Statistics...', () => G.reactOptionsSpellingAction('readability'), { class: 'btn small' })), ui.button('Recheck Document', () => G.reactOptionsSpellingAction('recheck'), { class: 'btn small' })) },
      { label: 'Track Changes', body: direct ? trackBody : h('div', { class: 'col' }, ui.group('Markup', G.f('Insertions color:', insColor, 'wide')), ui.group('Balloons', G.f('Use Balloons (Print and Web Layout):', balloons, 'wide'))) },
      { label: 'User Information', body: direct ? userBody : h('div', { class: 'col' }, G.f('Name:', uname, 'wide'), G.f('Initials:', uinit, 'wide'), h('label', { for: 'op-ua', text: 'Mailing address:' }), uaddr) },
      { label: 'Compatibility', body: direct ? compatBody : h('div', { class: 'col' }, G.f('Lay out this document as if created in:', compat, 'wide'), h('div', { class: 'tp-note', text: 'Word 2003 layout positions tables by their cell margins and keeps legacy spacing rules; later versions align tables with the text.' })) },
    ], { view: 0, general: 1, edit: 2, print: 3, save: 4, security: 5, spelling: 6, track: 7, user: 8, compat: 9 }[tab] || 0);
    tabs.classList.add('options-tabs');
    ui.dialog({ title: 'Options', width: 560, body: tabs, buttons: [{ label: 'OK', primary: true, onClick: () => {
      const newPw = direct ? securityRoot.querySelector('#react-wordx-options-password')?.value || '' : pwInp.value;
      if (newPw !== (d.openPassword || '')) {
        if (!newPw) { d.openPassword = null; A().status(L.N('The password to open was removed; the next save is not encrypted.')); }
        else {
          /* Word asks for the password a second time before using it */
          const confirmRoot = direct ? h('div', { id: 'react-wordx-confirm-password-root' }) : null;
          const again = direct ? null : h('input', { type: 'password', id: 'op-pw2', style: 'width:100%' });
          let val = null;
          const confirm = ui.dialog({ title: 'Confirm Password', width: 340, body: direct ? confirmRoot : h('div', { class: 'col' }, h('label', { for: 'op-pw2', text: 'Reenter password to open:' }), again, h('div', { class: 'tp-note', text: 'Caution: If you lose or forget the password, it cannot be recovered.' })), buttons: [{ label: 'OK', primary: true, onClick: () => { val = direct ? confirmRoot.querySelector('#react-wordx-confirm-password')?.value || '' : again.value; } }, { label: 'Cancel' }] });
          confirm.done.then(() => {
            if (val == null) return;
            if (val !== newPw) { ui.msg('The password confirmation does not match.', { icon: 'warn' }); return; }
            d.openPassword = newPw;
            A().dirty && A().dirty();
            A().status('The document will be encrypted with this password when you save it.');
          });
          if (!direct) setTimeout(() => again.focus(), 0);
        }
      }
      o.userName = ((direct ? userRoot.querySelector('#react-wordx-options-user-name')?.value : uname.value) || '').trim() || 'Wordx User';
      o.userInitials = ((direct ? userRoot.querySelector('#react-wordx-options-user-initials')?.value : uinit.value) || '').trim();
      o.autoSaveMin = direct ? (+saveRoot.querySelector('#react-wordx-options-autorecover-min')?.value || 10) : autoRec.get();
      L.store.set('userAddress', direct ? userRoot.querySelector('#react-wordx-options-user-address')?.value || '' : uaddr.value);
      A().saveOpts();
      A().applyOpts();
      if (direct) {
        if (suitePrefs && suitePrefs.look && generalLook !== suitePrefs.look.get()) suitePrefs.look.set(generalLook);
        if (suitePrefs && suitePrefs.lang && generalLanguage !== suitePrefs.lang.get()) { suitePrefs.lang.set(generalLanguage); ui.toast(L.t('The new language is used in the windows you open from now on.')); }
      } else if (look) look.apply();
      if (rerender) { LY.root.classList.toggle('marks', !!LY.opts.showMarks); LY.render(); E.restoreDom(true); }
    } }, { label: 'Cancel' }] });
  };

  /* ================= Properties ================= */
  G.properties = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const d = doc();
    const p = d.props;
    const st = D.stats(D.allParas(d));
    const fld = (k, l) => { const inp = h('input', { type: 'text', id: 'pr-' + k, value: p[k] || '', style: 'width:100%' }); return { k, inp, el: G.f(l, inp, 'wide') }; };
    const fields = [fld('title', '&Title:'), fld('subject', '&Subject:'), fld('creator', '&Author:'), fld('manager', '&Manager:'), fld('company', 'C&ompany:'), fld('category', 'Cat&egory:'), fld('keywords', '&Keywords:')];
    const comments = h('textarea', { rows: 3, id: 'pr-desc', style: 'width:100%' });
    comments.value = p.description || '';
    const hb = fld('hyperlinkBase', '&Hyperlink base:');
    const fmt = (s) => (s ? new Date(s).toLocaleString() : '');
    const statT = h('table', { class: 'wc-table' });
    for (const [k, v] of [['Pages:', LY.pages.length || 1], ['Paragraphs:', st.paras], ['Lines:', A().docStats().lines], ['Words:', st.words], ['Characters:', st.charsNoSp], ['Characters (with spaces):', st.chars]]) statT.appendChild(h('tr', null, h('td', { text: k }), h('td', { text: Number(v).toLocaleString() })));
    const heads = h('div', { class: 'tp-box', style: 'max-height:200px;overflow:auto' });
    D.reindex(d);
    heads.setAttribute('translate', 'no');   // the document's own headings
    for (const q of d.main.paras) { const lv = D.headingLevel(d, q); if (lv) heads.appendChild(h('div', { style: `padding-left:${(lv - 1) * 12}px` }, document.createTextNode(D.plainText(q).trim()))); }
    if (!heads.childNodes.length) heads.textContent = L.t('(no headings)');
    /* custom */
    const cName = h('input', { type: 'text', id: 'pr-cn', list: 'pr-cnl', style: 'width:100%' });
    const cnl = h('datalist', { id: 'pr-cnl' }, ...['Checked by', 'Client', 'Date completed', 'Department', 'Destination', 'Disposition', 'Division', 'Document number', 'Editor', 'Forward to', 'Group', 'Language', 'Mailstop', 'Matter', 'Office', 'Owner', 'Project', 'Publisher', 'Purpose', 'Received from', 'Recorded by', 'Recorded date', 'Reference', 'Source', 'Status', 'Telephone number', 'Typist'].map((x) => h('option', { value: x })));
    const cType = ui.select([['text', 'Text'], ['date', 'Date'], ['number', 'Number'], ['bool', 'Yes or no']], 'text', null, { id: 'pr-ct' });
    const cVal = h('input', { type: 'text', id: 'pr-cv', style: 'width:100%' });
    const custom = Object.assign({}, d.custom || {});
    const cList = G.listBox([], null, { height: 110, translate: 'no', onChange: (v) => { const c = custom[v]; cName.value = v; cVal.value = c && c.v != null ? c.v : c; cType.value = (c && c.type) || 'text'; } });
    const fillC = () => cList.setItems(Object.keys(custom).map((k) => ({ value: k, label: `${k}: ${custom[k] && custom[k].v != null ? custom[k].v : custom[k]}` })));
    fillC();
    const propertiesCatalog = () => ({ fileName: A().fileName, general: [['Type:', L.t('Microsoft Word Document')], ['Created:', fmt(p.created)], ['Modified:', fmt(p.modified)], ['Last printed:', fmt(p.printed)]], summary: fields.map((f) => ({ key: f.k, label: f.el.querySelector('label')?.textContent.replace(/:$/, '') || f.k, value: f.inp.value })), comments: comments.value, hyperlinkBase: hb.inp.value, template: p.template || 'Normal', statistics: [['Created:', fmt(p.created)], ['Modified:', fmt(p.modified)], ['Last saved by:', p.lastModifiedBy || ''], ['Revision number:', String(p.revision || 1)], ['Total editing time:', L.t('{0} Minutes', p.totalTime || 0)], ['Pages:', LY.pages.length || 1], ['Paragraphs:', st.paras], ['Lines:', A().docStats().lines], ['Words:', st.words], ['Characters:', st.charsNoSp], ['Characters (with spaces):', st.chars]], headings: d.main.paras.flatMap((para) => { const level = D.headingLevel(d, para); return level ? [{ level, text: D.plainText(para).trim() }] : []; }), custom: Object.entries(custom).map(([name, item]) => ({ name, type: item && item.type || 'text', value: item && item.v != null ? item.v : item })) });
    G.reactPropertiesCatalog = propertiesCatalog;
    G.reactPropertiesAction = (action, value) => {
      if (action === 'addCustom' && value && value.name.trim()) { custom[value.name.trim()] = { type: value.type, v: value.type === 'number' ? +value.value : value.type === 'bool' ? /^(y|yes|true|1)$/i.test(value.value) : value.value }; fillC(); }
      else if (action === 'deleteCustom' && value) { delete custom[value]; fillC(); }
      else if (action === 'save') {
        E.edit('Properties', () => { D.touchKey(d, 'props'); D.touchKey(d, 'custom'); for (const field of value.fields) d.props[field.key] = field.value; d.props.description = value.comments; d.props.hyperlinkBase = value.hyperlinkBase || undefined; d.custom = custom; return E.sel; });
        A().updateTitle();
      }
      return propertiesCatalog();
    };
    const reactRoot = h('div', { id: 'react-wordx-properties-root' });
    /* a label and its value as shown (translated where it is set, else the document's own) */
    const prop = ([a, b]) => h('tr', null, h('td', { text: a }), h('td', { translate: 'no' }, document.createTextNode(b)));
    const tabs = ui.tabs([
      { label: 'General', body: h('div', { class: 'col' }, h('div', { class: 'row', style: 'align-items:center' }, h('span', { html: L.icons.get('new') }), h('b', { translate: 'no' }, document.createTextNode(A().fileName + '.docx'))), h('table', { class: 'wc-table' }, ...[['Type:', L.t('Microsoft Word Document')], ['Created:', fmt(p.created)], ['Modified:', fmt(p.modified)], ['Last printed:', fmt(p.printed)]].map(prop))) },
      { label: 'Summary', body: h('div', { class: 'col' }, ...fields.map((f) => f.el), h('label', { for: 'pr-desc', html: ui.ampHTML('Co&mments:') }), comments, hb.el, h('div', { class: 'tp-note', text: L.t('Template: {0}', p.template || 'Normal') })) },
      { label: 'Statistics', body: h('div', { class: 'col' }, h('table', { class: 'wc-table' }, ...[['Created:', fmt(p.created)], ['Modified:', fmt(p.modified)], ['Last saved by:', p.lastModifiedBy || ''], ['Revision number:', String(p.revision || 1)], ['Total editing time:', L.t('{0} Minutes', p.totalTime || 0)]].map(prop)), h('div', { class: 'cd-cap', text: 'Statistics:' }), statT) },
      { label: 'Contents', body: h('div', { class: 'col' }, h('div', { text: 'Document contents:' }), heads) },
      { label: 'Custom', body: h('div', { class: 'col' }, cnl, G.f('&Name:', cName, 'wide'), G.f('&Type:', cType, 'wide'), G.f('&Value:', cVal, 'wide'), h('div', { class: 'tp-btns' }, ui.button('&Add', () => { if (!cName.value.trim()) return; custom[cName.value.trim()] = { type: cType.value, v: cType.value === 'number' ? +cVal.value : cType.value === 'bool' ? /^(y|yes|true|1)$/i.test(cVal.value) : cVal.value }; fillC(); }, { class: 'btn small' }), ui.button('&Delete', () => { delete custom[cName.value.trim()]; fillC(); }, { class: 'btn small' })), h('label', { text: 'Properties:' }), cList) },
    ], 1);
    ui.dialog({ title: L.t('{0} Properties', A().fileName), width: 420, body: direct ? reactRoot : tabs, buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (direct) { G.reactPropertiesAction('save', { fields: propertiesCatalog().summary.map(({ key }) => ({ key, value: reactRoot.querySelector('#react-wordx-property-' + key)?.value || '' })), comments: reactRoot.querySelector('#react-wordx-properties-comments')?.value || '', hyperlinkBase: reactRoot.querySelector('#react-wordx-properties-hyperlink-base')?.value || '' }); return; }
      E.edit('Properties', () => { D.touchKey(d, 'props'); D.touchKey(d, 'custom'); for (const f of fields) d.props[f.k] = f.inp.value; d.props.description = comments.value; d.props.hyperlinkBase = hb.inp.value || undefined; d.custom = custom; return E.sel; });
      A().updateTitle();
    } }, { label: 'Cancel' }] });
  };

  /* ================= Zoom ================= */
  G.zoom = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const cur = Math.round(LY.zoom * 100);
    const r = G.radios('zm', [['200', '&200%'], ['100', '&100%'], ['75', '&75%'], ['pw', '&Page width'], ['tw', '&Text width'], ['wp', '&Whole page'], ['many', '&Many pages']], [200, 100, 75].includes(cur) ? String(cur) : 'pct', (v) => { const n = { 200: 200, 100: 100, 75: 75 }[v]; if (n) pct.set(n); });
    const pct = ui.spin({ value: cur, min: 10, max: 500, step: 10, dec: 0, unit: '%', id: 'zm-pct' });
    const pv = h('div', { class: 'preview-box', style: 'width:150px;height:110px', translate: 'no' }, document.createTextNode('AaBbCcDdEeXxYyZz'));
    const reactRoot = h('div', { id: 'react-wordx-zoom-root' });
    G.reactZoomCatalog = () => ({ current: cur, preset: [200, 100, 75].includes(cur) ? String(cur) : 'pct' });
    ui.dialog({ title: 'Zoom', width: 360, body: direct ? reactRoot : G.row(ui.group('Zoom to', r), h('div', { class: 'col' }, G.f('Percen&t:', pct), h('div', { class: 'cd-cap', text: 'Preview' }), pv)), buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (direct) { const v = reactRoot.querySelector('input[name=react-wordx-zoom]:checked')?.value || 'pct'; const value = reactRoot.querySelector('#react-wordx-zoom-percent')?.value || String(cur); if (v === 'pw') A().zoomTo('Page Width'); else if (v === 'tw') A().zoomTo('Text Width'); else if (v === 'wp') A().zoomTo('Whole Page'); else if (v === 'many') A().zoomTo('Two Pages'); else A().zoomTo((v === 'pct' ? value : v) + '%'); return; }
      const v = r.get(); if (v === 'pw') A().zoomTo('Page Width'); else if (v === 'tw') A().zoomTo('Text Width'); else if (v === 'wp') A().zoomTo('Whole Page'); else if (v === 'many') A().zoomTo('Two Pages'); else A().zoomTo(pct.get() + '%');
    } }, { label: 'Cancel' }] });
  };

  /* ================= Print (to PDF) ================= */
  G.print = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const n = LY.pages.length || 1;
    const range = G.radios('pr-r', [['all', '&All'], ['current', 'Curr&ent page'], ['selection', '&Selection'], ['pages', 'Pa&ges:']], 'all');
    const pagesIn = h('input', { type: 'text', id: 'pr-pages', style: 'width:150px' });
    pagesIn.addEventListener('focus', () => range.set('pages'));
    const what = ui.select([['doc', 'Document'], ['markup', 'Document showing markup'], ['props', 'Document properties'], ['list', 'List of markup'], ['styles', 'Styles'], ['autotext', 'AutoText entries'], ['keys', 'Key assignments']], 'doc', null, { id: 'pr-what' });
    const which = ui.select([['all', 'All pages in range'], ['odd', 'Odd pages'], ['even', 'Even pages']], 'all', null, { id: 'pr-which' });
    const perSheet = ui.select([[1, '1 page'], [2, '2 pages'], [4, '4 pages'], [6, '6 pages'], [8, '8 pages'], [16, '16 pages']], 1, null, { id: 'pr-pps' });
    const copies = G.num(1, { min: 1, max: 99, id: 'pr-copies' });
    const root = h('div', { id: 'react-wordx-print-root' });
    ui.dialog({
      title: 'Print', width: 470,
      body: direct ? root : h('div', { class: 'col' }, ui.group('Printer', G.f('&Name:', ui.select([['pdf', 'Save as PDF (Wordx)']], 'pdf', null, { id: 'pr-printer' }), 'wide'), h('div', { class: 'tp-note', text: 'Status: Ready — the pages are saved as a PDF that you can print from any viewer.' })), G.row(ui.group('Page range', range, pagesIn, h('div', { class: 'tp-note', text: 'Enter page numbers and/or page ranges separated by commas. For example, 1,3,5–12' })), ui.group('Copies', G.f('Number of &copies:', copies))), G.row(G.f('Print &what:', what, 'wide'), G.f('P&rint:', which)), ui.group('Zoom', G.f('Pages per s&heet:', perSheet, 'wide'))),
      buttons: [{ label: '&Options...', onClick: () => { G.options('print'); return false; } }, { label: 'OK', primary: true, onClick: () => { doPrint(); } }, { label: 'Cancel' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
    async function doPrint() {
      const w = direct ? root.querySelector('#react-wordx-print-what')?.value || 'doc' : what.value;
      if (w !== 'doc' && w !== 'markup') { printList(w); return; }
      let pages = null;
      const rv = direct ? root.querySelector('input[name=react-wordx-print-range]:checked')?.value || 'all' : range.get();
      if (rv === 'current') pages = [Math.max(0, LY.pageOf(E.sel.f)) + 1];
      else if (rv === 'selection' && E.sel && !E.collapsed()) { const [a, b] = E.range(); const p0 = LY.pageOf(a), p1 = LY.pageOf(b); pages = []; for (let i = p0; i <= p1; i++) pages.push(i + 1); }
      else if (rv === 'pages') { pages = []; for (const part of (direct ? root.querySelector('#react-wordx-print-pages')?.value || '' : pagesIn.value).split(',')) { const m = /^\s*(\d+)\s*(?:[-–]\s*(\d+))?\s*$/.exec(part); if (m) { const a = +m[1], b = m[2] ? +m[2] : a; for (let i = a; i <= b; i++) if (i >= 1 && i <= n) pages.push(i); } } }
      const whichValue = direct ? root.querySelector('#react-wordx-print-which')?.value || 'all' : which.value;
      if (whichValue !== 'all') pages = (pages || Array.from({ length: n }, (_, i) => i + 1)).filter((i) => (whichValue === 'odd' ? i % 2 === 1 : i % 2 === 0));
      const prevMarkup = LY.opts.markup;
      if (w === 'doc' && prevMarkup !== 'final') { LY.opts.markup = 'final'; LY.render(); }
      try {
        const blob = await A().buildPDF(pages);
        L.ui.busy(false);
        doc().props.printed = new Date().toISOString();
        await L.saveFile(`${A().fileName}.pdf`, blob);
      } catch (e) { L.ui.busy(false); ui.msg(L.t('The PDF could not be created: {0}', e.message || e), { icon: 'error' }); }
      finally { if (LY.opts.markup !== prevMarkup) { LY.opts.markup = prevMarkup; LY.render(); E.restoreDom(true); } }
    }
    /* the printed lists in the UI's language: style kinds and revision kinds as shown */
    const KINDS = { paragraph: L.N('Paragraph'), character: L.N('Character'), table: L.N('Table'), numbering: L.N('List'), ins: L.N('Inserted'), del: L.N('Deleted'), chg: L.N('Formatted'), pchg: L.N('Formatted'), mins: L.N('Inserted'), mdel: L.N('Deleted') };
    function printList(kind) {
      const d = doc();
      const lines = [];
      if (kind === 'props') { for (const [k, v] of Object.entries(d.props)) if (v) lines.push(`${k}: ${v}`); }
      else if (kind === 'styles') { for (const k of Object.keys(d.styles).sort()) { const s = d.styles[k], ty = L.t(KINDS[s.type] || s.type); lines.push(s.basedOn ? L.t('{0} ({1}) — based on {2}', D.styleDisplayName(s), ty, s.basedOn) : `${D.styleDisplayName(s)} (${ty})`); } }
      else if (kind === 'autotext') { for (const e of L.autocorrect.autoTextEntries()) lines.push(`${e.cat}: ${e.name}`); }
      else if (kind === 'keys') { for (const [k, id] of Object.entries(A().SHORTCUTS)) if (ui.cmds[id]) lines.push(`${ui.keyText(k)}\t${ui.stripAmp(ui.cmds[id].label)}`); }
      else if (kind === 'list') { for (const r of L.review.list(d.main)) lines.push(`${L.t(KINDS[r.type] || r.type)} — ${r.stamp.author || ''}: ${r.text}`); for (const k in d.comments) lines.push(L.t('Comment {0}: {1}', d.comments[k].author, d.comments[k].blocks.map((b) => (b.t === 'p' ? D.plainText(b) : '')).join(' '))); }
      const nd = D.newDoc();
      const whatLabel = direct ? root.querySelector('#react-wordx-print-what')?.selectedOptions[0]?.textContent || kind : what.selectedOptions[0].text;
      nd.main.blocks = [D.para([D.text(A().fileName + ' — ' + whatLabel, { b: true, sz: 14 })])].concat(lines.map((l) => D.para([D.text(l)])));
      const keep = D.doc;
      D.doc = nd;
      LY.render();
      A().buildPDF().then(async (blob) => { L.ui.busy(false); D.doc = keep; LY.render(); E.restoreDom(true); await L.saveFile(`${A().fileName} (${kind}).pdf`, blob); }).catch(() => { D.doc = keep; LY.render(); L.ui.busy(false); });
    }
  };

  /* ================= Save As ================= */
  G.saveAs = function (cb) {
    const direct = !!(window.VO && VO.reactDirectHost);
    const name = h('input', { type: 'text', id: 'sa-name', value: A().fileName, style: 'width:100%' });
    const initialType = ['docm', 'dotx', 'dotm', 'md', 'mdzip'].includes(A().fileType) ? A().fileType : 'docx';
    const type = ui.select(L.file.types.map((t) => [t.id, t.label]), initialType, null, { id: 'sa-type' });
    const note = h('div', { class: 'tp-note', text: 'Your browser asks where to save the file, or saves it to the Downloads folder.' });
    const root = h('div', { id: 'react-wordx-save-as-root' });
    G.reactSaveAsCatalog = () => ({ name: A().fileName, type: initialType, types: L.file.types.map((item) => ({ id: item.id, label: item.label })), note: L.t('Your browser asks where to save the file, or saves it to the Downloads folder.') });
    G.reactSaveAsAction = (fileName, fileType) => { const normalized = String(fileName || '').trim().replace(/\.(docx|docm|dotx|dotm|htm|html|rtf|txt|md|zip|pdf)$/i, '') || 'Document'; cb(normalized, fileType); };
    return ui.dialog({ title: 'Save As', width: 430, focus: direct ? root : name, body: direct ? root : h('div', { class: 'col' }, G.f('File &name:', name, 'wide'), G.f('Save as &type:', type, 'wide'), note), buttons: [{ label: '&Save', primary: true, onClick: () => { if (direct) { G.reactSaveAsAction(root.querySelector('#react-wordx-save-as-name')?.value || '', root.querySelector('#react-wordx-save-as-type')?.value || initialType); return; } const n = name.value.trim().replace(/\.(docx|docm|dotx|dotm|htm|html|rtf|txt|md|zip|pdf)$/i, '') || 'Document'; cb(n, type.value); } }, { label: 'Cancel' }] });
  };
  G.pasteSpecial = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const lb = G.listBox([{ value: 'rich', label: 'Formatted Text (RTF)' }, { value: 'html', label: 'HTML Format' }, { value: 'text', label: 'Unformatted Text' }, { value: 'utext', label: 'Unformatted Unicode Text' }], 'rich', { height: 110 });
    const clip = E.clip;
    const root = h('div', { id: 'react-wordx-paste-special-root' });
    G.reactPasteSpecialCatalog = () => ({ hasClipboard: !!clip });
    ui.dialog({ title: 'Paste Special', width: 380, body: direct ? root : h('div', { class: 'col' }, h('div', { text: clip ? 'Source: Wordx document' : 'Source: (the Wordx clipboard is empty — copy something first, or press Ctrl+V)' }), G.radios('ps-m', [['paste', '&Paste:'], ['link', 'Paste &link:']], 'paste'), lb, h('div', { class: 'tp-note', text: 'Formatted text keeps fonts and formatting; unformatted text takes the formatting of the destination.' })), buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (!clip) return;
      const v = direct ? root.querySelector('#react-wordx-paste-special-format')?.value || 'rich' : lb.get();
      if (v === 'text' || v === 'utext') E.pasteText(clip.text);
      else E.pasteBlocks(clip.blocks, { keepFormat: v === 'rich' || v === 'html' });
    } }, { label: 'Cancel' }] });
  };
  G.about = () => ui.about({ name: 'Wordx', text: [
    L.N('A word processor that runs entirely in your browser. Documents are read and written as Office Open XML (.docx, .docm), RTF, HTML, text and Markdown; nothing is uploaded.'),
    L.N('Spelling, AutoCorrect and the thesaurus work offline, in English (U.S.) and English (U.K.).')] });
  G.newFromTemplate = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const cats = L.templates.categories();
    let pick = 'blank';
    const create = (id) => { if (id === 'blank') A().newDocument(); else if (id === 'web') { A().newDocument().then(() => A().setView('web')); } else if (id === 'mail') A().newDocument({ build: (d) => { d.main.blocks = [D.para([D.text('To: ')]), D.para([D.text('Subject: ')]), D.para()]; } }); else L.templates.create(id); };
    const reactRoot = h('div', { id: 'react-wordx-template-picker-root' });
    G.reactTemplatePickerCatalog = () => ({ categories: cats.map((category) => ({ name: category, items: category === 'General' ? [{ id: 'blank', name: 'Blank Document' }, { id: 'web', name: 'Web Page' }, { id: 'mail', name: 'E-mail Message' }] : L.templates.list().filter((template) => template.cat === category).map(({ id, name }) => ({ id, name })) })) });
    let okButton = null;
    reactRoot.dataset.selected = 'blank';
    G.reactTemplatePickerAction = (action, id) => { if (action === 'select' && id) reactRoot.dataset.selected = id; else if (action === 'accept' && id) { reactRoot.dataset.selected = id; okButton?.click(); } };
    if (direct) {
      const dlg = ui.dialog({ title: 'Templates', width: 520, body: reactRoot, buttons: [{ label: 'OK', primary: true, onClick: () => create(reactRoot.dataset.selected || 'blank') }, { label: 'Cancel' }] });
      okButton = dlg.buttons[0];
      return;
    }
    const pv = h('div', { class: 'preview-box', style: 'width:150px;height:190px' });
    const drawPv = (id) => { L.clear(pv); const page = h('div', { class: 'ps-page', style: 'width:110px;height:140px' }); for (let i = 0; i < (id === 'blank' ? 0 : 14); i++) page.appendChild(h('i', { style: `top:${14 + i * 8}px;${i % 5 === 4 ? 'right:40px' : ''}` })); pv.appendChild(page); };
    const tabs = ui.tabs(cats.map((c) => {
      const grid = h('div', { class: 'preset-grid' });
      const items = c === 'General' ? [{ id: 'blank', name: 'Blank Document' }, { id: 'web', name: 'Web Page' }, { id: 'mail', name: 'E-mail Message' }] : L.templates.list().filter((t) => t.cat === c);
      for (const t of items) { const b = h('button', { type: 'button', class: t.id === pick ? 'on' : '' }, h('span', { html: L.icons.get('new', 32) || L.icons.get('new') }), h('span', { text: t.name })); b.addEventListener('click', () => { pick = t.id; L.$$('.preset-grid button', tabs).forEach((x) => x.classList.toggle('on', x === b)); drawPv(t.id); }); b.addEventListener('dblclick', () => ok.click()); grid.appendChild(b); }
      return { label: c, body: grid };
    }));
    drawPv('blank');
    const dlg = ui.dialog({ title: 'Templates', width: 520, body: G.row(h('div', { style: 'flex:1;min-width:240px' }, tabs), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Preview' }), pv)), buttons: [{ label: 'OK', primary: true, onClick: () => create(pick) }, { label: 'Cancel' }] });
    const ok = dlg.buttons[0];
  };
})();
