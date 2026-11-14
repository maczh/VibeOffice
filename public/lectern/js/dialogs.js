/* PPTX — dialog boxes. */
(function () {
  'use strict';
  const L = window.L;
  const { h } = L;
  const ui = L.ui;
  const D = (L.dlg = {});
  const E = () => L.ed;
  let reactSaveAs = null;
  let reactSaveAsCallback = null;
  let reactSaveAsResolve = null;
  let reactSaveAsReject = null;
  let reactZoom = null;
  let reactGrid = null;
  let reactDateTime = null;
  let reactSymbol = null;
  let reactSymbolCallback = null;
  let reactChangeCase = null;
  let reactLineSpacing = null;
  let reactHeaderFooter = null;
  let reactSetupShow = null;
  let reactReplaceFonts = null;
  let reactInsertTable = null;
  let reactInsertTableCallback = null;
  let reactFont = null;
  let reactBullets = null;
  let reactActionSettings = null;
  let reactActionShapes = null;
  let reactPrint = null;
  let reactProperties = null;
  let reactAutoCorrect = null;
  let reactFind = null;
  let reactPageSetup = null;
  let reactHyperlink = null;
  let reactHyperlinkApply = null;
  let reactWordartText = null;
  let reactWordartTextCallback = null;
  let reactWordartGallery = null;
  let reactWordartGalleryCallback = null;
  let reactPhotoAlbum = false;
  let reactEffectOptions = null;
  let reactEffectOptionsCallback = null;
  let reactSlideFinder = null;
  let reactOptions = null;
  let reactOptionsConfirm = null;
  let reactBackground = null;
  let reactFillEffects = null;
  let reactFillEffectsResolve = null;
  let reactFormatShape = null;
  D.reactSaveAsCatalog = () => reactSaveAs;
  D.reactSaveAsAction = async (action, values) => {
    if (!reactSaveAs) return false;
    if (action === 'close') {
      reactSaveAs = null;
      reactSaveAsCallback = null;
      L.bus.emit('pptx-save-as-dialog');
      if (reactSaveAsResolve) reactSaveAsResolve('declined');
      reactSaveAsResolve = null;
      reactSaveAsReject = null;
      return 'declined';
    }
    if (action !== 'save' || !values || typeof values.name !== 'string' || !values.name.trim() || !reactSaveAs.types.some(([id]) => id === values.type)) return false;
    const name = values.name.trim().replace(/\.(pptx|pptm|ppsx|ppsm|potx|potm)$/i, '') || reactSaveAs.name;
    const type = values.type;
    const save = reactSaveAsCallback;
    reactSaveAs = null;
    reactSaveAsCallback = null;
    L.bus.emit('pptx-save-as-dialog');
    try {
      const result = save ? await save(name, type) : false;
      if (reactSaveAsResolve) reactSaveAsResolve(result);
      reactSaveAsResolve = null;
      reactSaveAsReject = null;
      return result;
    } catch (error) {
      if (reactSaveAsReject) reactSaveAsReject(error);
      reactSaveAsResolve = null;
      reactSaveAsReject = null;
      throw error;
    }
  };
  D.reactZoomCatalog = () => reactZoom;
  D.reactZoomApply = (value) => {
    if (!reactZoom || (value !== 'fit' && (!Number.isFinite(value) || value < 10 || value > 400))) return false;
    E().setZoom(value === 'fit' ? 'fit' : value / 100);
    reactZoom = null;
    L.bus.emit('pptx-zoom-dialog');
    return true;
  };
  D.reactZoomClose = () => { if (!reactZoom) return; reactZoom = null; L.bus.emit('pptx-zoom-dialog'); };
  D.reactGridCatalog = () => reactGrid;
  D.reactGridApply = (value) => {
    if (!reactGrid || !value || typeof value.snap !== 'boolean' || typeof value.show !== 'boolean' || typeof value.guides !== 'boolean' || ![4.5, 6, 7.2, 9, 12, 14.4, 18, 24, 36, 72].includes(value.size)) return false;
    Object.assign(E().grid, value);
    E().layout();
    L.ui.refresh();
    L.app.saveOpts();
    reactGrid = null;
    L.bus.emit('pptx-grid-dialog');
    return true;
  };
  D.reactGridClose = () => { if (!reactGrid) return; reactGrid = null; L.bus.emit('pptx-grid-dialog'); };
  D.reactDateTimeCatalog = () => reactDateTime;
  D.reactDateTimeApply = (format, automatic) => {
    if (!reactDateTime || !reactDateTime.formats.some((item) => item.id === format) || typeof automatic !== 'boolean') return false;
    if (!L.te.active()) L.app.ensureTextTarget();
    if (!L.te.active()) return false;
    L.hist.push('Insert Date and Time');
    if (automatic) L.te.insertField(format); else L.te.insertText(L.fmtDate(new Date(), format));
    reactDateTime = null;
    L.bus.emit('pptx-date-time-dialog');
    return true;
  };
  D.reactDateTimeClose = () => { if (!reactDateTime) return; reactDateTime = null; L.bus.emit('pptx-date-time-dialog'); };
  D.reactSymbolCatalog = () => reactSymbol;
  D.reactSymbolAction = (action, value) => {
    if (!reactSymbol) return null;
    if (action === 'close') { reactSymbol = null; reactSymbolCallback = null; L.bus.emit('pptx-symbol-dialog'); return null; }
    if (action === 'select') {
      if (typeof value === 'string') { reactSymbol.selected = value; reactSymbol.code = value.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); }
    } else if (action === 'font' && reactSymbol.fonts.includes(value)) reactSymbol.font = value;
    else if (action === 'subset' && Number.isInteger(value) && value >= 0 && value < reactSymbol.subsets.length) {
      reactSymbol.subset = value;
      const [, first, last] = SYM_SETS[value];
      reactSymbol.glyphs = Array.from({ length: last - first + 1 }, (_, i) => String.fromCodePoint(first + i));
    }
    else if (action === 'insert' && reactSymbol.selected) {
      const character = reactSymbol.selected;
      const index = reactSymbol.recent.indexOf(character); if (index >= 0) reactSymbol.recent.splice(index, 1);
      reactSymbol.recent.unshift(character); reactSymbol.recent.length = 16;
      recentSyms.splice(0, recentSyms.length, ...reactSymbol.recent);
      if (reactSymbolCallback) reactSymbolCallback(character);
      else if (L.te.active()) { L.hist.push('Insert Symbol'); L.te.insertText(character); }
      reactSymbol = null; reactSymbolCallback = null; L.bus.emit('pptx-symbol-dialog'); return null;
    }
    return reactSymbol;
  };
  D.reactChangeCaseCatalog = () => reactChangeCase;
  D.reactChangeCaseAction = (action, mode) => {
    if (!reactChangeCase) return false;
    if (action === 'close') { reactChangeCase = null; L.bus.emit('pptx-change-case-dialog'); return true; }
    if (action === 'apply') {
      const selected = mode || reactChangeCase.mode;
      if (!['sentence', 'lower', 'upper', 'title', 'toggle'].includes(selected)) return false;
      D.applyCase(selected);
      reactChangeCase = null;
      L.bus.emit('pptx-change-case-dialog');
      return true;
    }
    if (action === 'select' && ['sentence', 'lower', 'upper', 'title', 'toggle'].includes(mode)) reactChangeCase.mode = mode;
    return true;
  };
  D.reactLineSpacingCatalog = () => reactLineSpacing;
  D.reactLineSpacingAction = (action, values) => {
    if (!reactLineSpacing) return false;
    if (action === 'close') { reactLineSpacing = null; L.bus.emit('pptx-line-spacing-dialog'); return true; }
    if (action !== 'apply' || !values || !['line', 'before', 'after'].every((key) => Number.isFinite(values[key]?.value) && values[key].value >= 0 && values[key].value <= 1584 && ['l', 'p'].includes(values[key].unit))) return false;
    const value = (item) => item.unit === 'p' ? { pts: L.round(item.value, 2) } : { pct: L.round(item.value * 100, 1) };
    L.fmt.para((p) => { p.pp = Object.assign({}, p.pp, { lnSpc: value(values.line), spcBef: value(values.before), spcAft: value(values.after) }); }, 'Line Spacing');
    reactLineSpacing = null;
    L.bus.emit('pptx-line-spacing-dialog');
    return true;
  };
  D.reactHeaderFooterCatalog = () => reactHeaderFooter;
  D.reactHeaderFooterAction = (action, values, all) => {
    if (!reactHeaderFooter) return false;
    if (action === 'close') { reactHeaderFooter = null; L.bus.emit('pptx-header-footer-dialog'); return true; }
    if (action !== 'apply' || !values || typeof values.dt !== 'boolean' || typeof values.dtAuto !== 'boolean' || !L.DATE_FORMATS.includes(values.dtFmt) || typeof values.dtText !== 'string' || typeof values.num !== 'boolean' || typeof values.ftr !== 'boolean' || typeof values.ftrText !== 'string' || typeof values.notOnTitle !== 'boolean') return false;
    L.hist.push('Header and Footer');
    if (all) { L.pres.hf = values; for (const slide of L.pres.slides) delete slide.hf; }
    else { const slide = E().slide(); if (slide) slide.hf = values; }
    E().render(); L.bus.emit('slides-changed');
    reactHeaderFooter = null; L.bus.emit('pptx-header-footer-dialog');
    return true;
  };
  D.reactSetupShowCatalog = () => reactSetupShow;
  D.reactSetupShowResolveColor = (color) => L.model.resolveColor(color, design());
  D.reactSetupShowOpenPenColor = (anchor) => {
    if (!reactSetupShow || !(anchor instanceof HTMLElement)) return false;
    ui.colorMenu(anchor, { mode: 'plain', grid: true }, (value) => {
      if (typeof value !== 'string') return;
      reactSetupShow.penColor = value;
      L.bus.emit('pptx-setup-show-dialog');
    });
    return true;
  };
  D.reactSetupShowAction = (action, values) => {
    if (!reactSetupShow) return false;
    if (action === 'close') { reactSetupShow = null; L.bus.emit('pptx-setup-show-dialog'); return true; }
    if (action !== 'apply' || !values || typeof values.kiosk !== 'boolean' || typeof values.loop !== 'boolean' || typeof values.noAnimation !== 'boolean' || typeof values.range !== 'boolean' || !Number.isFinite(values.from) || !Number.isFinite(values.to) || typeof values.useTimings !== 'boolean' || typeof values.penColor !== 'string' || !values.penColor || values.penColor.length > 40) return false;
    const show = L.pres.show;
    L.hist.push('Set Up Show');
    Object.assign(show, { kiosk: values.kiosk, loop: values.loop || values.kiosk, noAnim: values.noAnimation, penColor: values.penColor[0] === '#' ? values.penColor : L.model.resolveColor(values.penColor, design()), from: values.range ? L.clamp(values.from, 1, reactSetupShow.slideCount) : 0, to: values.range ? L.clamp(values.to, 1, reactSetupShow.slideCount) : 0, useTimings: values.useTimings });
    reactSetupShow = null; L.bus.emit('pptx-setup-show-dialog');
    return true;
  };
  D.reactReplaceFontsCatalog = () => reactReplaceFonts;
  D.reactReplaceFontsAction = (action, from, to) => {
    if (!reactReplaceFonts) return false;
    if (action === 'close') { reactReplaceFonts = null; L.bus.emit('pptx-replace-fonts-dialog'); return true; }
    if (action !== 'apply' || !reactReplaceFonts.used.includes(from) || !L.FONT_LIST.includes(to)) return false;
    replaceFont(from, to);
    reactReplaceFonts = null; L.bus.emit('pptx-replace-fonts-dialog');
    return true;
  };
  D.reactInsertTableCatalog = () => reactInsertTable;
  D.reactInsertTableAction = (action, rows, columns) => {
    if (!reactInsertTable) return false;
    if (action === 'close') { reactInsertTable = null; reactInsertTableCallback = null; L.bus.emit('pptx-insert-table-dialog'); return true; }
    if (action !== 'apply' || !Number.isInteger(rows) || !Number.isInteger(columns) || rows < 1 || rows > 75 || columns < 1 || columns > 75) return false;
    const callback = reactInsertTableCallback;
    reactInsertTable = null; reactInsertTableCallback = null; L.bus.emit('pptx-insert-table-dialog');
    callback?.(rows, columns);
    return true;
  };
  D.reactFontCatalog = () => reactFont;
  D.reactFontOpenColor = (anchor) => {
    if (!reactFont || !(anchor instanceof HTMLElement)) return false;
    ui.colorMenu(anchor, { mode: 'plain', grid: true, automatic: true }, (value) => {
      if (value && value.auto) reactFont.color = null;
      else if (typeof value === 'string') reactFont.color = value;
      else return;
      L.bus.emit('pptx-font-dialog');
    });
    return true;
  };
  D.reactFontResolveColor = (color) => L.model.resolveColor(color, design());
  D.reactFontAction = (action, values) => {
    if (!reactFont) return false;
    if (action === 'close') { reactFont = null; L.bus.emit('pptx-font-dialog'); return true; }
    if (action !== 'apply' || !values || typeof values.font !== 'string' || !Number.isFinite(values.size) || typeof values.bold !== 'boolean' || typeof values.italic !== 'boolean' || typeof values.underline !== 'boolean' || typeof values.shadow !== 'boolean' || typeof values.emboss !== 'boolean' || ![null, 'super', 'sub'].includes(values.script) || !Number.isFinite(values.offset) || (values.color !== null && typeof values.color !== 'string')) return false;
    const props = { font: values.font.trim() || reactFont.font, sz: L.clamp(values.size || reactFont.size, 1, 4000), b: values.bold, i: values.italic, u: values.underline ? 'sng' : undefined, shd: values.shadow || undefined, emb: values.emboss || undefined, base: values.script === 'super' ? L.clamp(values.offset, 1, 100) * 1000 : values.script === 'sub' ? -L.clamp(values.offset, 1, 100) * 1000 : undefined, color: values.color || undefined };
    applyFont(props);
    reactFont = null; L.bus.emit('pptx-font-dialog');
    return true;
  };
  D.reactBulletsCatalog = () => reactBullets;
  D.reactBulletsOpenColor = (anchor) => {
    if (!reactBullets || !(anchor instanceof HTMLElement)) return false;
    ui.colorMenu(anchor, { mode: 'plain', grid: true, automatic: true }, (value) => {
      if (value && value.auto) reactBullets.color = null;
      else if (typeof value === 'string') reactBullets.color = value;
      else return;
      L.bus.emit('pptx-bullets-dialog');
    });
    return true;
  };
  D.reactBulletsCustomize = () => {
    if (!reactBullets) return false;
    D.symbol((character) => { reactBullets.customChar = character; reactBullets.pick = character; reactBullets.kind = 'char'; BULLET_SETS[7] = character; L.bus.emit('pptx-bullets-dialog'); });
    return true;
  };
  D.reactBulletsAction = (action, values) => {
    if (!reactBullets) return false;
    if (action === 'close') { reactBullets = null; L.bus.emit('pptx-bullets-dialog'); return true; }
    if (action === 'select' && values && ['char', 'num'].includes(values.kind) && (values.value === null || (values.kind === 'char' ? BULLET_SETS.includes(values.value) : NUM_SETS.includes(values.value)))) {
      reactBullets.kind = values.kind; reactBullets.pick = values.value;
      if (values.kind === 'char' && values.value) reactBullets.customChar = values.value;
      return true;
    }
    if (action !== 'apply' || !values || !['char', 'num'].includes(values.kind) || values.pick !== null && !(values.kind === 'char' ? BULLET_SETS.includes(values.pick) : NUM_SETS.includes(values.pick)) || !Number.isFinite(values.size) || values.size < 25 || values.size > 400 || !Number.isFinite(values.start) || values.start < 1 || values.start > 9999 || values.color !== null && typeof values.color !== 'string') return false;
    let bullet;
    if (values.pick == null) bullet = { t: 'none' };
    else if (values.kind === 'num') bullet = { t: 'num', scheme: values.pick, start: values.start, sz: values.size / 100, c: values.color || undefined };
    else bullet = { t: 'char', ch: values.pick || values.customChar, sz: values.size / 100, c: values.color || undefined };
    if (bullet.sz === 1) delete bullet.sz;
    applyBullets(bullet);
    reactBullets = null; L.bus.emit('pptx-bullets-dialog');
    return true;
  };
  D.reactActionSettingsCatalog = () => reactActionSettings;
  D.reactActionSettingsAction = (action, values) => {
    if (!reactActionSettings) return false;
    if (action === 'close') { reactActionSettings = null; reactActionShapes = null; L.bus.emit('pptx-action-settings-dialog'); return true; }
    if (action !== 'apply' || !values || typeof values.none !== 'boolean' || !['next', 'prev', 'first', 'last', 'lastViewed', 'end', 'slide', 'url'].includes(values.kind) || typeof values.slide !== 'string' || typeof values.url !== 'string') return false;
    const shapes = reactActionShapes;
    const link = values.none ? null : values.kind === 'slide' ? { slide: values.slide } : values.kind === 'url' ? { url: values.url } : { action: values.kind };
    E().commit('Action Settings', () => { for (const shape of shapes) { if (link) shape.link = link; else delete shape.link; } });
    reactActionSettings = null; reactActionShapes = null; L.bus.emit('pptx-action-settings-dialog');
    return true;
  };
  D.reactPrintCatalog = () => reactPrint;
  D.reactPrintAction = (action, values) => {
    if (!reactPrint) return false;
    if (action === 'close') { reactPrint = null; L.bus.emit('pptx-print-dialog'); return true; }
    if (action !== 'preview' || !values || !['slides', 'handouts', 'notes', 'outline'].includes(values.what) || ![1, 2, 3, 4, 6, 9].includes(values.per) || !['color', 'gray', 'bw'].includes(values.color) || typeof values.frame !== 'boolean' || typeof values.hidden !== 'boolean' || !['all', 'current', 'range'].includes(values.rangeMode) || typeof values.range !== 'string') return false;
    const range = values.rangeMode === 'all' ? null : values.rangeMode === 'current' ? [E().idx] : parseRange(values.range, reactPrint.slides);
    Object.assign(L.app.printOpts, { what: values.what, per: values.per, color: values.color, frame: values.frame, hidden: values.hidden, range });
    reactPrint = null; L.bus.emit('pptx-print-dialog');
    L.app.setView('preview');
    return true;
  };
  D.reactPropertiesCatalog = () => reactProperties;
  D.reactPropertiesUpdate = (key, value) => {
    if (!reactProperties || !['title', 'subject', 'author', 'company', 'category', 'keywords', 'comments'].includes(key) || typeof value !== 'string') return false;
    reactProperties.values[key] = value;
    L.pres.props[key] = value;
    if (key === 'title') L.pres.title = value;
    return true;
  };
  D.reactPropertiesClose = () => { if (!reactProperties) return; reactProperties = null; L.bus.emit('pptx-properties-dialog'); };
  D.reactAutoCorrectCatalog = () => reactAutoCorrect;
  D.reactAutoCorrectAction = (action, value, secondary) => {
    if (!reactAutoCorrect) return false;
    const options = L.app.opts;
    if (action === 'close') { reactAutoCorrect = null; L.bus.emit('pptx-autocorrect-dialog'); return true; }
    if (action === 'option' && ['autocorrect', 'capSentence', 'smartQuotes'].includes(value) && typeof secondary === 'boolean') { options[value] = secondary; reactAutoCorrect.options = { ...reactAutoCorrect.options, [value]: secondary }; return true; }
    if (action === 'select' && typeof value === 'string') { reactAutoCorrect.selected = value; reactAutoCorrect.replacement = L.te.AUTOCORRECT[value] || ''; return true; }
    if (action === 'edit' && typeof value === 'string' && typeof secondary === 'string') { reactAutoCorrect.selected = value; reactAutoCorrect.replacement = secondary; return true; }
    if (action === 'add' && typeof value === 'string' && value.trim() && typeof secondary === 'string') { L.te.AUTOCORRECT[value] = secondary; reactAutoCorrect.selected = value; reactAutoCorrect.replacement = secondary; reactAutoCorrect.entries = Object.entries(L.te.AUTOCORRECT).sort(([a], [b]) => a.localeCompare(b)); return true; }
    if (action === 'delete' && reactAutoCorrect.selected && Object.hasOwn(L.te.AUTOCORRECT, reactAutoCorrect.selected)) { delete L.te.AUTOCORRECT[reactAutoCorrect.selected]; reactAutoCorrect.selected = ''; reactAutoCorrect.replacement = ''; reactAutoCorrect.entries = Object.entries(L.te.AUTOCORRECT).sort(([a], [b]) => a.localeCompare(b)); return true; }
    if (action === 'save') { L.app.saveOpts(); reactAutoCorrect = null; L.bus.emit('pptx-autocorrect-dialog'); return true; }
    return false;
  };
  const M = L.model;
  const IN = (pt) => L.round(pt / 72, 2);
  const PT = (inch) => inch * 72;
  const design = () => M.design(L.pres, E().slide());
  const fillCSS = (f) => L.render.fillCSS(f, design());

  /* ---------- shared pickers ---------- */
  /** a dropdown button that edits a Fill; returns element with .get() */
  function fillPicker(fill, mode, onChange) {
    let cur = L.clone(fill || { t: 'none' });
    const sw = h('span', { class: 'fp-sw', style: 'display:inline-block;width:70px;height:14px;border:1px solid #555;vertical-align:middle' });
    const btn = h('button', { type: 'button', class: 'btn small', style: 'min-width:100px;display:inline-flex;align-items:center;gap:6px' }, sw, h('span', { class: 'dd-arrow' }));
    const paint = () => {
      if (!cur || cur.t === 'none') { sw.style.background = 'repeating-linear-gradient(45deg,#fff 0 4px,#eee 4px 8px)'; sw.title = mode === 'line' ? L.t('No Line') : L.t('No Fill'); }
      else sw.style.background = mode === 'line' ? L.model.resolveColor(cur.c, design()) : fillCSS(cur);
    };
    paint();
    btn.addEventListener('click', () => {
      ui.colorMenu(btn, { mode: mode === 'line' ? 'line' : mode === 'plain' ? 'plain' : 'fill', effects: mode !== 'plain' && mode !== 'line', grid: mode === 'plain' }, async (v) => {
        if (v && v.none) cur = mode === 'line' ? { t: 'none' } : { t: 'none' };
        else if (v && v.effects) { const f = await D.fillEffects(cur && cur.t !== 'none' ? cur : { t: 'solid', c: 'accent1', a: 1 }); if (!f) return; cur = f; }
        else if (typeof v === 'string') cur = mode === 'line' ? Object.assign({}, cur && cur.t !== 'none' ? cur : { w: 0.75, dash: 'solid' }, { c: v, t: undefined }) : { t: 'solid', c: v, a: cur && cur.a != null && cur.t === 'solid' ? cur.a : 1 };
        if (cur && cur.t === undefined) delete cur.t;
        paint();
        onChange && onChange(cur);
      });
    });
    btn.get = () => cur;
    btn.set = (f) => { cur = f; paint(); };
    return btn;
  }
  D.fillPicker = fillPicker;
  function colorPick(c, onChange, opts) {
    let cur = c;
    const sw = h('span', { style: 'display:inline-block;width:60px;height:14px;border:1px solid #555;vertical-align:middle' });
    const btn = h('button', { type: 'button', class: 'btn small', style: 'min-width:90px;display:inline-flex;align-items:center;gap:6px' }, sw, h('span', { class: 'dd-arrow' }));
    const paint = () => { sw.style.background = cur ? L.model.resolveColor(cur, design()) : 'repeating-linear-gradient(45deg,#fff 0 4px,#eee 4px 8px)'; };
    paint();
    btn.addEventListener('click', () => ui.colorMenu(btn, { mode: 'plain', grid: true, automatic: opts && opts.automatic }, (v) => { if (v && v.auto) cur = null; else if (typeof v === 'string') cur = v; else return; paint(); onChange && onChange(cur); }));
    btn.get = () => cur;
    return btn;
  }
  D.colorPick = colorPick;
  const LINE_WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
  const DASHES = [['solid', L.N('Solid')], ['sysDot', L.N('Round Dot')], ['sysDash', L.N('Square Dot')], ['dash', L.N('Dash')], ['dashDot', L.N('Dash Dot')], ['lgDash', L.N('Long Dash')], ['lgDashDot', L.N('Long Dash Dot')], ['lgDashDotDot', L.N('Long Dash Dot Dot')]];
  const ARROWS = [['none', L.N('None')], ['triangle', L.N('Arrow')], ['arrow', L.N('Open Arrow')], ['stealth', L.N('Stealth Arrow')], ['diamond', L.N('Diamond Arrow')], ['oval', L.N('Oval Arrow')]];
  const SIZES = [['sm', L.N('Small')], ['med', L.N('Medium')], ['lg', L.N('Large')]];
  D.LINE_WEIGHTS = LINE_WEIGHTS; D.DASHES = DASHES; D.ARROWS = ARROWS;

  /* ---------- Font ---------- */
  D.font = function () {
    const st = L.fmt.state();
    if (!st) { ui.msg('Select some text or a text object first.'); return; }
    const o = { font: st.font, sz: st.sz, b: st.b, i: st.i, u: st.u, shd: st.shd, emb: !!(st.r && st.r.emb), base: (st.r && st.r.base) || 0, color: st.r ? st.r.color : null };
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const fonts = [...L.FONT_LIST]; if (o.font && !fonts.includes(o.font)) fonts.unshift(o.font);
      const styles = [['regular', L.N('Regular')], ['italic', L.N('Italic')], ['bold', L.N('Bold')], ['bolditalic', 'Bold Italic']];
      reactFont = { font: o.font || fonts[0] || 'Arial', fonts, styles, style: o.b && o.i ? 'bolditalic' : o.b ? 'bold' : o.i ? 'italic' : 'regular', size: o.sz || 18, sizes: [...L.SIZE_LIST], underline: o.u === 'sng', shadow: !!o.shd, emboss: !!o.emb, script: o.base > 0 ? 'super' : o.base < 0 ? 'sub' : null, offset: o.base ? Math.abs(o.base / 1000) : 30, color: o.color || null };
      L.bus.emit('pptx-font-dialog');
      return;
    }
    const fontList = h('select', { id: 'fd-font', size: 7, style: 'width:180px' });
    L.FONT_LIST.forEach((f) => { const op = h('option', { value: f, text: f, translate: 'no' }); if (f === o.font) op.selected = true; fontList.appendChild(op); });
    if (!L.FONT_LIST.includes(o.font)) fontList.insertBefore(h('option', { value: o.font, text: o.font, selected: true, translate: 'no' }), fontList.firstChild);
    const fontIn = h('input', { type: 'text', id: 'fd-font-in', value: o.font, style: 'width:180px' });
    const styleSel = h('select', { id: 'fd-style', size: 4, style: 'width:110px' }, ...[L.N('Regular'), L.N('Italic'), L.N('Bold'), 'Bold Italic'].map((s) => h('option', { value: s, text: s })));
    styleSel.value = o.b && o.i ? 'Bold Italic' : o.b ? 'Bold' : o.i ? 'Italic' : 'Regular';
    const sizeIn = h('input', { type: 'text', id: 'fd-size', value: o.sz, style: 'width:60px' });
    const sizeSel = h('select', { size: 4, style: 'width:60px' }, ...L.SIZE_LIST.map((s) => h('option', { value: s, text: s })));
    sizeSel.value = String(o.sz);
    const preview = h('div', { class: 'preview-box font-preview', text: 'AaBbYyZz', translate: 'no' });
    const upd = () => {
      const r = preview.style;
      r.fontFamily = L.fontStack(fontIn.value);
      r.fontWeight = /Bold/.test(styleSel.value) ? '700' : '400';
      r.fontStyle = /Italic/.test(styleSel.value) ? 'italic' : 'normal';
      r.textDecoration = u.input.checked ? 'underline' : 'none';
      r.textShadow = s.input.checked ? '2px 2px 0 rgba(0,0,0,.35)' : emb.input.checked ? '-1px -1px 0 #fff, 1px 1px 0 #666' : 'none';
      r.color = col.get() ? L.model.resolveColor(col.get(), design()) : '#000';
      r.fontSize = Math.min(40, Math.max(10, parseFloat(sizeIn.value) || 18)) + 'px';
    };
    fontList.addEventListener('change', () => { fontIn.value = fontList.value; upd(); });
    fontIn.addEventListener('input', upd);
    styleSel.addEventListener('change', upd);
    sizeSel.addEventListener('change', () => { sizeIn.value = sizeSel.value; upd(); });
    sizeIn.addEventListener('input', upd);
    const u = ui.check('&Underline', o.u, upd), s = ui.check('&Shadow', o.shd, upd), emb = ui.check('&Emboss', o.emb, upd);
    const sup = ui.check('Su&perscript', o.base > 0, () => { if (sup.input.checked) sub.input.checked = false; }), sub = ui.check('Su&bscript', o.base < 0, () => { if (sub.input.checked) sup.input.checked = false; });
    const off = ui.spin({ value: o.base ? Math.abs(o.base / 1000) : 30, min: 1, max: 100, step: 1, unit: '%', dec: 0 });
    const col = colorPick(o.color, upd, { automatic: true });
    const body = h('div', { class: 'col' },
      h('div', { class: 'row' },
        h('div', { class: 'col' }, h('label', { for: 'fd-font-in', html: ui.ampHTML('&Font:') }), fontIn, fontList),
        h('div', { class: 'col' }, h('label', { for: 'fd-style', html: ui.ampHTML('F&ont style:') }), styleSel),
        h('div', { class: 'col' }, h('label', { for: 'fd-size', html: ui.ampHTML('&Size:') }), sizeIn, sizeSel)),
      h('div', { class: 'row' },
        ui.group('Effects', h('div', { class: 'row' }, h('div', { class: 'col' }, u, s, emb), h('div', { class: 'col' }, sup, sub, h('div', { class: 'field' }, h('label', { text: 'Offset:' }), off)))),
        h('div', { class: 'col' }, h('label', { html: ui.ampHTML('&Color:') }), col)),
      preview,
      h('div', { class: 'tp-note', text: 'This is a TrueType font. The same font will be used on both your printer and your screen.' }));
    upd();
    ui.dialog({
      title: 'Font', body, width: 470, buttons: [{
        label: 'OK', primary: true, onClick: () => {
          const props = {
            font: fontIn.value.trim() || o.font, sz: L.clamp(parseFloat(sizeIn.value) || o.sz, 1, 4000), b: /Bold/.test(styleSel.value), i: /Italic/.test(styleSel.value),
            u: u.input.checked ? 'sng' : undefined, shd: s.input.checked || undefined, emb: emb.input.checked || undefined,
            base: sup.input.checked ? off.get() * 1000 : sub.input.checked ? -off.get() * 1000 : undefined, color: col.get() || undefined,
          };
          applyFont(props);
        },
      }, { label: 'Cancel' }],
    });
  };
  function applyFont(props) {
    L.fmt.run((r) => { const n = Object.assign({}, r, props); for (const k in props) if (props[k] === undefined) delete n[k]; if (props.color) delete n.fill; return n; }, L.N('Font'), (wa) => { wa.font = props.font; wa.b = props.b; wa.i = props.i; });
  }

  /* ---------- Bullets and Numbering ---------- */
  const BULLET_SETS = [null, '•', '○', '■', '❖', '➢', '✓', '▪'];
  const NUM_SETS = [null, 'arabicPeriod', 'arabicParenR', 'romanUcPeriod', 'alphaUcPeriod', 'alphaLcParenR', 'alphaLcPeriod', 'romanLcPeriod'];
  function applyBullets(bu) {
    L.fmt.para((p, i, t) => {
      p.pp = Object.assign({}, p.pp, { bu });
      const cls = L.style.cls(t ? t.sh : null);
      if (bu.t !== 'none' && cls === 'other' && !(p.pp.marL > 0)) { p.pp.marL = 27 + (p.lvl || 0) * 36; p.pp.indent = -27; }
      if (bu.t === 'none' && cls === 'other') { delete p.pp.marL; delete p.pp.indent; }
    }, 'Bullets and Numbering');
  }
  D.bullets = function (tabIdx) {
    const st = L.fmt.state();
    if (!st || st.wa) { ui.msg('Select some text first.'); return; }
    const cur = st.ps.bu || { t: 'none' };
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactBullets = {
        kind: cur.t === 'num' ? 'num' : 'char', pick: cur.t === 'char' ? cur.ch : cur.t === 'num' ? cur.scheme : null,
        customChar: cur.t === 'char' ? cur.ch : '•', size: Math.round((cur.sz || 1) * 100), color: cur.c || null,
        start: cur.start || 1, tab: tabIdx != null ? tabIdx : cur.t === 'num' ? 1 : 0,
        bullets: BULLET_SETS.map((value, index) => ({ value, preview: value == null ? ['None'] : [1, 2, 3].map((n) => value || '•'), index })),
        numbers: NUM_SETS.map((value, index) => ({ value, preview: value == null ? ['None'] : [1, 2, 3].map((n) => L.render.numberText(value, n)), index })),
      };
      L.bus.emit('pptx-bullets-dialog');
      return;
    }
    let pick = cur.t === 'char' ? cur.ch : cur.t === 'num' ? cur.scheme : null;
    let kind = cur.t === 'num' ? 'num' : 'char';
    let customCh = cur.t === 'char' ? cur.ch : '•';
    const sizeSp = ui.spin({ value: Math.round((cur.sz || 1) * 100), min: 25, max: 400, step: 5, unit: '%', dec: 0 });
    const sizeSp2 = ui.spin({ value: Math.round((cur.sz || 1) * 100), min: 25, max: 400, step: 5, unit: '%', dec: 0 });
    const col = colorPick(cur.c || null, null, { automatic: true });
    const col2 = colorPick(cur.c || null, null, { automatic: true });
    const startSp = ui.spin({ value: cur.start || 1, min: 1, max: 9999, step: 1, dec: 0 });
    const bg = h('div', { class: 'bul-grid' });
    const ng = h('div', { class: 'num-grid' });
    const mkBtn = (grid, val, label, isNum) => {
      const b = h('button', { type: 'button' });
      if (val == null) b.appendChild(h('span', { text: 'None', style: 'font-size:13px;align-self:center' }));
      else for (let i = 1; i <= 3; i++) b.appendChild(h('div', null, h('span', { text: isNum ? L.render.numberText(val, i) : val, style: 'min-width:16px' }), h('i')));
      if (pick === val && ((isNum && kind === 'num') || (!isNum && kind === 'char') || val == null)) b.classList.add('on');
      b.addEventListener('click', () => { L.$$('button', bg).concat(L.$$('button', ng)).forEach((x) => x.classList.remove('on')); b.classList.add('on'); pick = val; kind = isNum ? 'num' : 'char'; if (!isNum && val) customCh = val; });
      b.addEventListener('dblclick', () => { okBtn.click(); });
      grid.appendChild(b);
      void label;
    };
    BULLET_SETS.forEach((v) => mkBtn(bg, v, '', false));
    NUM_SETS.forEach((v) => mkBtn(ng, v, '', true));
    const customize = ui.button('C&ustomize...', () => D.symbol((ch) => { customCh = ch; pick = ch; kind = 'char'; const b = L.$$('button', bg)[7]; L.clear(b); for (let i = 0; i < 3; i++) b.appendChild(h('div', null, h('span', { text: ch, style: 'min-width:16px' }), h('i'))); L.$$('button', bg).forEach((x) => x.classList.remove('on')); b.classList.add('on'); BULLET_SETS[7] = ch; }));
    const tabs = ui.tabs([
      { label: 'Bulle&ted', body: h('div', { class: 'col' }, bg, h('div', { class: 'row', style: 'align-items:center;margin-top:8px' }, h('label', { text: 'Size:' }), sizeSp, h('span', { text: '% of text' }), h('label', { text: 'Color:' }), col, customize)) },
      { label: '&Numbered', body: h('div', { class: 'col' }, ng, h('div', { class: 'row', style: 'align-items:center;margin-top:8px' }, h('label', { text: 'Size:' }), sizeSp2, h('span', { text: '% of text' }), h('label', { text: 'Color:' }), col2, h('label', { text: 'Start at:' }), startSp)) },
    ], tabIdx != null ? tabIdx : kind === 'num' ? 1 : 0);
    const d = ui.dialog({
      title: 'Bullets and Numbering', body: tabs, width: 420, buttons: [{
        label: 'OK', primary: true, onClick: () => {
          let bu;
          if (pick == null) bu = { t: 'none' };
          else if (kind === 'num') bu = { t: 'num', scheme: pick, start: startSp.get(), sz: sizeSp2.get() / 100, c: col2.get() || undefined };
          else bu = { t: 'char', ch: pick || customCh, sz: sizeSp.get() / 100, c: col.get() || undefined };
          if (bu.sz === 1) delete bu.sz;
          applyBullets(bu);
        },
      }, { label: 'Cancel' }],
    });
    const okBtn = d.buttons[0];
  };

  /* ---------- Line spacing ---------- */
  D.lineSpacing = function () {
    const st = L.fmt.state();
    if (!st || st.wa) { ui.msg('Select some text first.'); return; }
    const ps = st.ps;
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const read = (spacing, fallback) => spacing && spacing.pts != null ? { value: spacing.pts, unit: 'p' } : { value: spacing ? (spacing.pct || 0) / 100 : fallback, unit: 'l' };
      reactLineSpacing = { fontSize: st.sz || 18, line: read(ps.lnSpc, 1), before: read(ps.spcBef, 0.2), after: read(ps.spcAft, 0) };
      L.bus.emit('pptx-line-spacing-dialog');
      return;
    }
    const mk = (sp, def) => {
      const isPts = sp && sp.pts != null;
      const v = sp ? (isPts ? sp.pts : (sp.pct || 0) / 100) : def;
      const spin = ui.spin({ value: v, min: 0, max: 1584, step: isPts ? 1 : 0.1, dec: 2 });
      const unit = ui.select([['l', 'Lines'], ['p', 'Points']], isPts ? 'p' : 'l', (u2) => { spin.set(u2 === 'p' ? L.round(spin.get() * (st.sz || 18) * 1.2, 1) : L.round(spin.get() / ((st.sz || 18) * 1.2), 2)); });
      return { spin, unit, get: () => (unit.value === 'p' ? { pts: spin.get() } : { pct: L.round(spin.get() * 100, 1) }) };
    };
    const ls = mk(ps.lnSpc, 1), sb = mk(ps.spcBef, 0.2), sa = mk(ps.spcAft, 0);
    const body = h('div', { class: 'col' },
      ui.group('Line spacing', h('div', { class: 'row' }, ls.spin, ls.unit)),
      ui.group('Before paragraph', h('div', { class: 'row' }, sb.spin, sb.unit)),
      ui.group('After paragraph', h('div', { class: 'row' }, sa.spin, sa.unit)));
    ui.dialog({ title: 'Line Spacing', body, width: 280, buttons: [{ label: 'OK', primary: true, onClick: () => L.fmt.para((p) => { p.pp = Object.assign({}, p.pp, { lnSpc: ls.get(), spcBef: sb.get(), spcAft: sa.get() }); }, 'Line Spacing') }, { label: 'Cancel' }] });
  };

  /* ---------- Change Case ---------- */
  D.changeCase = function () {
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactChangeCase = { mode: 'sentence', options: [['sentence', '&Sentence case.'], ['lower', L.N('&lowercase')], ['upper', L.N('&UPPERCASE')], ['title', '&Title Case'], ['toggle', L.N('t&OGGLE cASE')]] };
      L.bus.emit('pptx-change-case-dialog');
      return;
    }
    let mode = 'sentence';
    const opts = [['sentence', '&Sentence case.'], ['lower', L.N('&lowercase')], ['upper', L.N('&UPPERCASE')], ['title', '&Title Case'], ['toggle', L.N('t&OGGLE cASE')]];
    const body = h('div', { class: 'col' }, ...opts.map(([k, l], i) => ui.radio('cc', l, i === 0, () => { mode = k; })));
    ui.dialog({ title: 'Change Case', body, width: 240, buttons: [{ label: 'OK', primary: true, onClick: () => D.applyCase(mode) }, { label: 'Cancel' }] });
  };
  D.applyCase = function (mode) {
    const fn = { sentence: L.sentenceCase, lower: (s) => s.toLowerCase(), upper: (s) => s.toUpperCase(), title: L.titleCase, toggle: L.toggleCase }[mode];
    L.fmt.run((r) => (r.fld || typeof r.t !== 'string' ? r : Object.assign({}, r, { t: fn(r.t) })), 'Change Case', (wa) => { wa.text = fn(wa.text); });
  };

  function replaceFont(a, b) {
    L.hist.push('Replace Fonts');
    for (const id in L.pres.designs) { const d = L.pres.designs[id]; if (d.fonts.major === a) d.fonts.major = b; if (d.fonts.minor === a) d.fonts.minor = b; }
    const fix = (tx, d) => { for (const p of tx.ps) { for (const r of p.rs) if (L.style.font(r.font || '+mn', d) === a && r.font && r.font[0] !== '+') r.font = b; if (p.end && p.end.font === a) p.end.font = b; } };
    for (const s of L.pres.slides) { const d = M.design(L.pres, s); M.walk(s.shapes, (sh) => { if (sh.tx) fix(sh.tx, d); if (sh.tbl) sh.tbl.rows.forEach((r) => r.cells.forEach((c) => fix(c.tx, d))); if (sh.wa && sh.wa.font === a) sh.wa.font = b; }); }
    E().render(); L.bus.emit('slides-changed');
  }

  /* ---------- Replace Fonts ---------- */
  D.replaceFonts = function () {
    const used = new Set();
    const scan = (tx, d) => { for (const p of tx.ps) for (const r of p.rs) used.add(L.style.font(r.font || '+mn', d)); };
    for (const s of L.pres.slides) { const d = M.design(L.pres, s); M.walk(s.shapes, (sh) => { if (sh.tx) scan(sh.tx, d); if (sh.tbl) sh.tbl.rows.forEach((r) => r.cells.forEach((c) => scan(c.tx, d))); }); used.add(d.fonts.major); used.add(d.fonts.minor); }
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactReplaceFonts = { used: Array.from(used), fonts: [...L.FONT_LIST] };
      L.bus.emit('pptx-replace-fonts-dialog');
      return;
    }
    const from = ui.select(Array.from(used), Array.from(used)[0]);
    const to = ui.select(L.FONT_LIST, 'Arial');
    const body = h('div', { class: 'col' }, ui.field('&Replace:', from), ui.field('&With:', to));
    ui.dialog({
      title: 'Replace Font', body, width: 330, buttons: [{
        label: 'Replace', primary: true, onClick: () => {
          replaceFont(from.value, to.value);
        },
      }, { label: 'Close' }],
    });
  };

  /* ---------- Fill Effects ---------- */
  const GRAD_PRESETS = [
    [L.N('Early Sunset'), [[0, '#000082'], [0.3, '#66008F'], [0.65, '#BA0066'], [1, '#FF8200']]],
    [L.N('Late Sunset'), [[0, '#000000'], [0.5, '#000040'], [1, '#FF8C00']]],
    [L.N('Nightfall'), [[0, '#000000'], [0.6, '#0A128C'], [1, '#181CC7']]],
    [L.N('Daybreak'), [[0, '#5E9EFF'], [0.4, '#85C2FF'], [0.7, '#C4D6EB'], [1, '#FFEBFA']]],
    [L.N('Horizon'), [[0, '#DCEBF5'], [0.4, '#83A7C3'], [0.5, '#768FB9'], [0.6, '#83A7C3'], [1, '#FFFFFF']]],
    [L.N('Desert'), [[0, '#9C3500'], [0.3, '#E8A15C'], [0.6, '#FFDF9C'], [1, '#FFFFFF']]],
    [L.N('Ocean'), [[0, '#03D4A8'], [0.5, '#21D6E0'], [1, '#0087E6']]],
    [L.N('Calm Water'), [[0, '#CCCCFF'], [0.5, '#99CCFF'], [1, '#CCCCFF']]],
    [L.N('Fire'), [[0, '#FFF200'], [0.45, '#FF7A00'], [0.7, '#FF0300'], [1, '#4D0808']]],
    [L.N('Fog'), [[0, '#8488C4'], [0.5, '#D4DEFF'], [1, '#96AB94']]],
    [L.N('Moss'), [[0, '#DDEBCF'], [0.5, '#9CB86E'], [1, '#156B13']]],
    [L.N('Peacock'), [[0, '#3399FF'], [0.2, '#00CCCC'], [0.5, '#9999FF'], [0.8, '#2E6792'], [1, '#3333CC']]],
    [L.N('Wheat'), [[0, '#FBEAC7'], [0.2, '#FEE7F2'], [0.5, '#FAC77D'], [0.8, '#FBA97D'], [1, '#FBD49C']]],
    [L.N('Parchment'), [[0, '#FFEFD1'], [0.65, '#F0EBD5'], [1, '#D1C39F']]],
    [L.N('Mahogany'), [[0, '#D6B19C'], [0.3, '#D49E6C'], [0.7, '#A65528'], [1, '#663012']]],
    [L.N('Rainbow'), [[0, '#A603AB'], [0.2, '#0819FB'], [0.4, '#1A8D48'], [0.6, '#FFFF00'], [0.8, '#EE3F17'], [1, '#E81766']]],
    [L.N('Gold'), [[0, '#E6DCAC'], [0.15, '#E6D78A'], [0.4, '#C7AC4C'], [0.6, '#E6D78A'], [1, '#E6DCAC']]],
    [L.N('Brass'), [[0, '#825600'], [0.15, '#FFA800'], [0.35, '#825600'], [0.55, '#FFA800'], [0.8, '#825600'], [1, '#FFA800']]],
    [L.N('Chrome'), [[0, '#FFFFFF'], [0.2, '#000000'], [0.4, '#FFFFFF'], [0.6, '#000000'], [0.8, '#FFFFFF'], [1, '#000000']]],
    [L.N('Silver'), [[0, '#FFFFFF'], [0.3, '#E6E6E6'], [0.6, '#7D8496'], [1, '#FFFFFF']]],
    [L.N('Sapphire'), [[0, '#000082'], [0.5, '#0047FF'], [1, '#000082']]],
  ];
  /* procedural textures (generated tiles) */
  const TEXTURES = [L.N('Canvas'), L.N('Denim'), L.N('Paper bag'), L.N('Newsprint'), L.N('Recycled paper'), L.N('Parchment'), L.N('Granite'), L.N('White marble'), L.N('Green marble'), L.N('Sand'), L.N('Cork'), L.N('Oak'), L.N('Walnut'), L.N('Water droplets'), L.N('Blue tissue paper'), L.N('Woven mat'), L.N('Stationery'), L.N('Fabric')];
  const texCache = {};
  function texture(name) {
    if (texCache[name]) return texCache[name];
    const S = 128;
    const c = document.createElement('canvas');
    c.width = S; c.height = S;
    const g = c.getContext('2d');
    let seed = name.split('').reduce((a, ch) => a * 31 + ch.charCodeAt(0), 7) >>> 0;
    const rnd = () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed / 4294967296; };
    const noise = (base, amt, scale) => { const img = g.getImageData(0, 0, S, S); const [r, gg, b] = L.color.hexToRgb(base); for (let i = 0; i < S * S; i++) { const n = (rnd() - 0.5) * amt * (scale || 1); img.data[i * 4] = r + n; img.data[i * 4 + 1] = gg + n; img.data[i * 4 + 2] = b + n; img.data[i * 4 + 3] = 255; } g.putImageData(img, 0, 0); };
    const P = {
      'Canvas': () => { noise('#E9E1CC', 30); g.globalAlpha = 0.15; for (let i = 0; i < S; i += 2) { g.fillStyle = i % 4 ? '#fff' : '#a08a60'; g.fillRect(0, i, S, 1); g.fillRect(i, 0, 1, S); } },
      'Denim': () => { noise('#3E5F8A', 40); g.globalAlpha = 0.25; g.strokeStyle = '#cfe0f5'; for (let i = -S; i < S; i += 3) { g.beginPath(); g.moveTo(i, 0); g.lineTo(i + S, S); g.stroke(); } },
      'Paper bag': () => { noise('#B8915F', 26); g.globalAlpha = 0.2; for (let i = 0; i < 40; i++) { g.fillStyle = rnd() > 0.5 ? '#8a6a3e' : '#d6b88a'; g.fillRect(rnd() * S, rnd() * S, rnd() * 20, 1); } },
      'Newsprint': () => { noise('#E3E0D8', 20); g.globalAlpha = 0.2; g.fillStyle = '#555'; for (let i = 0; i < 200; i++) g.fillRect(rnd() * S, rnd() * S, 1, 1); },
      'Recycled paper': () => { noise('#DCD3C0', 22); g.globalAlpha = 0.5; for (let i = 0; i < 80; i++) { g.fillStyle = ['#7a6b55', '#a49884', '#5e5444'][i % 3]; g.fillRect(rnd() * S, rnd() * S, 1 + rnd() * 2, 1); } },
      'Parchment': () => { const gr = g.createRadialGradient(S / 2, S / 2, 10, S / 2, S / 2, S * 0.75); gr.addColorStop(0, '#FBF1D4'); gr.addColorStop(1, '#E2CC97'); g.fillStyle = gr; g.fillRect(0, 0, S, S); g.globalAlpha = 0.08; noise('#C9B27A', 40); },
      'Granite': () => { noise('#9A9A9A', 60); g.globalAlpha = 0.6; for (let i = 0; i < 400; i++) { g.fillStyle = ['#2b2b2b', '#e8e8e8', '#6b4f4f'][i % 3]; g.fillRect(rnd() * S, rnd() * S, 1 + rnd() * 2, 1 + rnd() * 2); } },
      'White marble': () => { noise('#F2F2F0', 12); g.globalAlpha = 0.35; g.strokeStyle = '#9a9a9a'; for (let k = 0; k < 6; k++) { g.beginPath(); let x = rnd() * S, y = 0; g.moveTo(x, y); while (y < S) { x += (rnd() - 0.5) * 14; y += 6; g.lineTo(x, y); } g.stroke(); } },
      'Green marble': () => { noise('#2E6B4F', 30); g.globalAlpha = 0.45; g.strokeStyle = '#cfe8d8'; for (let k = 0; k < 6; k++) { g.beginPath(); let x = rnd() * S, y = 0; g.moveTo(x, y); while (y < S) { x += (rnd() - 0.5) * 14; y += 6; g.lineTo(x, y); } g.stroke(); } },
      'Sand': () => { noise('#D9C29A', 40); },
      'Cork': () => { noise('#B07A45', 50); g.globalAlpha = 0.5; for (let i = 0; i < 300; i++) { g.fillStyle = rnd() > 0.5 ? '#6b4423' : '#d9a066'; g.fillRect(rnd() * S, rnd() * S, 2, 2); } },
      'Oak': () => { g.fillStyle = '#C49A6C'; g.fillRect(0, 0, S, S); g.globalAlpha = 0.35; for (let y = 0; y < S; y += 2) { g.strokeStyle = rnd() > 0.5 ? '#8a5f34' : '#e0bd8f'; g.beginPath(); g.moveTo(0, y); for (let x = 0; x <= S; x += 8) g.lineTo(x, y + Math.sin(x / 15 + y / 20) * 2); g.stroke(); } },
      'Walnut': () => { g.fillStyle = '#5C3A21'; g.fillRect(0, 0, S, S); g.globalAlpha = 0.35; for (let y = 0; y < S; y += 2) { g.strokeStyle = rnd() > 0.5 ? '#3a2312' : '#8a5a36'; g.beginPath(); g.moveTo(0, y); for (let x = 0; x <= S; x += 8) g.lineTo(x, y + Math.sin(x / 12 + y / 18) * 3); g.stroke(); } },
      'Water droplets': () => { noise('#9CC3D9', 14); for (let i = 0; i < 40; i++) { const x = rnd() * S, y = rnd() * S, r = 2 + rnd() * 6; const gr = g.createRadialGradient(x - r / 3, y - r / 3, 0, x, y, r); gr.addColorStop(0, 'rgba(255,255,255,.9)'); gr.addColorStop(1, 'rgba(60,110,150,.35)'); g.fillStyle = gr; g.beginPath(); g.arc(x, y, r, 0, Math.PI * 2); g.fill(); } },
      'Blue tissue paper': () => { noise('#BFD3EE', 18); g.globalAlpha = 0.15; for (let i = 0; i < 30; i++) { g.strokeStyle = '#fff'; g.beginPath(); g.moveTo(rnd() * S, rnd() * S); g.lineTo(rnd() * S, rnd() * S); g.stroke(); } },
      'Woven mat': () => { g.fillStyle = '#D8C08E'; g.fillRect(0, 0, S, S); for (let y = 0; y < S; y += 16) for (let x = 0; x < S; x += 16) { g.fillStyle = (x + y) % 32 ? '#c2a46b' : '#e8d3a5'; g.fillRect(x, y, 16, 8); g.fillStyle = (x + y) % 32 ? '#e8d3a5' : '#c2a46b'; g.fillRect(x, y + 8, 16, 8); } g.globalAlpha = 0.1; noise('#8a6a3e', 50); },
      'Stationery': () => { g.fillStyle = '#FFFFF4'; g.fillRect(0, 0, S, S); g.strokeStyle = '#B9D3EE'; for (let y = 8; y < S; y += 16) { g.beginPath(); g.moveTo(0, y + 0.5); g.lineTo(S, y + 0.5); g.stroke(); } },
      'Fabric': () => { noise('#B04A5A', 30); g.globalAlpha = 0.2; for (let i = 0; i < S; i += 4) { g.fillStyle = '#fff'; g.fillRect(i, 0, 1, S); g.fillRect(0, i, S, 1); } },
    };
    (P[name] || P.Canvas)();
    g.globalAlpha = 1;
    texCache[name] = c.toDataURL('image/png');
    return texCache[name];
  }
  async function textureMedia(name) {
    const bytes = atob(texture(name).split(',', 2)[1]);
    const blob = new Blob([Uint8Array.from(bytes, (value) => value.charCodeAt(0))], { type: 'image/png' });
    return L.media.add(blob, name + '.png');
  }
  D.texture = texture;

  /** Fill Effects dialog → Promise<Fill|null> */
  D.fillEffects = function (initial) {
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      if (reactFillEffects) return Promise.resolve(null);
      reactFillEffects = { initial: initial && initial.t !== 'none' ? L.clone(initial) : null };
      const result = new Promise((resolve) => { reactFillEffectsResolve = resolve; });
      L.bus.emit('pptx-fill-effects-dialog');
      return result;
    }
    return new Promise((resolve) => {
      let res = null;
      const d0 = design();
      const base = initial && initial.t === 'solid' ? L.model.resolveColor(initial.c, d0) : '#3366CC';
      let g = initial && initial.t === 'grad' ? L.clone(initial) : { t: 'grad', stops: [{ p: 0, c: base, a: 1 }, { p: 1, c: L.color.lighten(base, 0.75), a: 1 }], ang: 90, path: 'lin' };
      let mode = 'two', c1 = g.stops[0].c, c2 = g.stops[g.stops.length - 1].c, dark = 0.5, preset = 0, style = g.path !== 'lin' ? (g.path === 'circle' ? 'center' : 'corner') : g.ang === 0 ? 'vert' : g.ang === 45 ? 'dup' : g.ang === 135 ? 'ddn' : 'horz', variant = 0;
      let tr1 = 0, tr2 = 0;
      let current = initial && initial.t !== 'none' ? L.clone(initial) : g;
      const sample = h('div', { class: 'preview-box', style: 'width:110px;height:110px' }, h('div', { class: 'fx-swatch' }));
      const paintSample = () => { sample.firstChild.style.background = fillCSS(current); };
      const variants = h('div', { class: 'row', style: 'gap:6px' });
      const buildGrad = (v) => {
        let stops;
        const cc1 = L.model.resolveColor(c1, d0), cc2 = mode === 'one' ? (dark < 0.5 ? L.color.darken(cc1, 1 - dark * 2) : L.color.lighten(cc1, (dark - 0.5) * 2)) : L.model.resolveColor(c2, d0);
        if (mode === 'preset') stops = GRAD_PRESETS[preset][1].map(([p, c]) => ({ p, c, a: 1 }));
        else stops = [{ p: 0, c: cc1, a: 1 - tr1 }, { p: 1, c: cc2, a: 1 - tr2 }];
        if (v === 1 || v === 3) stops = stops.map((s) => ({ p: 1 - s.p, c: s.c, a: s.a })).reverse();
        if (v >= 2) stops = stops.map((s) => ({ p: s.p / 2, c: s.c, a: s.a })).concat(stops.slice(0, -1).reverse().map((s) => ({ p: 1 - s.p / 2, c: s.c, a: s.a })));
        const ang = { horz: 90, vert: 0, dup: 315, ddn: 45, corner: 0, center: 0 }[style];
        return { t: 'grad', stops, ang, path: style === 'center' ? 'circle' : style === 'corner' ? 'rect' : 'lin' };
      };
      const refreshVariants = () => {
        L.clear(variants);
        for (let v = 0; v < 4; v++) {
          const gv = buildGrad(v);
          const b = h('button', { type: 'button', style: `width:46px;height:46px;border:1px solid #888;padding:0;background:${fillCSS(gv)};${v === variant ? 'outline:2px solid #0b3fa8;outline-offset:1px' : ''}` });
          b.addEventListener('click', () => { variant = v; current = buildGrad(v); paintSample(); refreshVariants(); });
          variants.appendChild(b);
        }
        current = buildGrad(variant); paintSample();
      };
      const p1 = colorPick(c1, (c) => { c1 = c; refreshVariants(); }), p2 = colorPick(c2, (c) => { c2 = c; refreshVariants(); });
      const darkR = h('input', { type: 'range', min: 0, max: 100, value: 50, id: 'fe-dark' });
      darkR.addEventListener('input', () => { dark = darkR.value / 100; refreshVariants(); });
      const presetSel = ui.select(GRAD_PRESETS.map((p, i) => [i, p[0]]), 0, (v) => { preset = +v; refreshVariants(); });
      const t1 = ui.spin({ value: 0, min: 0, max: 100, step: 5, unit: ' %', dec: 0, onChange: (v) => { tr1 = v / 100; refreshVariants(); } });
      const t2 = ui.spin({ value: 0, min: 0, max: 100, step: 5, unit: ' %', dec: 0, onChange: (v) => { tr2 = v / 100; refreshVariants(); } });
      const colorsArea = h('div', { class: 'col' });
      const setMode = (m) => {
        mode = m;
        L.clear(colorsArea);
        if (m === 'one') colorsArea.append(ui.field('Color &1:', p1), h('div', { class: 'row', style: 'align-items:center' }, h('span', { text: 'Dark' }), darkR, h('span', { text: 'Light' })));
        if (m === 'two') colorsArea.append(ui.field('Color &1:', p1), ui.field('Color &2:', p2));
        if (m === 'preset') colorsArea.append(ui.field('Preset colors:', presetSel));
        refreshVariants();
      };
      const shadingRadios = h('div', { class: 'col' }, ...[['horz', '&Horizontal'], ['vert', '&Vertical'], ['dup', 'Diagonal &up'], ['ddn', 'Dia&gonal down'], ['corner', '&From corner'], ['center', 'Fro&m center']].map(([k, l]) => ui.radio('fe-shade', l, style === k, () => { style = k; variant = 0; refreshVariants(); })));
      const gradTab = h('div', { class: 'col' },
        h('div', { class: 'row' },
          ui.group('Colors', h('div', { class: 'col' }, ui.radio('fe-mode', '&One color', false, () => setMode('one')), ui.radio('fe-mode', '&Two colors', true, () => setMode('two')), ui.radio('fe-mode', 'Pr&eset', false, () => setMode('preset')))),
          colorsArea),
        ui.group('Transparency', h('div', { class: 'col' }, ui.field('F&rom:', t1), ui.field('T&o:', t2))),
        h('div', { class: 'row' }, ui.group('Shading styles', shadingRadios), ui.group('Variants', variants)));
      /* texture */
      const texGrid = h('div', { class: 'tex-grid' });
      let texSel = null;
      TEXTURES.forEach((n) => {
        const b = h('button', { type: 'button', 'data-tip': n, 'aria-label': n, style: `background-image:url(${texture(n)});background-size:44px 44px` });
        b.addEventListener('click', async () => { L.$$('button', texGrid).forEach((x) => x.classList.remove('on')); b.classList.add('on'); texSel = n; texName.textContent = L.t(n); const id = await textureMedia(n); current = { t: 'img', media: id, tile: true }; paintSample(); });
        texGrid.appendChild(b);
      });
      const texName = h('div', { class: 'tp-note', text: ' ' });
      const otherTex = ui.button('&Other Texture...', async () => { const f = await L.pickFiles('image/*'); if (f[0]) { const id = L.media.add(f[0], f[0].name); current = { t: 'img', media: id, tile: true }; paintSample(); texName.textContent = f[0].name; } });
      const texTab = h('div', { class: 'col' }, h('label', { text: 'Texture:' }), texGrid, texName, otherTex);
      /* pattern */
      let pfg = '#000000', pbg = '#FFFFFF', pprst = initial && initial.t === 'patt' ? initial.prst : 'pct50';
      const pattGrid = h('div', { class: 'patt-grid' });
      const pattLabel = (k) => (L.geom.PATTERN_NAMES || {})[k] || k;   /* Office's name of a preset pattern */
      const pattName = h('div', { class: 'tp-note', text: pattLabel(pprst) });
      const drawPatt = () => {
        L.clear(pattGrid);
        Object.keys(L.render.PATTERNS).forEach((k) => {
          const b = h('button', { type: 'button', 'aria-label': pattLabel(k), 'data-tip': pattLabel(k), class: k === pprst ? 'on' : '', style: `background:${L.render.patternURL(k, L.model.resolveColor(pfg, d0), L.model.resolveColor(pbg, d0))}` });
          b.addEventListener('click', () => { pprst = k; pattName.textContent = L.t(pattLabel(k)); current = { t: 'patt', prst: k, fg: pfg, bg: pbg }; paintSample(); drawPatt(); });
          pattGrid.appendChild(b);
        });
      };
      drawPatt();
      const fgP = colorPick(pfg, (c) => { pfg = c; current = { t: 'patt', prst: pprst, fg: pfg, bg: pbg }; paintSample(); drawPatt(); });
      const bgP = colorPick(pbg, (c) => { pbg = c; current = { t: 'patt', prst: pprst, fg: pfg, bg: pbg }; paintSample(); drawPatt(); });
      const pattTab = h('div', { class: 'col' }, h('label', { text: 'Pattern:' }), pattGrid, pattName, h('div', { class: 'row' }, ui.field('&Foreground:', fgP), ui.field('&Background:', bgP)));
      /* picture */
      const picPrev = h('div', { class: 'preview-box', style: 'width:220px;height:150px' });
      const picTab = h('div', { class: 'col' }, ui.button('Select &Picture...', async () => { const f = await L.pickFiles('image/*'); if (!f[0]) return; const id = L.media.add(f[0], f[0].name); current = { t: 'img', media: id }; picPrev.style.background = `url(${L.media.url(id)}) center/contain no-repeat`; paintSample(); }), picPrev);
      const tabs = ui.tabs([{ label: '&Gradient', body: gradTab }, { label: 'Te&xture', body: texTab }, { label: '&Pattern', body: pattTab }, { label: 'P&icture', body: picTab }], initial && initial.t === 'patt' ? 2 : initial && initial.t === 'img' ? (initial.tile ? 1 : 3) : 0);
      setMode('two');
      if (initial && initial.t !== 'grad') { current = L.clone(initial); paintSample(); }
      const body = h('div', { class: 'row', style: 'flex-wrap:nowrap' }, tabs, h('div', { class: 'col', style: 'flex:none;padding-top:26px' }, h('label', { text: 'Sample:' }), sample));
      ui.dialog({ title: 'Fill Effects', body, width: 560, buttons: [{ label: 'OK', primary: true, onClick: () => { res = current; } }, { label: 'Cancel' }] }).done.then(() => resolve(res));
      void texSel;
    });
  };

  /* ---------- Format AutoShape / Picture / Text Box ---------- */
  function applyShapeChanges(shapes, ch, title) {
    E().commit(title, () => {
      for (const s of shapes) {
        const tgt = s.type === 'wordart' ? s.wa : s;
        if (ch.fill !== undefined) tgt.fill = L.clone(ch.fill);
        if (ch.alpha != null && tgt.fill && tgt.fill.t === 'solid') tgt.fill.a = ch.alpha;
        if (ch.alpha != null && tgt.fill && tgt.fill.t === 'grad') tgt.fill.stops.forEach((st) => { st.a = ch.alpha; });
        if (ch.lineColor !== undefined) tgt.line = ch.lineColor && ch.lineColor.t === 'none' ? { t: 'none', w: (tgt.line && tgt.line.w) || 0.75 } : Object.assign({ w: 0.75, dash: 'solid' }, tgt.line && tgt.line.t !== 'none' ? tgt.line : {}, { c: ch.lineColor.c, a: ch.lineColor.a });
        if (tgt.line && tgt.line.t !== 'none') {
          if (ch.lineW != null) tgt.line.w = ch.lineW;
          if (ch.dash) tgt.line.dash = ch.dash;
          if (ch.cmpd) tgt.line.cmpd = ch.cmpd;
          if (ch.head) tgt.line.head = ch.head === 'none' ? undefined : { type: ch.head, w: ch.headSz || 'med', len: ch.headSz || 'med' };
          if (ch.tail) tgt.line.tail = ch.tail === 'none' ? undefined : { type: ch.tail, w: ch.tailSz || 'med', len: ch.tailSz || 'med' };
          if (!ch.head && ch.headSz && tgt.line.head) tgt.line.head.w = tgt.line.head.len = ch.headSz;
          if (!ch.tail && ch.tailSz && tgt.line.tail) tgt.line.tail.w = tgt.line.tail.len = ch.tailSz;
        } else if (ch.lineW != null && (!tgt.line || tgt.line.t === 'none') && ch.lineColor === undefined && s.type !== 'image') { tgt.line = { c: 'tx1', w: ch.lineW, dash: ch.dash || 'solid' }; }
        if (shapes.length === 1) {
          const ob = { x: s.x, y: s.y, w: s.w, h: s.h };
          if (ch.w != null) s.w = Math.max(0, ch.w);
          if (ch.h != null) s.h = Math.max(0, ch.h);
          if (ch.x != null) s.x = ch.x;
          if (ch.y != null) s.y = ch.y;
          if (s.type === 'group') { if (ch.x != null || ch.y != null) M.translate(s, 0, 0); M.scaleGroup(s, ob, s); }
          if (s.type === 'table' && (ch.w != null || ch.h != null)) E().scaleTable(s, Object.assign(L.clone(s), ob));
        } else { if (ch.w != null) s.w = ch.w; if (ch.h != null) s.h = ch.h; }
        if (ch.rot != null && M.canRotate(s)) s.rot = ch.rot;
        if (ch.lockAspect != null) s.lockAspect = ch.lockAspect;
        if (s.type === 'image') {
          if (ch.crop) {
            const c = s.crop || { l: 0, t: 0, r: 0, b: 0 };
            const fullW = s.w / (1 - c.l - c.r), fullH = s.h / (1 - c.t - c.b);
            s.crop = ch.crop;
            s.w = fullW * (1 - ch.crop.l - ch.crop.r); s.h = fullH * (1 - ch.crop.t - ch.crop.b);
          }
          s.img = Object.assign({}, s.img || {}, ch.img || {});
          if (ch.imgMode != null) { if (ch.imgMode) s.img.mode = ch.imgMode; else delete s.img.mode; }
          if (ch.bright != null) s.img.bright = ch.bright;
          if (ch.contrast != null) s.img.contrast = ch.contrast;
        }
        if (s.tx) {
          if (ch.ins) s.tx.ins = ch.ins.map((v) => L.round(v, 2));
          if (ch.anchor) { const [a, c] = ch.anchor.split('|'); s.tx.anchor = a; s.tx.anchorCtr = !!c; }
          if (ch.wrap != null) s.tx.wrap = ch.wrap;
          if (ch.autofit) s.tx.autofit = ch.autofit;
          if (ch.vert != null) { if (ch.vert) s.tx.vert = 'vert'; else delete s.tx.vert; }
        }
        if (ch.alt != null) s.alt = ch.alt;
        const f = E().find(s.id);
        if (f && f.parent) E().fitGroup(f.parent);
      }
    });
  }

  D.formatShape = function (tabName) {
    const shapes = E().selected();
    if (!shapes.length) return;
    const s0 = shapes[0];
    const isPic = s0.type === 'image', isLine = s0.type === 'line', isWA = s0.type === 'wordart', isTable = s0.type === 'table';
    const title = isPic ? 'Format Picture' : isLine ? 'Format AutoShape' : isWA ? 'Format LettersArt' : s0.ph ? 'Format Placeholder' : s0.type === 'text' ? 'Format Text Box' : isTable ? 'Format Table' : 'Format AutoShape';
    const ch = {}; /* changes */
    const srcFill = isWA ? s0.wa.fill : s0.fill;
    const srcLine = isWA ? s0.wa.line : s0.line;
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const d = design(), crop = s0.crop || { l: 0, t: 0, r: 0, b: 0 };
      const fullW = s0.w / (1 - crop.l - crop.r), fullH = s0.h / (1 - crop.t - crop.b), tx = s0.tx;
      const colors = [
        ...(ui.schemeColors ? ui.schemeColors(d).map((color) => ({ value: color.ref, label: color.tip || color.ref, color: color.hex })) : []),
        ...L.color.STANDARD.map((color) => ({ value: color, label: ui.colorName(color), color })),
      ];
      reactFormatShape = { shapes, shape: s0, title, isPic, isLine, isWA, hasText: !!tx,
        fill: L.clone(srcFill || { t: 'none' }), line: L.clone(srcLine || { t: 'none' }), colors,
        lineWidth: srcLine && srcLine.w != null ? srcLine.w : 0.75, dash: srcLine && srcLine.dash || 'solid', compound: srcLine && srcLine.cmpd || 'sng',
        head: srcLine && srcLine.head && srcLine.head.type || 'none', tail: srcLine && srcLine.tail && srcLine.tail.type || 'none',
        headSize: srcLine && srcLine.head && srcLine.head.w || 'med', tailSize: srcLine && srcLine.tail && srcLine.tail.w || 'med',
        showArrows: isLine || (s0.type === 'shape' && L.geom.get(s0.geom).line) || s0.geom === 'custom', canRotate: shapes.every(M.canRotate),
        width: IN(s0.w), height: IN(s0.h), x: IN(s0.x), y: IN(s0.y), rotation: s0.rot || 0, lockAspect: !!s0.lockAspect,
        alpha: srcFill && srcFill.t === 'solid' && srcFill.a != null ? srcFill.a : 1,
        crop: { l: IN(crop.l * fullW), r: IN(crop.r * fullW), t: IN(crop.t * fullH), b: IN(crop.b * fullH) }, cropFullWidthPt: fullW, cropFullHeightPt: fullH,
        imageMode: s0.img && s0.img.mode || '', brightness: Math.round((s0.img && s0.img.bright || 0) * 100), contrast: Math.round((s0.img && s0.img.contrast || 0) * 100),
        margins: tx && tx.ins ? tx.ins.map(IN) : [0.1, 0.05, 0.1, 0.05], anchor: tx ? tx.anchor + (tx.anchorCtr ? '|c' : '') : 't',
        wrap: tx ? tx.wrap !== false : true, autofit: !!(tx && tx.autofit === 'shape'), verticalText: !!(tx && tx.vert), alt: s0.alt || '',
      };
      if (isPic && L.media.has(s0.media)) L.loadImage(L.media.url(s0.media)).then((image) => {
        if (!reactFormatShape || reactFormatShape.shape !== s0) return;
        reactFormatShape.original = { width: IN(image.naturalWidth * 0.75), height: IN(image.naturalHeight * 0.75) };
        L.bus.emit('pptx-format-shape-dialog');
      }).catch(() => {});
      L.bus.emit('pptx-format-shape-dialog');
      return;
    }
    /* colors & lines */
    const fp = fillPicker(srcFill, 'fill', (f) => { ch.fill = f; trans.set(f && f.a != null ? Math.round((1 - f.a) * 100) : 0); });
    const trans = ui.spin({ value: srcFill && srcFill.a != null && srcFill.t === 'solid' ? Math.round((1 - srcFill.a) * 100) : 0, min: 0, max: 100, step: 5, unit: ' %', dec: 0, onChange: (v) => { ch.alpha = 1 - v / 100; } });
    const lp = fillPicker(srcLine && srcLine.t !== 'none' ? srcLine : { t: 'none' }, 'line', (l) => { ch.lineColor = l; });
    const lnW = ui.spin({ value: srcLine && srcLine.w != null ? srcLine.w : 0.75, min: 0, max: 1584, step: 0.25, unit: ' pt', dec: 2, onChange: (v) => { ch.lineW = v; } });
    const dash = ui.select(DASHES, (srcLine && srcLine.dash) || 'solid', (v) => { ch.dash = v; });
    const cmpd = ui.select([['sng', 'Single'], ['dbl', 'Double'], ['thickThin', 'Thick Thin'], ['thinThick', 'Thin Thick'], ['tri', 'Triple']], (srcLine && srcLine.cmpd) || 'sng', (v) => { ch.cmpd = v; });
    const aB = ui.select(ARROWS, (srcLine && srcLine.head && srcLine.head.type) || 'none', (v) => { ch.head = v; });
    const aE = ui.select(ARROWS, (srcLine && srcLine.tail && srcLine.tail.type) || 'none', (v) => { ch.tail = v; });
    const sB = ui.select(SIZES, (srcLine && srcLine.head && srcLine.head.w) || 'med', (v) => { ch.headSz = v; });
    const sE = ui.select(SIZES, (srcLine && srcLine.tail && srcLine.tail.w) || 'med', (v) => { ch.tailSz = v; });
    const colorsTab = h('div', { class: 'col' },
      !isLine && !isPic ? ui.group('Fill', ui.field('C&olor:', fp), ui.field('&Transparency:', trans)) : null,
      isPic ? ui.group('Fill', ui.field('C&olor:', fp)) : null,
      ui.group('Line', h('div', { class: 'row' }, h('div', { class: 'col' }, ui.field('Co&lor:', lp), ui.field('&Dashed:', dash)), h('div', { class: 'col' }, ui.field('&Style:', cmpd), ui.field('&Weight:', lnW)))),
      isLine || (s0.type === 'shape' && L.geom.get(s0.geom).line) || (s0.geom === 'custom') ? ui.group('Arrows', h('div', { class: 'row' }, h('div', { class: 'col' }, ui.field('&Begin style:', aB), ui.field('Begin si&ze:', sB)), h('div', { class: 'col' }, ui.field('&End style:', aE), ui.field('E&nd size:', sE)))) : null);
    /* size */
    const hS = ui.spin({ value: IN(s0.h), min: 0, max: 56, step: 0.01, unit: '"', onChange: (v) => { ch.h = PT(v); if (lock.input.checked && s0.h) { ch.w = s0.w * (PT(v) / s0.h); wS.set(IN(ch.w)); } } });
    const wS = ui.spin({ value: IN(s0.w), min: 0, max: 56, step: 0.01, unit: '"', onChange: (v) => { ch.w = PT(v); if (lock.input.checked && s0.w) { ch.h = s0.h * (PT(v) / s0.w); hS.set(IN(ch.h)); } } });
    const rS = ui.spin({ value: s0.rot || 0, min: -360, max: 360, step: 1, unit: '°', dec: 1, onChange: (v) => { ch.rot = ((v % 360) + 360) % 360; } });
    rS.setDisabled(!shapes.every(M.canRotate));
    const lock = ui.check('Lock &aspect ratio', !!s0.lockAspect, (v) => { ch.lockAspect = v; });
    let natW = 0, natH = 0;
    const origLbl = h('div', { class: 'tp-note' });
    if (isPic && L.media.has(s0.media)) L.loadImage(L.media.url(s0.media)).then((im) => { natW = im.naturalWidth * 0.75; natH = im.naturalHeight * 0.75; origLbl.textContent = L.t('Original size: Height {0}"  Width {1}"', IN(natH), IN(natW)); }).catch(() => {});
    const sizeTab = h('div', { class: 'col' },
      ui.group('Size and rotate', h('div', { class: 'row' }, ui.field('H&eight:', hS), ui.field('Wi&dth:', wS)), ui.field('Rota&tion:', rS)),
      ui.group('Scale', lock),
      isPic ? h('div', null, origLbl, ui.button('&Reset', () => { if (natW) { ch.w = natW; ch.h = natH; wS.set(IN(natW)); hS.set(IN(natH)); ch.crop = { l: 0, t: 0, r: 0, b: 0 }; ch.img = {}; } })) : null);
    /* position */
    const xS = ui.spin({ value: IN(s0.x), min: -56, max: 56, step: 0.01, unit: '"', onChange: (v) => { ch.x = PT(v); } });
    const yS = ui.spin({ value: IN(s0.y), min: -56, max: 56, step: 0.01, unit: '"', onChange: (v) => { ch.y = PT(v); } });
    const posTab = h('div', { class: 'col' }, ui.group('Position on slide', ui.field('&Horizontal:', xS), ui.field('&Vertical:', yS), h('div', { class: 'tp-note', text: 'From: Top Left Corner' })));
    /* picture */
    let picTab = null;
    if (isPic) {
      const c = s0.crop || { l: 0, t: 0, r: 0, b: 0 };
      const fullW = s0.w / (1 - c.l - c.r), fullH = s0.h / (1 - c.t - c.b);
      const cr = (k, full) => ui.spin({ value: IN(c[k] * full), min: -20, max: 20, step: 0.01, unit: '"', onChange: (v) => { ch.crop = Object.assign({}, ch.crop || c); ch.crop[k] = PT(v) / full; } });
      const im = s0.img || {};
      const mode = ui.select([['', 'Automatic'], ['gray', 'Grayscale'], ['bw', 'Black & White'], ['wash', 'Washout']], im.mode || '', (v) => { ch.imgMode = v; });
      const br = h('input', { type: 'range', min: -100, max: 100, value: Math.round((im.bright || 0) * 100), id: 'fp-bright' });
      const co = h('input', { type: 'range', min: -100, max: 100, value: Math.round((im.contrast || 0) * 100), id: 'fp-contrast' });
      br.addEventListener('input', () => { ch.bright = br.value / 100; });
      co.addEventListener('input', () => { ch.contrast = co.value / 100; });
      picTab = h('div', { class: 'col' },
        ui.group('Crop from', h('div', { class: 'row' }, h('div', { class: 'col' }, ui.field('&Left:', cr('l', fullW)), ui.field('&Right:', cr('r', fullW))), h('div', { class: 'col' }, ui.field('&Top:', cr('t', fullH)), ui.field('&Bottom:', cr('b', fullH))))),
        ui.group('Image control', ui.field('&Color:', mode), ui.field('Brig&htness:', br), ui.field('Co&ntrast:', co)));
    }
    /* text box */
    let tbTab = null;
    if (s0.tx) {
      const tx = s0.tx;
      const ins = tx.ins || [7.2, 3.6, 7.2, 3.6];
      const insSp = (i) => ui.spin({ value: IN(ins[i]), min: 0, max: 20, step: 0.01, unit: '"', onChange: (v) => { ch.ins = (ch.ins || ins.slice()); ch.ins[i] = PT(v); } });
      const anchor = ui.select([['t', 'Top'], ['ctr', 'Middle'], ['b', 'Bottom'], ['t|c', 'Top Centered'], ['ctr|c', 'Middle Centered'], ['b|c', 'Bottom Centered']], tx.anchor + (tx.anchorCtr ? '|c' : ''), (v) => { ch.anchor = v; });
      tbTab = h('div', { class: 'col' },
        ui.field('Text &anchor point:', anchor, { cls: 'wide' }),
        ui.group('Internal margin', h('div', { class: 'row' }, h('div', { class: 'col' }, ui.field('&Left:', insSp(0)), ui.field('&Right:', insSp(2))), h('div', { class: 'col' }, ui.field('&Top:', insSp(1)), ui.field('&Bottom:', insSp(3))))),
        ui.check('&Word wrap text in AutoShape', tx.wrap !== false, (v) => { ch.wrap = v; }),
        ui.check('Resize AutoShape to &fit text', tx.autofit === 'shape', (v) => { ch.autofit = v ? 'shape' : 'none'; }),
        ui.check('R&otate text within AutoShape by 90°', !!tx.vert, (v) => { ch.vert = v; }));
    }
    const alt = h('textarea', { id: 'fs-alt', rows: 4, style: 'width:100%' });
    alt.value = s0.alt || '';
    alt.addEventListener('input', () => { ch.alt = alt.value; });
    const webTab = h('div', { class: 'col' }, h('label', { for: 'fs-alt', text: 'Alternative text:' }), alt, h('div', { class: 'tp-note', text: 'Screen readers and web pages use alternative text to describe the object.' }));
    const tabs = [{ label: 'Colors and Lines', body: colorsTab }, { label: 'Size', body: sizeTab }, { label: 'Position', body: posTab }];
    if (picTab) tabs.push({ label: 'Picture', body: picTab });
    if (tbTab) tabs.push({ label: 'Text Box', body: tbTab });
    tabs.push({ label: 'Web', body: webTab });
    const ti = Math.max(0, tabs.findIndex((t) => t.label === tabName));
    const tabsEl = ui.tabs(tabs, ti);
    ui.dialog({ title, body: tabsEl, width: 470, buttons: [{ label: 'OK', primary: true, onClick: () => applyShapeChanges(shapes, ch, title) }, { label: 'Cancel' }] });
  };

  /* ---------- Background ---------- */
  D.background = function () {
    const slide = E().slide();
    if (!slide) return;
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const d = design();
      const colors = [
        ...(ui.schemeColors ? ui.schemeColors(d).map((color) => ({ value: color.ref, label: color.tip || color.ref, color: color.hex })) : []),
        ...L.color.STANDARD.map((color) => ({ value: color, label: ui.colorName(color), color })),
      ];
      reactBackground = { slide, fill: L.clone(slide.bg || d.bg), omit: !!slide.hideMaster, colors };
      L.bus.emit('pptx-background-dialog');
      return;
    }
    const d = design();
    let fill = L.clone(slide.bg || d.bg);
    let omit = !!slide.hideMaster;
    const prev = h('div', { class: 'preview-box', style: 'width:220px;height:150px;position:relative' });
    const paint = () => {
      L.clear(prev);
      const tmp = Object.assign({}, slide, { bg: fill, hideMaster: omit });
      prev.appendChild(L.render.thumb(L.pres, tmp, 218));
    };
    const fp = fillPicker(fill, 'fill', (f) => { fill = f && f.t === 'none' ? { t: 'solid', c: 'bg1', a: 1 } : f; paint(); });
    const om = ui.check('Omit background &graphics from master', omit, (v) => { omit = v; paint(); });
    paint();
    const apply = (all) => {
      L.hist.push(L.N('Background'));
      const targets = all ? L.pres.slides : (L.app.selectedSlides && L.app.selectedSlides().length > 1 ? L.app.selectedSlides() : [slide]);
      if (E().view === 'master') { const ms = E().masterSlide; ms.bg = fill; E().syncMaster(); }
      else if (all) { for (const s of L.pres.slides) { delete s.bg; s.hideMaster = omit || undefined; } for (const id in L.pres.designs) { if (L.pres.slides.some((s) => s.design === id)) L.pres.designs[id].bg = L.clone(fill); } }
      else for (const s of targets) { s.bg = L.clone(fill); s.hideMaster = omit || undefined; }
      E().render(); L.bus.emit('slides-changed');
    };
    ui.dialog({ title: 'Background', body: h('div', { class: 'row' }, h('div', { class: 'col' }, h('label', { text: 'Background fill' }), prev, fp, om)), width: 380, buttons: [{ label: 'Apply to &All', onClick: () => apply(true) }, { label: '&Apply', primary: true, onClick: () => apply(false) }, { label: 'Cancel' }] });
  };

  /* ---------- Page Setup ---------- */
  D.pageSetup = function () {
    const p = L.pres;
    let cur = M.SLIDE_SIZES.find((s) => Math.abs(s.w - p.W) < 0.6 && Math.abs(s.h - p.H) < 0.6 && s.key !== 'letter' && s.key !== 'overhead') || M.SLIDE_SIZES[M.SLIDE_SIZES.length - 1];
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const defaultSize = L.app.opts.slideSize || M.DEFAULT_SIZE;
      reactPageSetup = { sizes: M.SLIDE_SIZES.map((size) => ({ key: size.key, name: size.name, width: IN(size.w), height: IN(size.h) })), preset: cur.key, width: IN(p.W), height: IN(p.H), firstNum: p.firstNum || 1, portrait: p.H > p.W, asDefault: !!(defaultSize && Math.abs(defaultSize.w - p.W) < 0.6 && Math.abs(defaultSize.h - p.H) < 0.6) };
      L.bus.emit('pptx-page-setup-dialog');
      return;
    }
    const w = ui.spin({ value: IN(p.W), min: 1, max: 56, step: 0.1, unit: '"' }), hh = ui.spin({ value: IN(p.H), min: 1, max: 56, step: 0.1, unit: '"' });
    const sel = ui.select(M.SLIDE_SIZES.map((s) => [s.key, s.name]), cur.key, (k) => { const s = M.SLIDE_SIZES.find((x) => x.key === k); if (k !== 'custom') { w.set(IN(s.w)); hh.set(IN(s.h)); } });
    const first = ui.spin({ value: p.firstNum || 1, min: 0, max: 9999, step: 1, dec: 0 });
    let portrait = p.H > p.W;
    const dflt = L.app.opts.slideSize || M.DEFAULT_SIZE;
    const asDefault = ui.check('&Use for new presentations', !!(dflt && Math.abs(dflt.w - p.W) < 0.6 && Math.abs(dflt.h - p.H) < 0.6));
    const body = h('div', { class: 'row' },
      h('div', { class: 'col' }, ui.field('&Slides sized for:', sel, { cls: 'wide' }), ui.field('&Width:', w, { cls: 'wide' }), ui.field('H&eight:', hh, { cls: 'wide' }), ui.field('&Number slides from:', first, { cls: 'wide' }), asDefault),
      ui.group('Orientation', h('div', { class: 'col' }, h('b', { text: 'Slides' }), ui.radio('ps-or', '&Portrait', portrait, () => { portrait = true; if (w.get() > hh.get()) { const t = w.get(); w.set(hh.get()); hh.set(t); } }), ui.radio('ps-or', '&Landscape', !portrait, () => { portrait = false; if (hh.get() > w.get()) { const t = w.get(); w.set(hh.get()); hh.set(t); } }))));
    ui.dialog({ title: 'Page Setup', body, width: 480, buttons: [{ label: 'OK', primary: true, onClick: () => {
      const preset = M.SLIDE_SIZES.find((x) => x.key === sel.value);
      /* use exact preset points (inch spinners round to 0.01") */
      const W = preset && preset.key !== 'custom' && Math.abs(IN(preset.w) - w.get()) < 0.011 && Math.abs(IN(preset.h) - hh.get()) < 0.011 ? preset.w : PT(w.get());
      const H = preset && preset.key !== 'custom' && Math.abs(IN(preset.w) - w.get()) < 0.011 && Math.abs(IN(preset.h) - hh.get()) < 0.011 ? preset.h : PT(hh.get());
      const W2 = portrait ? Math.min(W, H) : Math.max(W, H), H2 = portrait ? Math.max(W, H) : Math.min(W, H);
      L.app.resizeSlides(W2, H2);
      L.pres.firstNum = first.get();
      if (asDefault.input.checked) L.app.opts.slideSize = { w: W2, h: H2 };
      else if (dflt && Math.abs(dflt.w - W2) < 0.6 && Math.abs(dflt.h - H2) < 0.6) delete L.app.opts.slideSize;
      L.app.saveOpts();
      E().render(); L.bus.emit('slides-changed');
    } }, { label: 'Cancel' }] });
  };

  /* ---------- Header and Footer ---------- */
  D.headerFooter = function () {
    const slide = E().slide();
    const hf = Object.assign({}, L.pres.hf, slide && slide.hf);
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactHeaderFooter = {
        date: !!hf.dt, automatic: hf.dtAuto !== false, format: hf.dtFmt || 'datetime1',
        fixedText: hf.dtText || L.fmtDate(new Date(), 'datetime1'), number: !!hf.num,
        footer: !!hf.ftr, footerText: hf.ftrText || '', notOnTitle: !!hf.notOnTitle,
        formats: L.DATE_FORMATS.map((id) => ({ id, label: L.fmtDate(new Date(), id) })),
      };
      L.bus.emit('pptx-header-footer-dialog');
      return;
    }
    const dt = ui.check('&Date and time', !!hf.dt, () => sync());
    const auto = ui.radio('hf-dt', 'Update &automatically', hf.dtAuto !== false, () => sync());
    const fmt = ui.select(L.DATE_FORMATS.map((f) => [f, L.fmtDate(new Date(), f)]), hf.dtFmt || 'datetime1', null, { translate: 'no' });
    const fixed = ui.radio('hf-dt', 'Fi&xed', hf.dtAuto === false, () => sync());
    const fixedIn = h('input', { type: 'text', id: 'hf-fixed', value: hf.dtText || L.fmtDate(new Date(), 'datetime1'), style: 'width:220px' });
    const num = ui.check('Slide &number', !!hf.num);
    const ftr = ui.check('&Footer', !!hf.ftr, () => sync());
    const ftrIn = h('input', { type: 'text', id: 'hf-ftr', value: hf.ftrText || '', style: 'width:260px' });
    const nt = ui.check('Don\'t show on title &slide', !!hf.notOnTitle);
    const sync = () => { const on = dt.input.checked; auto.input.disabled = fixed.input.disabled = !on; fmt.disabled = !on || !auto.input.checked; fixedIn.disabled = !on || !fixed.input.checked; ftrIn.disabled = !ftr.input.checked; };
    sync();
    const read = () => ({ dt: dt.input.checked, dtAuto: auto.input.checked, dtFmt: fmt.value, dtText: fixedIn.value, num: num.input.checked, ftr: ftr.input.checked, ftrText: ftrIn.value, notOnTitle: nt.input.checked });
    const body = ui.tabs([{ label: 'Slide', body: h('div', { class: 'col' }, h('b', { text: 'Include on slide' }), dt, h('div', { class: 'col', style: 'margin-left:20px' }, auto, fmt, fixed, fixedIn), num, ftr, h('div', { style: 'margin-left:20px' }, ftrIn), nt) }]);
    ui.dialog({
      title: 'Header and Footer', body, width: 420, buttons: [
        { label: 'Apply to &All', primary: true, onClick: () => { L.hist.push('Header and Footer'); L.pres.hf = read(); for (const s of L.pres.slides) delete s.hf; E().render(); L.bus.emit('slides-changed'); } },
        { label: '&Apply', onClick: () => { L.hist.push('Header and Footer'); if (slide) slide.hf = read(); E().render(); L.bus.emit('slides-changed'); } },
        { label: 'Cancel' }],
    });
  };

  /* ---------- Find / Replace ---------- */
  const findState = { what: '', with: '', mc: false, ww: false, pos: null };
  function findNext(fs) {
    const slides = L.pres.slides;
    if (!slides.length || !fs.what) return false;
    const flags = fs.mc ? 'g' : 'gi';
    const esc = fs.what.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(fs.ww ? `\\b${esc}\\b` : esc, flags);
    const startSlide = E().idx;
    let after = fs.pos;
    if (L.te.active()) { const s = L.te.getSel(); const t = L.te.target(); if (s && t) after = { slide: E().idx, sid: t.sh.id, cell: L.te.state.cell, p: Math.max(s.p0, s.p1), o: Math.max(s.o1, s.o0) }; }
    for (let k = 0; k <= slides.length; k++) {
      const si = (startSlide + k) % slides.length;
      const sl = slides[si];
      const targets = [];
      M.walk(sl.shapes, (sh) => {
        if (sh.tx) targets.push({ sh, tx: sh.tx, cell: null });
        if (sh.tbl) sh.tbl.rows.forEach((r, ri) => r.cells.forEach((c, ci) => targets.push({ sh, tx: c.tx, cell: [ri, ci] })));
      });
      let passed = !after || after.slide !== si || k > 0 && k === slides.length;
      for (const t of targets) {
        if (!passed) { if (t.sh.id === after.sid && String(t.cell) === String(after.cell)) passed = 'here'; else continue; }
        for (let pi = 0; pi < t.tx.ps.length; pi++) {
          const text = L.txt.paraText(t.tx.ps[pi]);
          re.lastIndex = 0;
          let m;
          while ((m = re.exec(text))) {
            if (passed === 'here' && (pi < after.p || (pi === after.p && m.index < after.o))) { if (!m[0].length) re.lastIndex++; continue; }
            fs.pos = { slide: si, sid: t.sh.id, cell: t.cell, p: pi, o: m.index + m[0].length };
            if (L.app.view !== 'normal') L.app.setView('normal');
            if (E().idx !== si) E().goto(si);
            L.te.begin(t.sh.id, { cell: t.cell || undefined, sel: { p0: pi, o0: m.index, p1: pi, o1: m.index + m[0].length } });
            return true;
          }
        }
        passed = true;
      }
      after = null;
    }
    return false;
  }
  function findNotFound() { ui.msg(L.t('{0} has finished searching the presentation. The search item wasn\'t found.', L.APP), { icon: 'info' }); }
  function replaceOne() {
    if (L.te.active()) {
      const sel = L.te.selectedText();
      const match = findState.mc ? sel === findState.what : sel.toLowerCase() === findState.what.toLowerCase();
      if (match && sel) { L.hist.push(L.N('Replace')); L.te.insertText(findState.with); }
    }
    if (!findNext(findState)) findNotFound();
  }
  function replaceAll() {
    if (!findState.what) return 0;
    if (L.te.active()) L.te.end();
    const esc = findState.what.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(findState.ww ? `\\b${esc}\\b` : esc, findState.mc ? 'g' : 'gi');
    let n = 0;
    L.hist.push('Replace All');
    const fix = (tx) => { for (const p of tx.ps) for (const r of p.rs) { if (r.fld) continue; r.t = r.t.replace(re, () => { n++; return findState.with; }); } };
    for (const s of L.pres.slides) M.walk(s.shapes, (sh) => { if (sh.tx) fix(sh.tx); if (sh.tbl) sh.tbl.rows.forEach((r) => r.cells.forEach((c) => fix(c.tx))); if (sh.wa) sh.wa.text = sh.wa.text.replace(re, () => { n++; return findState.with; }); });
    if (!n) L.hist.undo.pop();
    E().render(); L.bus.emit('slides-changed');
    return n;
  }
  D.reactFindCatalog = () => reactFind;
  D.reactFindAction = (action, values) => {
    if (!reactFind) return false;
    if (action === 'close') { reactFind = null; L.bus.emit('pptx-find-dialog'); return true; }
    if (action === 'switch') { reactFind.replace = true; L.bus.emit('pptx-find-dialog'); return true; }
    if (action === 'update' && values && typeof values === 'object' && ['what', 'with', 'mc', 'ww'].every((key) => key in values)) {
      if (typeof values.what !== 'string' || typeof values.with !== 'string' || typeof values.mc !== 'boolean' || typeof values.ww !== 'boolean') return false;
      Object.assign(reactFind, values); return true;
    }
    Object.assign(findState, { what: reactFind.what, with: reactFind.with, mc: reactFind.mc, ww: reactFind.ww });
    if (action === 'find-next') { if (!findNext(findState)) findNotFound(); return true; }
    if (action === 'replace') { replaceOne(); return true; }
    if (action === 'replace-all') {
      const count = replaceAll();
      ui.msg(count > 1 ? L.t('{0} has finished searching the presentation and made {1} replacements.', L.APP, count) : count ? L.t('{0} has finished searching the presentation and made 1 replacement.', L.APP) : 'The search item wasn\'t found.', { icon: 'info' });
      return true;
    }
    return false;
  };
  D.reactPageSetupCatalog = () => reactPageSetup;
  D.reactPageSetupAction = (action, values) => {
    if (!reactPageSetup) return false;
    if (action === 'close') { reactPageSetup = null; L.bus.emit('pptx-page-setup-dialog'); return true; }
    if (action !== 'apply' || !values || !M.SLIDE_SIZES.some((size) => size.key === values.preset) || !Number.isFinite(values.width) || values.width < 1 || values.width > 56 || !Number.isFinite(values.height) || values.height < 1 || values.height > 56 || !Number.isInteger(values.firstNum) || values.firstNum < 0 || values.firstNum > 9999 || typeof values.portrait !== 'boolean' || typeof values.asDefault !== 'boolean') return false;
    const preset = M.SLIDE_SIZES.find((size) => size.key === values.preset);
    const exact = preset && preset.key !== 'custom' && Math.abs(IN(preset.w) - values.width) < 0.011 && Math.abs(IN(preset.h) - values.height) < 0.011;
    const width = exact ? preset.w : PT(values.width), height = exact ? preset.h : PT(values.height);
    const slideWidth = values.portrait ? Math.min(width, height) : Math.max(width, height), slideHeight = values.portrait ? Math.max(width, height) : Math.min(width, height);
    L.app.resizeSlides(slideWidth, slideHeight);
    L.pres.firstNum = values.firstNum;
    if (values.asDefault) L.app.opts.slideSize = { w: slideWidth, h: slideHeight };
    else if (L.app.opts.slideSize && Math.abs(L.app.opts.slideSize.w - slideWidth) < 0.6 && Math.abs(L.app.opts.slideSize.h - slideHeight) < 0.6) delete L.app.opts.slideSize;
    L.app.saveOpts(); E().render(); L.bus.emit('slides-changed');
    reactPageSetup = null; L.bus.emit('pptx-page-setup-dialog');
    return true;
  };
  D.reactHyperlinkCatalog = () => reactHyperlink;
  D.reactHyperlinkAction = (action, values) => {
    if (!reactHyperlink) return false;
    if (action === 'close') { reactHyperlink = null; reactHyperlinkApply = null; L.bus.emit('pptx-hyperlink-dialog'); return true; }
    if (action === 'remove') {
      if (!reactHyperlink.existing) return false;
      reactHyperlinkApply?.(null, reactHyperlink.display);
      reactHyperlink = null; reactHyperlinkApply = null; L.bus.emit('pptx-hyperlink-dialog'); return true;
    }
    if (action !== 'apply' || !values || !['web', 'place', 'mail'].includes(values.kind) || typeof values.display !== 'string' || typeof values.address !== 'string' || typeof values.tip !== 'string' || typeof values.email !== 'string' || typeof values.subject !== 'string' || typeof values.place !== 'string') return false;
    let link = null;
    if (values.kind === 'web') { const url = values.address.trim(); if (url && url !== 'https://') link = { url: /^[a-z]+:/i.test(url) ? url : 'https://' + url }; if (link && values.tip) link.tip = values.tip; }
    if (values.kind === 'mail' && values.email.trim()) link = { url: 'mailto:' + values.email.trim() + (values.subject ? '?subject=' + encodeURIComponent(values.subject) : '') };
    if (values.kind === 'place') link = values.place.startsWith('s:') ? { slide: values.place.slice(2) } : { action: values.place.slice(2) };
    if (link) reactHyperlinkApply?.(link, values.display);
    reactHyperlink = null; reactHyperlinkApply = null; L.bus.emit('pptx-hyperlink-dialog');
    return true;
  };
  D.reactWordartTextCatalog = () => reactWordartText;
  D.reactWordartTextAction = (action, values) => {
    if (!reactWordartText) return false;
    if (action === 'close') { reactWordartText = null; reactWordartTextCallback = null; L.bus.emit('pptx-wordart-text-dialog'); return true; }
    if (action !== 'apply' || !values || typeof values.text !== 'string' || !L.FONT_LIST.includes(values.font) || !L.SIZE_LIST.includes(values.size) || typeof values.b !== 'boolean' || typeof values.i !== 'boolean') return false;
    reactWordartTextCallback?.({ text: values.text, font: values.font, b: values.b, i: values.i, size: values.size });
    reactWordartText = null; reactWordartTextCallback = null; L.bus.emit('pptx-wordart-text-dialog');
    return true;
  };
  D.reactWordartGalleryCatalog = () => reactWordartGallery;
  D.reactWordartGalleryAction = (action, index) => {
    if (!reactWordartGallery) return false;
    if (action === 'close') { reactWordartGallery = null; reactWordartGalleryCallback = null; L.bus.emit('pptx-wordart-gallery-dialog'); return true; }
    if (action === 'select' && Number.isInteger(index) && index >= 0 && index < reactWordartGallery.previews.length) { reactWordartGallery.selected = index; return true; }
    if (action === 'apply' && Number.isInteger(reactWordartGallery.selected) && reactWordartGallery.selected >= 0 && reactWordartGallery.selected < reactWordartGallery.previews.length) {
      reactWordartGalleryCallback?.(reactWordartGallery.selected);
      reactWordartGallery = null; reactWordartGalleryCallback = null; L.bus.emit('pptx-wordart-gallery-dialog');
      return true;
    }
    return false;
  };
  D.reactPhotoAlbumCatalog = () => reactPhotoAlbum ? { layouts: [['fit', 'Fit to slide'], ['1', '1 picture'], ['2', '2 pictures'], ['4', '4 pictures'], ['1t', '1 picture with title'], ['2t', '2 pictures with title'], ['4t', '4 pictures with title']], frames: [['rect', 'Rectangle'], ['roundRect', 'Rounded Rectangle'], ['simple', 'Simple Frame, White'], ['black', 'Simple Frame, Black'], ['shadow', 'Soft Shadow']] } : null;
  D.reactPhotoAlbumAction = async (action, photos, layout, frame, captions) => {
    if (!reactPhotoAlbum) return false;
    if (action === 'close') { reactPhotoAlbum = false; L.bus.emit('pptx-photo-album-dialog'); return true; }
    const layouts = ['fit', '1', '2', '4', '1t', '2t', '4t'], frames = ['rect', 'roundRect', 'simple', 'black', 'shadow'];
    if (action !== 'create' || !Array.isArray(photos) || !photos.length || !photos.every((photo) => photo && photo.file instanceof File && typeof photo.name === 'string' && photo.file.type.startsWith('image/') && typeof photo.url === 'string' && photo.url.startsWith('blob:')) || !layouts.includes(layout) || !frames.includes(frame) || typeof captions !== 'boolean') return false;
    L.ui.busy(true, L.t('Creating Photo Album...'));
    try { await L.app.createPhotoAlbum(photos, layout, frame, captions); }
    catch (error) { L.ui.msg(L.t('The photo album could not be created: {0}', error?.message || String(error)), { icon: 'error' }); return false; }
    finally { L.ui.busy(false); }
    reactPhotoAlbum = false; L.bus.emit('pptx-photo-album-dialog');
    return true;
  };
  D.reactEffectOptionsCatalog = () => reactEffectOptions;
  D.reactEffectOptionsAction = (action, values) => {
    if (!reactEffectOptions) return false;
    if (action === 'close') { reactEffectOptions = null; reactEffectOptionsCallback = null; L.bus.emit('pptx-effect-options-dialog'); return true; }
    if (action !== 'apply' || !values || !['click', 'with', 'after'].includes(values.start) || !Number.isFinite(values.delay) || values.delay < 0 || values.delay > 60000 || !reactEffectOptions.speeds.some((speed) => speed.value === values.dur) || ![1, 2, 3, 4, 5, 10].includes(values.repeat) || !['all', 'para'].includes(values.by)) return false;
    if (reactEffectOptions.directions.length && !reactEffectOptions.directions.some(([value]) => value === values.dir)) return false;
    if (reactEffectOptions.amountEnabled && (!Number.isInteger(values.amount) || values.amount < 10 || values.amount > 1440)) return false;
    const result = { dir: reactEffectOptions.dir, amount: reactEffectOptions.amount, start: values.start, delay: values.delay, dur: values.dur, repeat: values.repeat, by: values.by };
    if (reactEffectOptions.directions.length) result.dir = values.dir;
    if (reactEffectOptions.amountEnabled) result.amount = values.amount;
    reactEffectOptionsCallback?.(result);
    reactEffectOptions = null; reactEffectOptionsCallback = null; L.bus.emit('pptx-effect-options-dialog');
    return true;
  };
  D.reactSlideFinderCatalog = () => reactSlideFinder ? {
    name: reactSlideFinder.name, keep: reactSlideFinder.keep,
    slides: reactSlideFinder.source.slides.map((slide, index) => ({ index, title: M.slideTitle(slide) || L.t('Slide'), thumb: L.render.thumb(reactSlideFinder.source, slide, 124).outerHTML })),
    selected: [...reactSlideFinder.selected],
  } : null;
  D.reactSlideFinderAction = (action, index, value) => {
    if (!reactSlideFinder) return false;
    if (action === 'close') { reactSlideFinder = null; L.bus.emit('pptx-slide-finder-dialog'); return true; }
    if (action === 'keep' && typeof value === 'boolean') { reactSlideFinder.keep = value; return true; }
    if (action === 'select' && Number.isInteger(index) && index >= 0 && index < reactSlideFinder.source.slides.length && typeof value === 'boolean') {
      if (value) reactSlideFinder.selected.add(index); else reactSlideFinder.selected.delete(index);
      return true;
    }
    if (action !== 'insert' && action !== 'insert-all') return false;
    const indices = action === 'insert-all' ? reactSlideFinder.source.slides.map((_, i) => i) : [...reactSlideFinder.selected].sort((a, b) => a - b);
    if (!indices.length) return false;
    try { L.app.insertSlidesFrom(reactSlideFinder.source, indices, reactSlideFinder.keep); }
    catch (error) { L.ui.msg(L.t('The selected slides could not be inserted: {0}', error?.message || String(error)), { icon: 'error' }); return false; }
    reactSlideFinder = null; L.bus.emit('pptx-slide-finder-dialog');
    return true;
  };
  D.reactOptionsCatalog = () => reactOptions;
  D.reactOptionsConfirmCatalog = () => reactOptionsConfirm ? { target: reactOptionsConfirm } : null;
  D.reactOptionsAction = (action, values) => {
    if (!reactOptions) return false;
    if (action === 'close') { L.app.loadOpts(); reactOptions = null; L.bus.emit('pptx-options-dialog'); return true; }
    if (action !== 'apply' || !values || typeof values.user !== 'string' || values.user.length > 255 || !Number.isInteger(values.undo) || values.undo < 3 || values.undo > 150 || typeof values.password !== 'string' || !reactOptions.optionKeys.every((key) => typeof values.options?.[key] === 'boolean') || (values.look && !reactOptions.looks.some(([id]) => id === values.look)) || (values.lang && !reactOptions.languages.some(([id]) => id === values.lang))) return false;
    Object.assign(L.app.opts, values.options);
    L.hist.limit = values.undo;
    L.pres.props.author = values.user;
    L.app.saveOpts();
    L.app.applyOpts();
    const VO = window.VO;
    if (VO?.look && values.look) { if (values.look !== VO.look.get()) VO.look.set(values.look); }
    if (VO?.lang && values.lang) { if (values.lang !== VO.lang.get()) { VO.lang.set(values.lang); L.ui.toast('The new language is used in the windows you open from now on.'); } }
    const nextPassword = values.password;
    const oldPassword = L.pres.password || '';
    reactOptions = null;
    L.bus.emit('pptx-options-dialog');
    if (nextPassword === oldPassword) return true;
    if (!nextPassword) { delete L.pres.password; L.hist.dirty = true; return true; }
    reactOptionsConfirm = nextPassword;
    L.bus.emit('pptx-options-confirm-dialog');
    return true;
  };
  D.reactOptionsConfirmAction = (action, value) => {
    if (!reactOptionsConfirm) return false;
    if (action === 'close') { reactOptionsConfirm = null; L.bus.emit('pptx-options-confirm-dialog'); return true; }
    if (action !== 'confirm' || typeof value !== 'string') return false;
    if (value !== reactOptionsConfirm) { L.ui.msg('Confirmation password is not identical.', { icon: 'warn' }); return false; }
    L.pres.password = reactOptionsConfirm;
    L.hist.dirty = true;
    reactOptionsConfirm = null;
    L.bus.emit('pptx-options-confirm-dialog');
    return true;
  };
  const reactBackgroundCatalog = () => reactBackground ? {
    fill: L.clone(reactBackground.fill), omit: reactBackground.omit,
    swatch: reactBackground.fill && reactBackground.fill.t !== 'none' ? fillCSS(reactBackground.fill) : 'repeating-linear-gradient(45deg,#fff 0 4px,#eee 4px 8px)',
    colors: reactBackground.colors,
    preview: L.render.thumb(L.pres, Object.assign({}, reactBackground.slide, { bg: reactBackground.fill, hideMaster: reactBackground.omit }), 218).outerHTML,
  } : null;
  D.reactFillEffectsCatalog = () => reactFillEffects ? {
    initial: L.clone(reactFillEffects.initial),
    presets: GRAD_PRESETS.map(([name, stops]) => ({ name, stops: L.clone(stops) })),
    textures: TEXTURES.map((name) => ({ name, src: texture(name) })),
    patterns: Object.keys(L.render.PATTERNS).map((key) => ({ key, label: (L.geom.PATTERN_NAMES || {})[key] || key, src: L.render.patternURL(key, '#000000', '#FFFFFF') })),
    colors: [
      ...(ui.schemeColors ? ui.schemeColors(design()).map((color) => ({ value: color.ref, label: color.tip || color.ref, color: color.hex })) : []),
      ...L.color.STANDARD.map((color) => ({ value: color, label: ui.colorName(color), color })),
    ],
  } : null;
  D.reactFillEffectsBuild = (config) => {
    if (!reactFillEffects || !config || typeof config !== 'object') return null;
    const d = design();
    const mode = ['one', 'two', 'preset'].includes(config.mode) ? config.mode : 'two';
    const presets = GRAD_PRESETS;
    const presetIndex = Number.isInteger(config.preset) ? Math.max(0, Math.min(presets.length - 1, config.preset)) : 0;
    const style = ['horz', 'vert', 'dup', 'ddn', 'corner', 'center'].includes(config.style) ? config.style : 'horz';
    const c1 = typeof config.c1 === 'string' ? config.c1 : '#3366CC';
    const c2 = typeof config.c2 === 'string' ? config.c2 : '#99B3E5';
    const dark = Number.isFinite(config.dark) ? Math.max(0, Math.min(1, config.dark)) : .5;
    const tr1 = Number.isFinite(config.tr1) ? Math.max(0, Math.min(1, config.tr1)) : 0;
    const tr2 = Number.isFinite(config.tr2) ? Math.max(0, Math.min(1, config.tr2)) : 0;
    const cc1 = L.model.resolveColor(c1, d);
    const cc2 = mode === 'one' ? (dark < .5 ? L.color.darken(cc1, 1 - dark * 2) : L.color.lighten(cc1, (dark - .5) * 2)) : L.model.resolveColor(c2, d);
    const baseStops = mode === 'preset'
      ? presets[presetIndex][1].map(([p, c]) => ({ p, c, a: 1 }))
      : [{ p: 0, c: cc1, a: 1 - tr1 }, { p: 1, c: cc2, a: 1 - tr2 }];
    const angle = { horz: 90, vert: 0, dup: 315, ddn: 45, corner: 0, center: 0 }[style];
    const variants = Array.from({ length: 4 }, (_, variant) => {
      let stops = baseStops;
      if (variant === 1 || variant === 3) stops = stops.map((stop) => ({ p: 1 - stop.p, c: stop.c, a: stop.a })).reverse();
      if (variant >= 2) stops = stops.map((stop) => ({ p: stop.p / 2, c: stop.c, a: stop.a })).concat(stops.slice(0, -1).reverse().map((stop) => ({ p: 1 - stop.p / 2, c: stop.c, a: stop.a })));
      const fill = { t: 'grad', stops, ang: angle, path: style === 'center' ? 'circle' : style === 'corner' ? 'rect' : 'lin' };
      return { fill, css: fillCSS(fill) };
    });
    return { variants, patterns: Object.keys(L.render.PATTERNS).map((key) => ({ key, label: (L.geom.PATTERN_NAMES || {})[key] || key, src: L.render.patternURL(key, L.model.resolveColor(config.fg || '#000000', d), L.model.resolveColor(config.bg || '#FFFFFF', d)) })) };
  };
  D.reactFillEffectsTexture = async (name) => {
    if (!reactFillEffects || !TEXTURES.includes(name)) return null;
    const bytes = atob(texture(name).split(',', 2)[1]);
    const blob = new Blob([Uint8Array.from(bytes, (value) => value.charCodeAt(0))], { type: 'image/png' });
    return { t: 'img', media: L.media.add(blob, name + '.png'), tile: true };
  };
  D.reactFillEffectsImage = (file, tile) => {
    if (!reactFillEffects || !(file instanceof Blob) || !file.type.startsWith('image/')) return null;
    return { t: 'img', media: L.media.add(file, file instanceof File ? file.name : 'picture.png'), ...(tile ? { tile: true } : {}) };
  };
  D.reactFillEffectsFinish = (fill) => {
    if (!reactFillEffects) return false;
    const resolve = reactFillEffectsResolve;
    reactFillEffectsResolve = null;
    reactFillEffects = null;
    L.bus.emit('pptx-fill-effects-dialog');
    resolve?.(fill ? L.clone(fill) : null);
    return true;
  };
  D.reactFillEffectsPreview = (fill) => fill && fill.t ? fillCSS(fill) : '';
  D.reactBackgroundCatalog = reactBackgroundCatalog;
  D.reactBackgroundColor = (value) => {
    if (!reactBackground || typeof value !== 'string' || (!reactBackground.colors.some((color) => color.value === value) && !/^#[0-9a-f]{6}$/i.test(value))) return false;
    reactBackground.fill = { t: 'solid', c: value, a: reactBackground.fill?.t === 'solid' && reactBackground.fill.a != null ? reactBackground.fill.a : 1 };
    L.bus.emit('pptx-background-dialog');
    return true;
  };
  D.reactBackgroundEffects = async () => {
    if (!reactBackground) return false;
    const fill = await D.fillEffects(reactBackground.fill && reactBackground.fill.t !== 'none' ? reactBackground.fill : { t: 'solid', c: 'accent1', a: 1 });
    if (!fill || !reactBackground) return false;
    reactBackground.fill = fill;
    L.bus.emit('pptx-background-dialog');
    return true;
  };
  D.reactBackgroundAction = (action, value) => {
    if (!reactBackground) return false;
    if (action === 'close') { reactBackground = null; L.bus.emit('pptx-background-dialog'); return true; }
    if (action === 'omit' && typeof value === 'boolean') { reactBackground.omit = value; L.bus.emit('pptx-background-dialog'); return true; }
    if (action !== 'apply' && action !== 'apply-all') return false;
    const slide = reactBackground.slide, fill = reactBackground.fill, omit = reactBackground.omit, all = action === 'apply-all';
    L.hist.push(L.N('Background'));
    const targets = all ? L.pres.slides : (L.app.selectedSlides && L.app.selectedSlides().length > 1 ? L.app.selectedSlides() : [slide]);
    if (E().view === 'master') { const master = E().masterSlide; master.bg = fill; E().syncMaster(); }
    else if (all) { for (const item of L.pres.slides) { delete item.bg; item.hideMaster = omit || undefined; } for (const id in L.pres.designs) if (L.pres.slides.some((item) => item.design === id)) L.pres.designs[id].bg = L.clone(fill); }
    else for (const item of targets) { item.bg = L.clone(fill); item.hideMaster = omit || undefined; }
    E().render(); L.bus.emit('slides-changed');
    reactBackground = null; L.bus.emit('pptx-background-dialog');
    return true;
  };
  D.reactFormatShapeCatalog = () => {
    if (!reactFormatShape) return null;
    const { shapes, shape, ...catalog } = reactFormatShape;
    return Object.assign({}, catalog, { fill: L.clone(catalog.fill), line: L.clone(catalog.line), colors: catalog.colors.map((color) => Object.assign({}, color)) });
  };
  D.reactFormatShapePreview = (fill) => fill ? fillCSS(fill) : '';
  D.reactFormatShapeEffects = async (fill) => {
    if (!reactFormatShape) return null;
    return D.fillEffects(fill && fill.t !== 'none' ? fill : { t: 'solid', c: 'accent1', a: 1 });
  };
  D.reactFormatShapeAction = (action, changes) => {
    if (!reactFormatShape) return false;
    if (action === 'close') { reactFormatShape = null; L.bus.emit('pptx-format-shape-dialog'); return true; }
    if (action !== 'apply' || !changes || typeof changes !== 'object') return false;
    const converted = Object.assign({}, changes);
    for (const key of ['w', 'h', 'x', 'y']) if (converted[key] != null) converted[key] = PT(converted[key]);
    if (converted.cropLeft != null || converted.cropRight != null || converted.cropTop != null || converted.cropBottom != null) {
      const crop = reactFormatShape.crop, fullW = reactFormatShape.cropFullWidthPt, fullH = reactFormatShape.cropFullHeightPt;
      converted.crop = {
        l: PT(converted.cropLeft == null ? crop.l : converted.cropLeft) / fullW,
        r: PT(converted.cropRight == null ? crop.r : converted.cropRight) / fullW,
        t: PT(converted.cropTop == null ? crop.t : converted.cropTop) / fullH,
        b: PT(converted.cropBottom == null ? crop.b : converted.cropBottom) / fullH,
      };
      delete converted.cropLeft; delete converted.cropRight; delete converted.cropTop; delete converted.cropBottom;
    }
    const { shapes, title } = reactFormatShape;
    applyShapeChanges(shapes, converted, title);
    reactFormatShape = null; L.bus.emit('pptx-format-shape-dialog');
    return true;
  };
  D.find = function (replace) {
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactFind = { what: (L.te.active() && L.te.selectedText()) || findState.what, with: findState.with, mc: findState.mc, ww: findState.ww, replace: !!replace };
      L.bus.emit('pptx-find-dialog');
      return;
    }
    const what = h('input', { type: 'text', id: 'fd-what', value: (L.te.active() && L.te.selectedText()) || findState.what, style: 'width:260px' });
    const repl = h('input', { type: 'text', id: 'fd-with', value: findState.with, style: 'width:260px' });
    const mc = ui.check('Matc&h case', findState.mc), ww = ui.check('Find &whole words only', findState.ww);
    const read = () => { findState.what = what.value; findState.with = repl.value; findState.mc = mc.input.checked; findState.ww = ww.input.checked; };
    const body = h('div', { class: 'col' }, ui.field('Fi&nd what:', what, { cls: 'wide' }), replace ? ui.field('Re&place with:', repl, { cls: 'wide' }) : null, mc, ww);
    const btns = [{ label: '&Find Next', primary: true, onClick: () => { read(); if (!findNext(findState)) findNotFound(); return false; } }];
    if (replace) {
      btns.push({ label: '&Replace', onClick: () => { read(); replaceOne(); return false; } });
      btns.push({ label: 'Replace &All', onClick: () => { read(); const n = replaceAll(); ui.msg(n > 1 ? L.t('{0} has finished searching the presentation and made {1} replacements.', L.APP, n) : n ? L.t('{0} has finished searching the presentation and made 1 replacement.', L.APP) : 'The search item wasn\'t found.', { icon: 'info' }); return false; } });
    } else btns.push({ label: 'Re&place...', onClick: (close) => { read(); close(null); setTimeout(() => D.find(true), 0); return false; } });
    btns.push({ label: 'Close' });
    ui.dialog({ title: replace ? 'Replace' : 'Find', body, width: 440, buttons: btns });
  };

  /* ---------- Insert Table ---------- */
  D.insertTable = function (cb) {
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactInsertTable = { rows: 2, columns: 3 };
      reactInsertTableCallback = typeof cb === 'function' ? cb : null;
      L.bus.emit('pptx-insert-table-dialog');
      return;
    }
    const c = ui.spin({ value: 3, min: 1, max: 75, step: 1, dec: 0 }), r = ui.spin({ value: 2, min: 1, max: 75, step: 1, dec: 0 });
    ui.dialog({ title: 'Insert Table', body: h('div', { class: 'col' }, ui.field('Number of &columns:', c, { cls: 'wide' }), ui.field('Number of &rows:', r, { cls: 'wide' })), width: 280, buttons: [{ label: 'OK', primary: true, onClick: () => cb(r.get(), c.get()) }, { label: 'Cancel' }] });
  };

  /* ---------- Hyperlink ---------- */
  D.hyperlink = function () {
    const ed = E();
    const editing = L.te.active();
    const sh = ed.primary();
    if (!sh && !editing) { ui.msg('Select text or an object to link.'); return; }
    const runProps = editing ? L.te.runPropsAtSel() : {};
    const existing = editing ? (L.te.target() && (() => { const t = L.te.target(); const s = L.te.currentSel(); const p = t.tx.ps[s.p0]; let acc = 0; for (const r of p.rs) { acc += r.t.length; if (acc >= s.o0 && r.link) return r.link; } return null; })()) : sh.link;
    void runProps;
    let kind = existing && existing.slide ? 'place' : existing && existing.url && /^mailto:/.test(existing.url) ? 'mail' : existing && existing.action ? 'place' : 'web';
    const places = [['a:first', 'First Slide'], ['a:last', 'Last Slide'], ['a:next', 'Next Slide'], ['a:prev', 'Previous Slide']].concat(L.pres.slides.map((s, i) => ['s:' + s.id, `${i + 1}. ${M.slideTitle(s) || L.t('Slide {0}', i + 1)}`]));
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const place = existing && existing.slide ? 's:' + existing.slide : existing && existing.action ? 'a:' + existing.action : 'a:next';
      reactHyperlink = { title: existing ? 'Edit Hyperlink' : 'Insert Hyperlink', existing: !!existing, editing, kind, display: editing ? L.te.selectedText() : L.t('<<Selection in Document>>'), address: existing && existing.url && !/^mailto:/.test(existing.url) ? existing.url : 'https://', tip: (existing && existing.tip) || '', email: existing && existing.url && /^mailto:/.test(existing.url) ? existing.url.replace(/^mailto:/, '').split('?')[0] : '', subject: '', place, places: places.map(([value, label]) => ({ value, label })) };
      reactHyperlinkApply = (link, display) => {
        if (editing) {
          const selection = L.te.currentSel();
          if (selection && L.txt.collapsed(selection) && display) { L.hist.push(L.N('Hyperlink')); L.te.insertText(display); const next = L.te.currentSel(); L.te.applyRun((run) => { const result = Object.assign({}, run); if (link) result.link = link; else delete result.link; return result; }, 'Hyperlink', { p0: next.p0, o0: next.o0 - display.length, p1: next.p1, o1: next.o1 }); }
          else L.te.applyRun((run) => { const result = Object.assign({}, run); if (link) result.link = link; else delete result.link; return result; }, 'Hyperlink');
        } else ed.commit(L.N('Hyperlink'), () => { for (const shape of ed.selected()) { if (link) shape.link = link; else delete shape.link; } });
      };
      L.bus.emit('pptx-hyperlink-dialog');
      return;
    }
    const disp = h('input', { type: 'text', id: 'hl-disp', value: editing ? L.te.selectedText() : L.t('<<Selection in Document>>'), style: 'width:100%' });
    if (!editing) disp.disabled = true;
    const addr = h('input', { type: 'text', id: 'hl-addr', value: existing && existing.url && !/^mailto:/.test(existing.url) ? existing.url : 'https://', style: 'width:100%' });
    const tip = h('input', { type: 'text', id: 'hl-tip', value: (existing && existing.tip) || '', style: 'width:100%' });
    const mail = h('input', { type: 'text', id: 'hl-mail', value: existing && existing.url && /^mailto:/.test(existing.url) ? existing.url.replace(/^mailto:/, '').split('?')[0] : '', style: 'width:100%' });
    const subj = h('input', { type: 'text', id: 'hl-subj', value: '', style: 'width:100%' });
    const placeSel = h('select', { id: 'hl-place', size: 10, style: 'width:100%' }, ...places.map(([v, l], i) => h('option', { value: v, text: l, translate: i < 4 ? null : 'no' })));
    if (existing && existing.slide) placeSel.value = 's:' + existing.slide; else if (existing && existing.action) placeSel.value = 'a:' + existing.action; else placeSel.selectedIndex = 2;
    const panes = {
      web: h('div', { class: 'col' }, h('label', { for: 'hl-addr', html: ui.ampHTML('Addr&ess:') }), addr, h('label', { for: 'hl-tip', text: 'ScreenTip:' }), tip),
      place: h('div', { class: 'col' }, h('label', { for: 'hl-place', html: ui.ampHTML('Select a place in this d&ocument:') }), placeSel),
      mail: h('div', { class: 'col' }, h('label', { for: 'hl-mail', html: ui.ampHTML('&E-mail address:') }), mail, h('label', { for: 'hl-subj', html: ui.ampHTML('Su&bject:') }), subj),
    };
    const host = h('div', { style: 'flex:1;min-width:0' });
    const side = h('div', { class: 'list-box', style: 'width:130px;min-height:180px' });
    const show = (k) => { kind = k; L.clear(host).appendChild(panes[k]); L.$$('.li', side).forEach((x) => x.classList.toggle('on', x.dataset.k === k)); };
    [['web', 'Existing File or Web Page'], ['place', 'Place in This Document'], ['mail', 'E-mail Address']].forEach(([k, l]) => { const it = h('div', { class: 'li', 'data-k': k, text: l, style: 'white-space:normal;padding:8px 6px' }); it.addEventListener('click', () => show(k)); side.appendChild(it); });
    show(kind);
    const body = h('div', { class: 'col' }, h('div', { class: 'field' }, h('label', { for: 'hl-disp', html: ui.ampHTML('&Text to display:') }), disp), h('div', { class: 'row', style: 'flex-wrap:nowrap' }, h('div', { class: 'col' }, h('label', { text: 'Link to:' }), side), host));
    const apply = (link) => {
      if (editing) {
        const t = L.te.target();
        const s = L.te.currentSel();
        if (s && L.txt.collapsed(s) && disp.value) { L.hist.push(L.N('Hyperlink')); L.te.insertText(disp.value); const s2 = L.te.currentSel(); L.te.applyRun((r) => { const n = Object.assign({}, r); if (link) n.link = link; else delete n.link; return n; }, 'Hyperlink', { p0: s2.p0, o0: s2.o0 - disp.value.length, p1: s2.p1, o1: s2.o1 }); }
        else L.te.applyRun((r) => { const n = Object.assign({}, r); if (link) n.link = link; else delete n.link; return n; }, 'Hyperlink');
        void t;
      } else ed.commit(L.N('Hyperlink'), () => { for (const x of ed.selected()) { if (link) x.link = link; else delete x.link; } });
    };
    const buttons = [{
      label: 'OK', primary: true, onClick: () => {
        let link = null;
        if (kind === 'web') { const u = addr.value.trim(); if (u && u !== 'https://') link = { url: /^[a-z]+:/i.test(u) ? u : 'https://' + u }; if (link && tip.value) link.tip = tip.value; }
        if (kind === 'mail' && mail.value.trim()) link = { url: 'mailto:' + mail.value.trim() + (subj.value ? '?subject=' + encodeURIComponent(subj.value) : '') };
        if (kind === 'place') { const v = placeSel.value || 'a:next'; link = v.startsWith('s:') ? { slide: v.slice(2) } : { action: v.slice(2) }; }
        if (link) apply(link);
      },
    }, { label: 'Cancel' }];
    if (existing) buttons.unshift({ label: '&Remove Link', onClick: () => apply(null) });
    ui.dialog({ title: existing ? 'Edit Hyperlink' : 'Insert Hyperlink', body, width: 560, buttons });
  };

  /* ---------- Action Settings ---------- */
  D.actionSettings = function () {
    const shapes = E().selected();
    if (!shapes.length) { ui.msg('Select an object first.'); return; }
    const cur = shapes[0].link;
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const currentKind = cur ? (cur.slide ? 'slide' : cur.url ? 'url' : cur.action) : 'next';
      reactActionShapes = shapes;
      reactActionSettings = {
        none: !cur, kind: ['next', 'prev', 'first', 'last', 'lastViewed', 'end', 'slide', 'url'].includes(currentKind) ? currentKind : 'next',
        slide: cur && cur.slide || L.pres.slides[0]?.id || '', url: cur && cur.url || 'https://',
        actions: [['next', 'Next Slide'], ['prev', 'Previous Slide'], ['first', 'First Slide'], ['last', 'Last Slide'], ['lastViewed', 'Last Slide Viewed'], ['end', 'End Show'], ['slide', 'Slide...'], ['url', L.N('URL...')]],
        slides: L.pres.slides.map((slide, index) => ({ id: slide.id, label: `${index + 1}. ${M.slideTitle(slide) || L.t('Slide {0}', index + 1)}` })),
      };
      L.bus.emit('pptx-action-settings-dialog');
      return;
    }
    let none = !cur;
    const opts = [['next', 'Next Slide'], ['prev', 'Previous Slide'], ['first', 'First Slide'], ['last', 'Last Slide'], ['lastViewed', 'Last Slide Viewed'], ['end', 'End Show'], ['slide', 'Slide...'], ['url', L.N('URL...')]];
    const sel = ui.select(opts, cur ? (cur.slide ? 'slide' : cur.url ? 'url' : cur.action) : 'next', () => sync());
    const slideSel = ui.select(L.pres.slides.map((s, i) => [s.id, `${i + 1}. ${M.slideTitle(s) || L.t('Slide {0}', i + 1)}`]), cur && cur.slide, null, { translate: 'no' });
    const url = h('input', { type: 'text', id: 'as-url', value: (cur && cur.url) || 'https://', style: 'width:240px' });
    const rn = ui.radio('as-k', '&None', none, () => { none = true; sync(); }), rl = ui.radio('as-k', '&Hyperlink to:', !none, () => { none = false; sync(); });
    const sync = () => { sel.disabled = none; slideSel.hidden = none || sel.value !== 'slide'; url.hidden = none || sel.value !== 'url'; };
    sync();
    const body = ui.tabs([{ label: 'Mouse Click', body: h('div', { class: 'col' }, h('b', { text: 'Action on click' }), rn, rl, h('div', { class: 'col', style: 'margin-left:20px' }, sel, slideSel, url)) }]);
    ui.dialog({
      title: 'Action Settings', body, width: 360, buttons: [{
        label: 'OK', primary: true, onClick: () => {
          let link = null;
          if (!none) link = sel.value === 'slide' ? { slide: slideSel.value } : sel.value === 'url' ? { url: url.value } : { action: sel.value };
          E().commit('Action Settings', () => { for (const s of shapes) { if (link) s.link = link; else delete s.link; } });
        },
      }, { label: 'Cancel' }],
    });
  };

  /* ---------- Symbol ---------- */
  const SYM_SETS = [
    [L.N('Latin-1 Supplement'), 0x00a1, 0x00ff], [L.N('Latin Extended-A'), 0x0100, 0x017f], [L.N('Greek and Coptic'), 0x0391, 0x03c9], [L.N('General Punctuation'), 0x2010, 0x2044],
    [L.N('Currency Symbols'), 0x20a0, 0x20bf], [L.N('Letterlike Symbols'), 0x2100, 0x214f], [L.N('Number Forms'), 0x2150, 0x218b], [L.N('Arrows'), 0x2190, 0x21ff], [L.N('Mathematical Operators'), 0x2200, 0x22ff],
    [L.N('Box Drawing'), 0x2500, 0x257f], [L.N('Geometric Shapes'), 0x25a0, 0x25ff], [L.N('Miscellaneous Symbols'), 0x2600, 0x26ff], [L.N('Dingbats'), 0x2700, 0x27bf],
  ];
  const recentSyms = ['€', '£', '¥', '©', '®', '™', '±', '≠', '≤', '≥', '÷', '×', '∞', 'µ', 'α', 'β'];
  D.symbol = function (cb) {
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactSymbolCallback = typeof cb === 'function' ? cb : null;
      reactSymbol = {
        font: 'normal', fonts: ['normal', 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI Symbol'],
        subset: 0, subsets: SYM_SETS.map((s) => s[0]), selected: null, code: '',
        recent: [...recentSyms],
        glyphs: [],
      };
      const [, first, last] = SYM_SETS[0];
      reactSymbol.glyphs = Array.from({ length: last - first + 1 }, (_, i) => String.fromCodePoint(first + i));
      L.bus.emit('pptx-symbol-dialog');
      return;
    }
    let pick = null;
    const grid = h('div', { class: 'sym-grid', translate: 'no' });
    const font = ui.select([['normal', '(normal text)'], 'Arial', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI Symbol'], 'normal', () => fill(sub.value));
    const sub = ui.select(SYM_SETS.map((s, i) => [i, s[0]]), 0, (v) => fill(v));
    const code = h('input', { type: 'text', id: 'sym-code', style: 'width:70px', readonly: true });
    const recent = h('div', { class: 'row', style: 'gap:2px' });
    const fill = (si) => {
      const [, a, b] = SYM_SETS[+si];
      L.clear(grid);
      grid.style.fontFamily = font.value === 'normal' ? '' : L.fontStack(font.value);
      for (let c = a; c <= b; c++) {
        const ch = String.fromCodePoint(c);
        const btn = h('button', { type: 'button', text: ch, 'aria-label': 'U+' + c.toString(16).toUpperCase(), translate: 'no' });
        btn.addEventListener('click', () => { L.$$('button.on', grid).forEach((x) => x.classList.remove('on')); btn.classList.add('on'); pick = ch; code.value = c.toString(16).toUpperCase().padStart(4, '0'); });
        btn.addEventListener('dblclick', () => { pick = ch; insertBtn.click(); });
        grid.appendChild(btn);
      }
    };
    fill(0);
    for (const ch of recentSyms) { const b = h('button', { type: 'button', class: 'btn small', text: ch, style: 'min-width:24px;padding:0' }); b.addEventListener('click', () => { pick = ch; code.value = ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); }); recent.appendChild(b); }
    const body = h('div', { class: 'col' }, h('div', { class: 'row' }, ui.field('&Font:', font), ui.field('Su&bset:', sub)), grid, h('label', { text: 'Recently used symbols:' }), recent, ui.field('&Character code:', code));
    const d = ui.dialog({
      title: 'Symbol', body, width: 470, buttons: [{
        label: '&Insert', primary: true, onClick: () => {
          if (!pick) return false;
          const i = recentSyms.indexOf(pick); if (i >= 0) recentSyms.splice(i, 1); recentSyms.unshift(pick); recentSyms.length = 16;
          if (cb) { cb(pick); return; }
          if (L.te.active()) { L.hist.push('Insert Symbol'); L.te.insertText(pick); }
        },
      }, { label: 'Cancel' }],
    });
    const insertBtn = d.buttons[0];
  };

  /* ---------- Date and Time ---------- */
  D.dateTime = function () {
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const now = new Date();
      reactDateTime = { formats: L.DATE_FORMATS.map((id) => ({ id, label: L.fmtDate(now, id) })), selected: L.DATE_FORMATS[0] || '' };
      L.bus.emit('pptx-date-time-dialog');
      return;
    }
    const list = h('select', { id: 'dt-list', size: 12, style: 'width:260px', translate: 'no' }, ...L.DATE_FORMATS.map((f) => h('option', { value: f, text: L.fmtDate(new Date(), f), translate: 'no' })));
    list.selectedIndex = 0;
    const upd = ui.check('&Update automatically', false);
    ui.dialog({
      title: 'Date and Time', body: h('div', { class: 'col' }, h('label', { for: 'dt-list', html: ui.ampHTML('&Available formats:') }), list, upd), width: 300, buttons: [{
        label: 'OK', primary: true, onClick: () => {
          const f = list.value || 'datetime1';
          if (!L.te.active()) { L.app.ensureTextTarget(); }
          if (!L.te.active()) return;
          L.hist.push('Insert Date and Time');
          if (upd.input.checked) L.te.insertField(f); else L.te.insertText(L.fmtDate(new Date(), f));
        },
      }, { label: 'Cancel' }],
    });
  };

  /* ---------- Set Up Show ---------- */
  D.setupShow = function () {
    const s = L.pres.show;
    const n = L.pres.slides.length;
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactSetupShow = { kiosk: !!s.kiosk, loop: !!s.loop, noAnimation: !!s.noAnim, penColor: s.penColor || '#FF0000', range: !!(s.from && s.to), from: s.from || 1, to: s.to || n, useTimings: s.useTimings !== false, slideCount: n };
      L.bus.emit('pptx-setup-show-dialog');
      return;
    }
    const speaker = ui.radio('su-type', '&Presented by a speaker (full screen)', !s.kiosk), kiosk = ui.radio('su-type', 'Browsed at a &kiosk (full screen)', !!s.kiosk);
    const loop = ui.check('&Loop continuously until \'Esc\'', !!s.loop), noAnim = ui.check('Show without &animation', !!s.noAnim);
    const pen = colorPick(s.penColor || '#FF0000');
    const all = ui.radio('su-rng', '&All', !(s.from && s.to)), rng = ui.radio('su-rng', '&From:', !!(s.from && s.to));
    const from = ui.spin({ value: s.from || 1, min: 1, max: Math.max(1, n), step: 1, dec: 0 }), to = ui.spin({ value: s.to || n, min: 1, max: Math.max(1, n), step: 1, dec: 0 });
    const man = ui.radio('su-adv', '&Manually', s.useTimings === false), tim = ui.radio('su-adv', '&Using timings, if present', s.useTimings !== false);
    const body = h('div', { class: 'row' },
      h('div', { class: 'col' }, ui.group('Show type', speaker, kiosk), ui.group('Show options', loop, noAnim, ui.field('P&en color:', pen))),
      h('div', { class: 'col' }, ui.group('Show slides', all, h('div', { class: 'row', style: 'align-items:center' }, rng, from, h('span', { text: 'To:' }), to)), ui.group('Advance slides', man, tim)));
    ui.dialog({ title: 'Set Up Show', body, width: 560, buttons: [{ label: 'OK', primary: true, onClick: () => { L.hist.push('Set Up Show'); Object.assign(s, { kiosk: kiosk.input.checked, loop: loop.input.checked || kiosk.input.checked, noAnim: noAnim.input.checked, penColor: pen.get() && pen.get()[0] === '#' ? pen.get() : L.model.resolveColor(pen.get(), design()), from: rng.input.checked ? from.get() : 0, to: rng.input.checked ? to.get() : 0, useTimings: tim.input.checked }); } }, { label: 'Cancel' }] });
  };

  /* ---------- Options ---------- */
  D.options = function () {
    const o = L.app.opts;
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const look = window.VO?.look, lang = window.VO?.lang;
      const optionKeys = ['startupPane', 'layoutPaneOnNew', 'paneFollows', 'statusBar', 'ruler', 'showPopup', 'showPopbar', 'endBlack', 'smartQuotes', 'autocorrect', 'capSentence', 'spell'];
      reactOptions = { options: Object.fromEntries(optionKeys.map((key) => [key, !!o[key]])), optionKeys, undo: L.hist.limit, user: L.pres.props.author || '', password: L.pres.password || '', looks: look?.list?.().map((item) => [item.id, item.name]) || [], languages: lang?.list?.().map((item) => [item.id, item.id === 'auto' ? L.t('Auto (as the browser)') : item.name]) || [], look: look?.get?.() || '', lang: lang?.get?.() || '' };
      L.bus.emit('pptx-options-dialog');
      return;
    }
    const c = (label, key) => ui.check(label, !!o[key], (v) => { o[key] = v; });
    const undo = ui.spin({ value: L.hist.limit, min: 3, max: 150, step: 1, dec: 0 });
    const user = h('input', { type: 'text', id: 'op-user', value: L.pres.props.author || '', style: 'width:220px' });
    const pwOpen = h('input', { type: 'password', id: 'op-pw', value: L.pres.password || '', autocomplete: 'new-password', style: 'width:180px' });
    const look = ui.lookField();
    const tabs = ui.tabs([
      { label: 'View', body: h('div', { class: 'col' }, ui.group('Show', c('Startup &Task Pane', 'startupPane'), c('&Slide Layout task pane when inserting new slides', 'layoutPaneOnNew'), c('Task pane f&ollows the selected object', 'paneFollows'), c('Stat&us bar', 'statusBar'), c('&Vertical ruler', 'ruler')), ui.group('Slide show', c('&Popup menu on right mouse click', 'showPopup'), c('Show popup menu &button', 'showPopbar'), c('&End with black slide', 'endBlack'))) },
      { label: 'General', body: h('div', { class: 'col' }, look && look.el, ui.group('User information', ui.field('&Name:', user))) },
      { label: 'Edit', body: h('div', { class: 'col' }, ui.group('Text', c('Replace straight &quotes with smart quotes', 'smartQuotes'), c('&AutoCorrect as you type', 'autocorrect'), c('Capitalize first letter of &sentences', 'capSentence')), ui.group('Undo', ui.field('&Maximum number of undos:', undo, { cls: 'wide' }))) },
      { label: 'Spelling and Style', body: h('div', { class: 'col' }, ui.group('Spelling', c('Check spelling as you &type', 'spell')), h('div', { class: 'tp-note', text: 'Spelling suggestions come from your browser. Hold Shift while right-clicking a word to see them.' })) },
      { label: 'Security', body: h('div', { class: 'col' }, ui.group('File encryption settings for this document', ui.field('Password to &open:', pwOpen, { cls: 'wide' })), h('div', { class: 'tp-note', text: 'The presentation is encrypted (AES-256) the next time you save it as a PowerPoint file, the way PowerPoint 2007 and later do. Leave the box empty to save it without a password. If you lose or forget the password, it cannot be recovered.' })) },
    ]);
    const applyPw = () => {
      const pw = pwOpen.value || '';
      if (pw === (L.pres.password || '')) return true;
      if (!pw) { delete L.pres.password; L.hist.dirty = true; return true; }
      /* confirm a new password, as PowerPoint does */
      const p2 = h('input', { type: 'password', id: 'pw-confirm', autocomplete: 'new-password', style: 'width:100%' });
      ui.dialog({ title: 'Confirm Password', width: 320, body: h('div', { class: 'col' }, h('label', { for: 'pw-confirm', text: 'Reenter password to open:' }), p2, h('div', { class: 'tp-note', text: 'Caution: If you lose or forget the password, it cannot be recovered.' })), buttons: [{ label: 'OK', primary: true, onClick: () => { if (p2.value !== pw) { ui.msg('Confirmation password is not identical.', { icon: 'warn' }); return false; } L.pres.password = pw; L.hist.dirty = true; } }, { label: 'Cancel' }] });
      setTimeout(() => p2.focus(), 0);
      return true;
    };
    ui.dialog({ title: 'Options', body: tabs, width: 460, buttons: [{ label: 'OK', primary: true, onClick: () => { L.hist.limit = undo.get(); L.pres.props.author = user.value; L.app.saveOpts(); L.app.applyOpts(); if (look) look.apply(); applyPw(); } }, { label: 'Cancel', onClick: () => { L.app.loadOpts(); } }] });
  };

  /* ---------- AutoCorrect ---------- */
  D.autocorrect = function () {
    const o = L.app.opts;
    const AC = L.te.AUTOCORRECT;
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactAutoCorrect = { options: { autocorrect: !!o.autocorrect, capSentence: !!o.capSentence, smartQuotes: !!o.smartQuotes }, entries: Object.entries(AC).sort(([a], [b]) => a.localeCompare(b)), selected: '', replacement: '' };
      L.bus.emit('pptx-autocorrect-dialog');
      return;
    }
    const list = h('div', { class: 'list-box', style: 'height:150px' });
    const rep = h('input', { type: 'text', id: 'ac-rep', style: 'width:120px' }), wit = h('input', { type: 'text', id: 'ac-with', style: 'width:200px' });
    const draw = () => { L.clear(list); Object.keys(AC).sort().forEach((k) => { const li = h('div', { class: 'li', translate: 'no' }, h('span', { text: k, style: 'width:110px;display:inline-block', translate: 'no' }), h('span', { text: AC[k], translate: 'no' })); li.addEventListener('click', () => { L.$$('.li', list).forEach((x) => x.classList.remove('on')); li.classList.add('on'); rep.value = k; wit.value = AC[k]; }); list.appendChild(li); }); };
    draw();
    const body = ui.tabs([{
      label: 'AutoCorrect', body: h('div', { class: 'col' },
        ui.check('Correct TWo INitial CApitals', !!o.autocorrect, (v) => { o.autocorrect = v; }),
        ui.check('Capitalize first letter of &sentences', !!o.capSentence, (v) => { o.capSentence = v; }),
        ui.check('Replace straight quotes with smart quotes', !!o.smartQuotes, (v) => { o.smartQuotes = v; }),
        h('div', { class: 'row' }, ui.field('&Replace:', rep), ui.field('&With:', wit)), list,
        h('div', { class: 'row' }, ui.button('&Add', () => { if (rep.value) { AC[rep.value] = wit.value; draw(); } }), ui.button('&Delete', () => { delete AC[rep.value]; draw(); }))),
    }]);
    ui.dialog({ title: 'AutoCorrect: English (U.S.)', body, width: 420, buttons: [{ label: 'OK', primary: true, onClick: () => L.app.saveOpts() }, { label: 'Cancel' }] });
  };

  /* ---------- Properties ---------- */
  D.properties = function () {
    const p = L.pres.props;
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const words = L.pres.slides.reduce((count, slide) => count + (M.allText(slide).match(/\S+/g) || []).length, 0);
      reactProperties = {
        title: L.t('{0} Properties', L.app.fileName || 'Presentation'),
        values: Object.fromEntries(['title', 'subject', 'author', 'company', 'category', 'keywords', 'comments'].map((key) => [key, p[key] || ''])),
        stats: [['Created:', p.created ? new Date(p.created).toLocaleString() : '—'], ['Slides:', L.pres.slides.length], ['Hidden slides:', L.pres.slides.filter((slide) => slide.hidden).length], ['Notes:', L.pres.slides.filter((slide) => slide.notes).length], ['Words:', words], ['Pictures:', L.media.all().size]],
      };
      L.bus.emit('pptx-properties-dialog');
      return;
    }
    const f = (k, label) => { const i = h('input', { type: 'text', id: 'pr-' + k, value: p[k] || '', style: 'width:260px' }); i.addEventListener('input', () => { p[k] = i.value; if (k === 'title') L.pres.title = i.value; }); return ui.field(label, i, { cls: 'wide' }); };
    const cm = h('textarea', { id: 'pr-cm', rows: 3, style: 'width:260px' }); cm.value = p.comments || ''; cm.addEventListener('input', () => { p.comments = cm.value; });
    const words = L.pres.slides.reduce((a, s) => a + (M.allText(s).match(/\S+/g) || []).length, 0);
    const tabs = ui.tabs([
      { label: 'Summary', body: h('div', { class: 'col' }, f('title', '&Title:'), f('subject', '&Subject:'), f('author', '&Author:'), f('company', 'C&ompany:'), f('category', 'Ca&tegory:'), f('keywords', '&Keywords:'), ui.field('&Comments:', cm, { cls: 'wide' })) },
      { label: 'Statistics', body: h('table', { class: 'kbd-table' }, ...[['Created:', p.created ? new Date(p.created).toLocaleString() : '—'], ['Slides:', L.pres.slides.length], ['Hidden slides:', L.pres.slides.filter((s) => s.hidden).length], ['Notes:', L.pres.slides.filter((s) => s.notes).length], ['Words:', words], ['Pictures:', L.media.all().size]].map(([a, b]) => h('tr', null, h('td', { text: a }), h('td', { text: String(b), translate: 'no' })))) },
    ]);
    ui.dialog({ title: L.t('{0} Properties', L.app.fileName || 'Presentation'), body: tabs, width: 430, buttons: [{ label: 'OK', primary: true }] });
  };

  /* ---------- Save As ---------- */
  D.saveAs = function (cb) {
    const base = (L.app.fileName || 'Presentation1').replace(/\.[^.]+$/, '');
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const types = L.file.types.map((t) => [t.id, L.t(t.label)]);
      const cur = types.some(([k]) => k === L.app.fileType) ? L.app.fileType : 'pptx';
      reactSaveAs = { name: base, type: cur, types, note: L.t('Your browser decides where downloaded files go.') };
      reactSaveAsCallback = cb;
      L.bus.emit('pptx-save-as-dialog');
      return new Promise((resolve, reject) => { reactSaveAsResolve = resolve; reactSaveAsReject = reject; });
    }
    const name = h('input', { type: 'text', id: 'sa-name', value: base, style: 'width:300px' });
    const types = L.file.types.map((t) => [t.id, t.label]);
    const cur = types.some(([k]) => k === L.app.fileType) ? L.app.fileType : 'pptx';
    const type = ui.select(types, cur);
    return ui.dialog({ title: 'Save As', body: h('div', { class: 'col' }, ui.field('File &name:', name, { cls: 'wide' }), ui.field('Save as &type:', type, { cls: 'wide' }), h('div', { class: 'tp-note', text: 'Your browser decides where downloaded files go.' })), width: 470, buttons: [{ label: '&Save', primary: true, onClick: () => cb((name.value.trim() || base).replace(/\.(pptx|pptm|ppsx|ppsm|potx|potm)$/i, ''), type.value) }, { label: 'Cancel' }] });
  };

  /* ---------- Zoom ---------- */
  D.zoom = function () {
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactZoom = { value: E().zoom === 'fit' ? 'fit' : Math.round(E().zoom * 100) };
      L.bus.emit('pptx-zoom-dialog');
      return;
    }
    const presets = [['fit', L.N('Fit')], [4, '400%'], [2, '200%'], [1, '100%'], [0.66, '66%'], [0.5, '50%'], [0.33, '33%']];
    let pick = E().zoom;
    const pct = ui.spin({ value: E().zoomPct(), min: 10, max: 400, step: 1, unit: '%', dec: 0, onChange: (v) => { pick = v / 100; } });
    const body = h('div', { class: 'row' }, ui.group('Zoom to', h('div', { class: 'col' }, ...presets.map(([v, l]) => ui.radio('zm', l, String(v) === String(E().zoom), () => { pick = v; if (v !== 'fit') pct.set(Math.round(v * 100)); })))), h('div', { class: 'col' }, ui.field('&Percent:', pct)));
    ui.dialog({ title: 'Zoom', body, width: 300, buttons: [{ label: 'OK', primary: true, onClick: () => E().setZoom(pick) }, { label: 'Cancel' }] });
  };

  /* ---------- Grid and Guides ---------- */
  D.grid = function () {
    const g = E().grid;
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactGrid = { snap: !!g.snap, size: g.size, show: !!g.show, guides: !!g.guides };
      L.bus.emit('pptx-grid-dialog');
      return;
    }
    const snap = ui.check('&Snap objects to grid', g.snap);
    const sp = ui.select([[4.5, '1/16"'], [6, '0.083"'], [7.2, '0.1"'], [9, '0.125"'], [12, '0.167"'], [14.4, '0.2"'], [18, '0.25"'], [24, '0.333"'], [36, '0.5"'], [72, '1"']], g.size);
    const disp = ui.check('&Display grid on screen', g.show), gd = ui.check('Display drawing &guides on screen', g.guides);
    ui.dialog({ title: 'Grid and Guides', body: h('div', { class: 'col' }, ui.group('Snap to', snap), ui.group('Grid settings', ui.field('Spac&ing:', sp), disp), ui.group('Guide settings', gd)), width: 320, buttons: [{ label: 'OK', primary: true, onClick: () => { g.snap = snap.input.checked; g.size = +sp.value; g.show = disp.input.checked; g.guides = gd.input.checked; E().layout(); L.ui.refresh(); L.app.saveOpts(); } }, { label: 'Cancel' }] });
  };

  /* ---------- WordArt ---------- */
  const WA_STYLES = (() => {
    const s = [];
    const grad = (a, b, ang) => ({ t: 'grad', stops: [{ p: 0, c: a, a: 1 }, { p: 1, c: b, a: 1 }], ang: ang == null ? 90 : ang, path: 'lin' });
    const solid = (c) => ({ t: 'solid', c, a: 1 });
    const ln = (c, w) => ({ c, w: w || 0.75, dash: 'solid' });
    const shd = { c: '#000000', a: 0.35, dx: 3, dy: 3, blur: 0 };
    const rows = [
      [solid('#000000'), null, null, 'textPlain'], [solid('#FFFFFF'), ln('#000000'), null, 'textPlain'], [grad('#3366FF', '#99CCFF'), null, null, 'textPlain'], [solid('#C0C0C0'), ln('#404040'), shd, 'textPlain'], [grad('#FF6600', '#FFCC00'), ln('#993300'), null, 'textArchUp'], [solid('#3366CC'), null, shd, 'textWave1'],
      [grad('#FFFFFF', '#808080', 90), ln('#000000'), shd, 'textPlain'], [solid('#CC0000'), ln('#660000'), null, 'textSlantUp'], [grad('#00CC99', '#006666'), null, shd, 'textPlain'], [solid('#FFFF00'), ln('#FF6600', 1.5), shd, 'textPlain'], [grad('#CC99FF', '#660099'), null, null, 'textArchDown'], [solid('#006600'), ln('#003300'), null, 'textTriangle'],
      [grad('#FFCC00', '#CC6600', 0), ln('#663300'), shd, 'textPlain'], [solid('#000080'), null, shd, 'textChevron'], [grad('#E6E6E6', '#333333'), ln('#000000'), null, 'textPlain'], [solid('#FF99CC'), ln('#CC0066'), shd, 'textWave1'], [grad('#66CCFF', '#003399', 0), null, shd, 'textCanUp'], [solid('#996633'), ln('#663300'), null, 'textSlantDown'],
      [grad('#FFFF99', '#FF9900'), ln('#996600'), shd, 'textArchUp'], [solid('#00CCFF'), ln('#006699'), null, 'textPlain'], [grad('#FF0000', '#660000'), null, shd, 'textPlain'], [solid('#808080'), null, { c: '#000000', a: 0.6, dx: 4, dy: 4, blur: 2 }, 'textPlain'], [grad('#99FF99', '#009900'), ln('#006600'), null, 'textCircle'], [solid('#FF6600'), null, shd, 'textTriangle'],
      [grad('#000000', '#808080'), null, shd, 'textPlain'], [solid('#FFFFFF'), ln('#3366CC', 1.5), shd, 'textArchUp'], [grad('#CC66FF', '#3399FF', 0), null, null, 'textWave1'], [solid('#003366'), ln('#99CCFF'), shd, 'textSlantUp'], [grad('#FFCCCC', '#FF3366'), ln('#990033'), null, 'textChevron'], [solid('#669900'), null, shd, 'textCanUp'],
    ];
    rows.forEach(([fill, line, shadow, warp], i) => s.push({ fill, line: line || { t: 'none' }, shadow, warp, b: i % 3 !== 1, font: ['Arial Black', 'Impact', 'Arial', 'Times New Roman', 'Georgia', 'Trebuchet MS'][i % 6] }));
    return s;
  })();
  D.WA_STYLES = WA_STYLES;
  D.wordartGallery = function (cb, current) {
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactWordartGallery = { selected: Number.isInteger(current) && current >= 0 && current < WA_STYLES.length ? current : 0, previews: WA_STYLES.map((style, index) => {
        const shape = { w: 66, h: 46, wa: Object.assign({}, style, { text: 'LettersArt', vert: false }) };
        return { label: L.t('LettersArt style {0}', index + 1), svg: L.render.wordart(shape, design(), {}).outerHTML };
      }) };
      reactWordartGalleryCallback = typeof cb === 'function' ? cb : null;
      L.bus.emit('pptx-wordart-gallery-dialog');
      return;
    }
    let pick = current != null ? current : 0;
    const grid = h('div', { class: 'wa-grid' });
    WA_STYLES.forEach((st, i) => {
      const b = h('button', { type: 'button', 'aria-label': L.t('LettersArt style {0}', i + 1), class: i === pick ? 'on' : '' });
      const sh = { w: 66, h: 46, wa: Object.assign({}, st, { text: 'LettersArt', vert: false }) };
      b.appendChild(L.render.wordart(sh, design(), {}));
      b.addEventListener('click', () => { L.$$('button', grid).forEach((x) => x.classList.remove('on')); b.classList.add('on'); pick = i; });
      b.addEventListener('dblclick', () => { okb.click(); });
      grid.appendChild(b);
    });
    const d = ui.dialog({ title: 'LettersArt Gallery', body: h('div', { class: 'col' }, h('label', { text: 'Select a LettersArt style:' }), grid), width: 470, buttons: [{ label: 'OK', primary: true, onClick: () => cb(pick) }, { label: 'Cancel' }] });
    const okb = d.buttons[0];
  };
  D.wordartText = function (init, cb) {
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactWordartText = { fonts: [...L.FONT_LIST], sizes: [...L.SIZE_LIST], font: init.font || 'Arial Black', size: 36, b: !!init.b, i: !!init.i, text: init.text || 'Your Text Here' };
      reactWordartTextCallback = typeof cb === 'function' ? cb : null;
      L.bus.emit('pptx-wordart-text-dialog');
      return;
    }
    const font = ui.select(L.FONT_LIST, init.font || 'Arial Black');
    const size = ui.select(L.SIZE_LIST, 36);
    const b = ui.check('&B', !!init.b), i = ui.check('&I', !!init.i);
    const ta = h('textarea', { id: 'wa-text', rows: 5, style: 'width:100%;font-size:18px' });
    ta.value = init.text || 'Your Text Here';
    const sync = () => { ta.style.fontFamily = L.fontStack(font.value); ta.style.fontWeight = b.input.checked ? '700' : '400'; ta.style.fontStyle = i.input.checked ? 'italic' : 'normal'; };
    font.addEventListener('change', sync); b.input.addEventListener('change', sync); i.input.addEventListener('change', sync);
    sync();
    ui.dialog({ title: 'Edit LettersArt Text', body: h('div', { class: 'col' }, h('div', { class: 'row' }, ui.field('&Font:', font), ui.field('&Size:', size), b, i), h('label', { for: 'wa-text', html: ui.ampHTML('&Text:') }), ta), width: 470, focus: '#wa-text', buttons: [{ label: 'OK', primary: true, onClick: () => cb({ text: ta.value, font: font.value, b: b.input.checked, i: i.input.checked, size: +size.value }) }, { label: 'Cancel' }] });
  };

  /* ---------- Photo Album ---------- */
  D.photoAlbum = function () {
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactPhotoAlbum = true;
      L.bus.emit('pptx-photo-album-dialog');
      return;
    }
    const pics = [];
    const list = h('div', { class: 'list-box', style: 'height:150px;width:200px' });
    const prev = h('div', { class: 'preview-box', style: 'width:180px;height:135px' });
    let sel = -1;
    const draw = () => {
      L.clear(list);
      pics.forEach((p, i) => { const li = h('div', { class: 'li' + (i === sel ? ' on' : ''), text: `${i + 1}  ${p.name}`, translate: 'no' }); li.addEventListener('click', () => { sel = i; draw(); prev.style.background = `url(${p.url}) center/contain no-repeat`; }); list.appendChild(li); });
    };
    const layout = ui.select([['fit', 'Fit to slide'], ['1', '1 picture'], ['2', '2 pictures'], ['4', '4 pictures'], ['1t', '1 picture with title'], ['2t', '2 pictures with title'], ['4t', '4 pictures with title']], '1');
    const frame = ui.select([['rect', 'Rectangle'], ['roundRect', 'Rounded Rectangle'], ['simple', 'Simple Frame, White'], ['black', 'Simple Frame, Black'], ['shadow', 'Soft Shadow']], 'rect');
    const caps = ui.check('Captions &below ALL pictures', false);
    const add = ui.button('&File/Disk...', async () => { const fs = await L.pickFiles('image/*', true); for (const f of fs) pics.push({ file: f, name: f.name, url: URL.createObjectURL(f) }); draw(); });
    const rm = ui.button('Re&move', () => { if (sel >= 0) { pics.splice(sel, 1); sel = -1; draw(); prev.style.background = ''; } });
    const up = ui.button('↑', () => { if (sel > 0) { [pics[sel - 1], pics[sel]] = [pics[sel], pics[sel - 1]]; sel--; draw(); } });
    const dn = ui.button('↓', () => { if (sel >= 0 && sel < pics.length - 1) { [pics[sel + 1], pics[sel]] = [pics[sel], pics[sel + 1]]; sel++; draw(); } });
    const body = h('div', { class: 'col' },
      ui.group('Album Content', h('div', { class: 'row' }, h('div', { class: 'col' }, h('label', { text: 'Insert picture from:' }), add), h('div', { class: 'col' }, h('label', { text: 'Pictures in album:' }), list, h('div', { class: 'row' }, up, dn, rm)), h('div', { class: 'col' }, h('label', { text: 'Preview:' }), prev))),
      ui.group('Album Layout', ui.field('&Picture layout:', layout, { cls: 'wide' }), ui.field('Frame s&hape:', frame, { cls: 'wide' }), caps));
    ui.dialog({ title: 'Photo Album', body, width: 600, buttons: [{ label: '&Create', primary: true, onClick: () => { if (!pics.length) { ui.msg('Add at least one picture.'); return false; } L.app.createPhotoAlbum(pics, layout.value, frame.value, caps.input.checked); } }, { label: 'Cancel' }] });
  };

  /* ---------- Slide Finder (Insert ▸ Slides from Files) ---------- */
  D.slideFinder = async function () {
    const files = await L.pickFiles('.pptx,.ppsx,.potx,.pptm');
    if (!files[0]) return;
    ui.busy(true, L.t('Opening {0}...', files[0].name));
    let src;
    try { src = await L.pptx.read(await L.readAsArrayBuffer(files[0])); } catch (e) { ui.busy(false); ui.msg(L.t('{0} could not be opened: {1}', files[0].name, e.message), { icon: 'error' }); return; }
    ui.busy(false);
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactSlideFinder = { name: files[0].name, source: src, selected: new Set(), keep: false };
      L.bus.emit('pptx-slide-finder-dialog');
      return;
    }
    const picked = new Set();
    const grid = h('div', { style: 'display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:10px;max-height:320px;overflow:auto;padding:4px;background:#fff;border:1px solid #888' });
    src.slides.forEach((s, i) => {
      const cell = h('label', { style: 'display:flex;flex-direction:column;gap:3px;align-items:flex-start;cursor:default' });
      const ck = h('input', { type: 'checkbox' });
      ck.addEventListener('change', () => { if (ck.checked) picked.add(i); else picked.delete(i); });
      const th = L.render.thumb(src, s, 124);
      th.style.border = '1px solid #999';
      cell.append(th, h('span', { style: 'display:flex;gap:4px;align-items:center' }, ck, h('span', { text: `${i + 1}. ${(M.slideTitle(s) || L.t('Slide')).slice(0, 24)}`, translate: 'no' })));
      grid.appendChild(cell);
    });
    const keep = ui.check('&Keep source formatting', false);
    const doInsert = (all) => {
      const idxs = all ? src.slides.map((_, i) => i) : Array.from(picked).sort((a, b) => a - b);
      if (!idxs.length) { ui.msg('Select at least one slide.'); return false; }
      L.app.insertSlidesFrom(src, idxs, keep.input.checked);
    };
    ui.dialog({ title: 'Slide Finder', body: h('div', { class: 'col' }, h('div', { text: L.t('Presentation: {0}', files[0].name), translate: 'no' }), grid, keep), width: 640, buttons: [{ label: '&Insert', primary: true, onClick: () => doInsert(false) }, { label: 'Insert &All', onClick: () => doInsert(true) }, { label: 'Close' }] });
  };

  /* ---------- Custom Animation effect options ---------- */
  D.effectOptions = function (a, onDone) {
    const info = L.anim.info(a.cls, a.eff) || {};
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      const shape = M.shapeById(E().slide(), a.sid);
      reactEffectOptions = {
        title: L.anim.name(a), directions: (info.dirs || []).map(([value, label]) => [value, L.t(label)]),
        dir: a.dir || info.dirs?.[0]?.[0] || '', amount: a.amount,
        amountEnabled: a.cls === 'emph' && (a.eff === 'growShrink' || a.eff === 'spin'), unit: a.eff === 'spin' ? '°' : '%',
        start: a.start || 'click', delay: a.delay || 0,
        speeds: L.anim.SPEEDS.map(([name, value]) => ({ name: L.t(name), value, label: L.t('{0} ({1} seconds)', L.t(name), value / 1000) })),
        dur: L.anim.SPEEDS.reduce((best, speed) => Math.abs(speed[1] - (a.dur || 500)) < Math.abs(best - (a.dur || 500)) ? speed[1] : best, 500),
        repeat: a.repeat || 1, by: a.by || 'all', hasText: !!(shape?.tx && shape.tx.ps.length >= 2),
      };
      reactEffectOptionsCallback = onDone;
      L.bus.emit('pptx-effect-options-dialog');
      return;
    }
    const dir = info.dirs ? ui.select(info.dirs.map((d) => [d[0], d[1]]), a.dir || info.dirs[0][0]) : null;
    const amt = a.cls === 'emph' && (a.eff === 'growShrink' || a.eff === 'spin') ? ui.spin({ value: a.amount || (a.eff === 'spin' ? 360 : 150), min: 10, max: 1440, step: 10, unit: a.eff === 'spin' ? '°' : '%', dec: 0 }) : null;
    const start = ui.select([['click', 'On Click'], ['with', 'With Previous'], ['after', 'After Previous']], a.start);
    const delay = ui.spin({ value: (a.delay || 0) / 1000, min: 0, max: 60, step: 0.5, fmt: (v) => L.t('{0} seconds', L.round(v, 1)) });
    const speed = ui.select(L.anim.SPEEDS.map(([n, ms]) => [ms, L.t('{0} ({1} seconds)', L.t(n), ms / 1000)]), L.anim.SPEEDS.reduce((best, s) => (Math.abs(s[1] - (a.dur || 500)) < Math.abs(best - (a.dur || 500)) ? s[1] : best), 500));
    const repeat = ui.select([[1, '(none)'], [2, '2'], [3, '3'], [4, '4'], [5, '5'], [10, '10']], a.repeat || 1);
    const sh = M.shapeById(E().slide(), a.sid);
    const grp = ui.select([['all', 'As One Object'], ['para', 'By 1st Level Paragraphs']], a.by || 'all');
    if (!sh || !sh.tx || sh.tx.ps.length < 2) grp.disabled = true;
    const tabs = ui.tabs([
      { label: 'Effect', body: h('div', { class: 'col' }, ui.group('Settings', dir ? ui.field('&Direction:', dir) : h('div', { class: 'tp-note', text: 'This effect has no direction settings.' }), amt ? ui.field(a.eff === 'spin' ? '&Amount:' : '&Size:', amt) : null)) },
      { label: 'Timing', body: h('div', { class: 'col' }, ui.field('&Start:', start), ui.field('&Delay:', delay), ui.field('Sp&eed:', speed), ui.field('&Repeat:', repeat)) },
      { label: 'Text Animation', body: h('div', { class: 'col' }, ui.field('&Group text:', grp, { cls: 'wide' })) },
    ]);
    ui.dialog({ title: L.anim.name(a), body: tabs, width: 380, buttons: [{ label: 'OK', primary: true, onClick: () => onDone({ dir: dir ? dir.value : a.dir, amount: amt ? amt.get() : a.amount, start: start.value, delay: delay.get() * 1000, dur: +speed.value, repeat: +repeat.value, by: grp.value }) }, { label: 'Cancel' }] });
  };

  /* ---------- Print ---------- */
  D.print = function () {
    const n = L.pres.slides.length;
    if (window.__VIBE_OFFICE_DIRECT_HOST__) {
      reactPrint = { slides: n, what: L.app.printOpts.what, per: L.app.printOpts.per, color: L.app.printOpts.color, frame: !!L.app.printOpts.frame, hidden: !!L.app.printOpts.hidden, range: `1-${n}`, rangeMode: 'all' };
      L.bus.emit('pptx-print-dialog');
      return;
    }
    const what = ui.select([['slides', 'Slides'], ['handouts', 'Handouts'], ['notes', 'Notes Pages'], ['outline', 'Outline View']], L.app.printOpts.what);
    const per = ui.select([[1, '1'], [2, '2'], [3, '3'], [4, '4'], [6, '6'], [9, '9']], L.app.printOpts.per);
    const color = ui.select([['color', 'Color'], ['gray', 'Grayscale'], ['bw', 'Pure Black and White']], L.app.printOpts.color);
    const frame = ui.check('Fra&me slides', L.app.printOpts.frame);
    const hid = ui.check('Print &hidden slides', L.app.printOpts.hidden);
    const all = ui.radio('pr-r', '&All', true), cur = ui.radio('pr-r', 'Curr&ent slide', false), rng = ui.radio('pr-r', '&Slides:', false);
    const rIn = h('input', { type: 'text', id: 'pr-range', value: `1-${n}`, style: 'width:120px' });
    const body = h('div', { class: 'col' }, ui.group('Print range', all, cur, h('div', { class: 'row', style: 'align-items:center' }, rng, rIn)), h('div', { class: 'row' }, ui.field('Print &what:', what), ui.field('Slides per page:', per)), ui.field('Color/gra&yscale:', color), frame, hid);
    ui.dialog({
      title: 'Print', body, width: 420, buttons: [{
        label: 'Preview', primary: true, onClick: () => {
          const range = all.input.checked ? null : cur.input.checked ? [E().idx] : parseRange(rIn.value, n);
          Object.assign(L.app.printOpts, { what: what.value, per: +per.value, color: color.value, frame: frame.input.checked, hidden: hid.input.checked, range });
          L.app.setView('preview');
        },
      }, { label: 'Cancel' }],
    });
  };
  function parseRange(s, n) {
    const out = [];
    for (const part of String(s).split(/[,;]/)) {
      const m = /^\s*(\d+)\s*(?:-\s*(\d+))?\s*$/.exec(part);
      if (!m) continue;
      const a = +m[1], b = m[2] ? +m[2] : a;
      for (let i = Math.min(a, b); i <= Math.max(a, b); i++) if (i >= 1 && i <= n) out.push(i - 1);
    }
    return out.length ? Array.from(new Set(out)) : null;
  }

  /* ---------- About ---------- */
  D.about = () => ui.about({ name: 'PPTX', text: [
    L.N('A presentation editor that runs entirely in your browser. Presentations are read and written as Office Open XML (.pptx, .pptm); nothing is uploaded.'),
    'Slide shows play in the browser, and presentations export to PDF, PNG pictures and web pages.'] });
})();
