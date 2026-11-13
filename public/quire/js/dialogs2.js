/* Wordx — dialogs, part 2: Insert menu (Break, Page Numbers, Date and Time, Field, Symbol,
 * Footnote, Caption, Cross-reference, Index and Tables, Bookmark, Hyperlink, Object, Diagram) and
 * the Table menu dialogs.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const { h } = L;
  const ui = L.ui;
  const G = L.dlg;
  const doc = () => D.doc;
  const A = () => L.app;
  const F = () => L.fields;

  /* ================= Break ================= */
  G.breakDlg = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const items = [['page', '&Page break'], ['column', '&Column break'], ['wrap', 'Text &wrapping break'], ['nextPage', '&Next page'], ['continuous', 'Con&tinuous'], ['evenPage', '&Even page'], ['oddPage', '&Odd page']];
    const r = direct ? null : G.radios('br', items, 'page');
    const els = r?.querySelectorAll('.ck');
    const root = h('div', { id: 'react-wordx-break-root' });
    const insert = (v) => {
      if (v === 'page' || v === 'column') { E.enter(v); return; }
      if (v === 'wrap') { E.enter('line'); return; }
      A().insertSectionBreak(v);
    };
    G.reactBreakCatalog = () => items.map(([value, label]) => ({ value, label: L.t(label.replace('&', '')) }));
    G.reactBreakAction = (value) => insert(value);
    const body = direct ? root : h('div', { class: 'col' }, ui.group('Break types', els[0], els[1], els[2]), ui.group('Section break types', els[3], els[4], els[5], els[6]));
    ui.dialog({ title: 'Break', width: 260, body, buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (direct) { G.reactBreakAction(root.querySelector('[data-break-type]:checked')?.getAttribute('data-break-type') || 'page'); return; }
      insert(r.get());
    } }, { label: 'Cancel' }] });
  };
  A().insertSectionBreak = function (type) {
    const d = doc();
    if (!E.sel || D.storyOf(d, E.sel.f.p) !== d.main || D.cellOf(d, E.sel.f.p)) { A().status('Section breaks can only be inserted in the main text, outside tables.'); return; }
    E.edit('Section Break', () => {
      const pos = E.deleteSelection();
      const cur = D.sectFor(d, pos.p);
      const np = O.splitPara(pos);
      D.touch(pos.p);
      pos.p.sect = L.opc.duplicate(cur);
      /* the new section (after the break) carries the requested start type */
      const after = A().curSect ? null : null; void after;
      const tb = D.topBlock(d, np.p);
      const sec = D.sectOfBlock(d, tb.i);
      const holder = sec.endPara || d;
      D.touchKey(holder, 'sect');
      (holder === d ? d.sect : holder.sect).type = type;
      return np;
    });
  };

  /* ================= Page Numbers ================= */
  G.pageNumbers = function (mode) {
    const d = doc();
    const s = A().curSect();
    if (mode === 'format') { formatPageNumbers(s); return; }
    const direct = !!(window.VO && VO.reactDirectHost);
    const pos = direct ? null : ui.select([['bottom', 'Bottom of page (Footer)'], ['top', 'Top of page (Header)']], 'bottom', () => draw(), { id: 'pn-pos' });
    const al = direct ? null : ui.select([['right', 'Right'], ['left', 'Left'], ['center', 'Center'], ['inside', 'Inside'], ['outside', 'Outside']], 'right', () => draw(), { id: 'pn-al' });
    const first = direct ? null : ui.check('&Show number on first page', true, null);
    const pv = h('div', { class: 'ps-page', style: 'width:80px;height:104px;margin:4px auto' });
    const root = h('div', { id: 'react-wordx-page-numbers-root' });
    const draw = () => { L.clear(pv); const y = pos.value === 'top' ? 'top:6px' : 'bottom:6px'; const x = { left: 'left:8px', right: 'right:8px', center: 'left:36px', inside: 'right:8px', outside: 'right:8px' }[al.value]; pv.appendChild(h('span', { style: `position:absolute;${y};${x};width:8px;height:8px;border:1px solid #000;background:#fff` })); };
    if (!direct) draw();
    G.reactPageNumberAction = (position, alignment, showFirst) => {
      const kind = position === 'top' ? 'hdr' : 'ftr';
      E.edit('Page Numbers', () => {
        D.touchKey(d, 'hf');
        const holder = s.holder;
        D.touchKey(holder, 'sect');
        const sect = holder === d ? d.sect : holder.sect;
        sect.refs = sect.refs || { hdr: {}, ftr: {} };
        sect.refs[kind] = Object.assign({}, sect.refs[kind]);
        const add = (ty, jc) => {
          let story = sect.refs[kind][ty] && d.hf[sect.refs[kind][ty]];
          if (!story) { const id = (kind === 'hdr' ? 'h' : 'f') + D.nid(); story = d.hf[id] = { kind, id, blocks: [D.para([], { style: D.useStyle(d, kind === 'hdr' ? 'Header' : 'Footer') })] }; sect.refs[kind][ty] = id; }
          const fid = D.newFid(), pn = D.useStyle(d, 'PageNumber');
          story.blocks.push(D.para([D.item('fb', { fid, instr: ' PAGE ' }, { style: pn }), D.item('fs', { fid }, { style: pn }), D.text('1', { style: pn }), D.item('fe', { fid }, { style: pn })], { style: D.useStyle(d, kind === 'hdr' ? 'Header' : 'Footer'), jc }));
        };
        const jc = (v, even) => (v === 'inside' ? (even ? 'right' : 'left') : v === 'outside' ? (even ? 'left' : 'right') : v);
        add('default', jc(alignment, false));
        if (d.settings.evenOdd) add('even', jc(alignment, true));
        sect.titlePg = !showFirst || sect.titlePg;
        if (!showFirst) { if (!sect.refs[kind].first) { const id = (kind === 'hdr' ? 'h' : 'f') + D.nid(); d.hf[id] = { kind, id, blocks: [D.para([], { style: kind === 'hdr' ? 'Header' : 'Footer' })] }; sect.refs[kind].first = id; } }
        else if (sect.titlePg && sect.refs[kind].first) add('first', jc(alignment, false));
        d._idxDirty = true;
        return E.sel;
      });
    };
    ui.dialog({
      title: 'Page Numbers', width: 360, body: direct ? root : G.row(h('div', { class: 'col' }, G.f('&Position:', pos), G.f('&Alignment:', al), first), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Preview' }), pv)),
      buttons: [{ label: '&Format...', onClick: () => { formatPageNumbers(s); return false; } }, { label: 'OK', primary: true, onClick: () => {
        const get = (id, fallback) => root.querySelector(`#${id}`)?.value || fallback;
        const showFirst = direct ? !!root.querySelector('input[type=checkbox]')?.checked : first.input.checked;
        G.reactPageNumberAction(direct ? get('react-wordx-page-position', 'bottom') : pos.value, direct ? get('react-wordx-page-alignment', 'right') : al.value, showFirst);
      } }, { label: 'Cancel' }],
      onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left'),
    });
  };
  function formatPageNumbers(s) {
    const d = doc();
    const sect = s.sect;
    const pn = Object.assign({ fmt: 'decimal', start: null }, sect.pgNum || {});
    const direct = !!(window.VO && VO.reactDirectHost);
    const formats = [['decimal', '1, 2, 3, ...'], ['numberInDash', '- 1 -, - 2 -, - 3 -, ...'], ['lowerLetter', 'a, b, c, ...'], ['upperLetter', 'A, B, C, ...'], ['lowerRoman', 'i, ii, iii, ...'], ['upperRoman', 'I, II, III, ...']];
    const fmt = direct ? null : ui.select(formats, pn.fmt, null, { id: 'pf-fmt' });
    const chap = direct ? null : ui.check('Include chapter &number', !!pn.chapStyle, null);
    const mode = direct ? null : G.radios('pf-m', [['cont', '&Continue from previous section'], ['start', 'Start &at:']], pn.start != null ? 'start' : 'cont');
    const startAt = direct ? null : G.num(pn.start != null ? pn.start : 1, { min: 0, id: 'pf-start' });
    const root = h('div', { id: 'react-wordx-page-number-format-root' });
    G.reactPageNumberFormatState = () => ({ fmt: pn.fmt, chapter: !!pn.chapStyle, mode: pn.start != null ? 'start' : 'cont', start: pn.start != null ? pn.start : 1 });
    G.reactPageNumberFormatCatalog = () => formats.map(([value, label]) => ({ value, label: L.t(label) }));
    G.reactPageNumberFormatAction = (format, chapter, numberingMode, start) => E.edit('Page Number Format', () => {
        D.touchKey(s.holder, 'sect');
        const sc = s.holder === d ? d.sect : s.holder.sect;
        sc.pgNum = Object.assign({}, sc.pgNum, { fmt: format, start: numberingMode === 'start' ? start : null });
        if (chapter !== !!pn.chapStyle) {
          if (chapter) sc.pgNum.chapStyle = 1;
          else { delete sc.pgNum.chapStyle; delete sc.pgNum.chapSep; }
        }
        return E.sel;
      });
    ui.dialog({ title: 'Page Number Format', width: 330, body: direct ? root : h('div', { class: 'col' }, G.f('Number &format:', fmt, 'wide'), chap, ui.group('Page numbering', mode, startAt)), buttons: [{ label: 'OK', primary: true, onClick: () => {
      const get = (id, fallback) => root.querySelector(`#${id}`)?.value || fallback;
      const directMode = root.querySelector('[data-page-number-mode]:checked')?.getAttribute('data-page-number-mode') || 'cont';
      G.reactPageNumberFormatAction(direct ? get('react-wordx-page-number-format', pn.fmt) : fmt.value, direct ? !!root.querySelector('input[type=checkbox]')?.checked : chap.input.checked, direct ? directMode : mode.get(), direct ? +(root.querySelector('#react-wordx-page-number-start')?.value || 1) : startAt.get());
    } }, { label: 'Cancel' }] });
  }

  /* ================= Date and Time ================= */
  G.dateTime = function () {
    const now = new Date();
    const pics = F().DATE_PICTURES;
    const def = L.store.get('defDatePic', pics[0]);
    const direct = !!(window.VO && VO.reactDirectHost);
    const lb = direct ? null : G.listBox(pics.map((p) => ({ value: p, label: F().formatDate(now, p) })), def, { height: 220, translate: 'no', onDbl: () => ok.click() });
    const lang = ui.select([['en-US', 'English (U.S.)'], ['en-GB', 'English (U.K.)']], 'en-US', null, { id: 'dt-lang' });
    const auto = ui.check('&Update automatically', !!L.store.get('defDateAuto', false), null);
    const root = h('div', { id: 'react-wordx-date-time-root' });
    G.reactDateTimeCatalog = () => pics.map((value) => ({ value, label: F().formatDate(now, value) }));
    G.reactDateTimeState = () => ({ format: def, autoUpdate: !!L.store.get('defDateAuto', false) });
    G.reactDateTimeAction = (action, pic, autoUpdate) => {
      if (action === 'default') { L.store.set('defDatePic', pic); L.store.set('defDateAuto', !!autoUpdate); A().status('Default date format saved.'); return; }
      if (action === 'insert') {
        if (autoUpdate) A().insertFieldNow(`${/[hHms]/.test(pic) && !/[dMy]/.test(pic) ? 'TIME' : 'DATE'} \\@ "${pic}"`, 'Insert Date and Time');
        else E.insertTextAt(F().formatDate(new Date(), pic), 'Insert Date and Time');
      }
    };
    const current = () => direct ? { pic: root.querySelector('select')?.value || def, autoUpdate: !!root.querySelector('input[type=checkbox]')?.checked } : { pic: lb.get(), autoUpdate: auto.input.checked };
    const dlg = ui.dialog({ title: 'Date and Time', width: 360, body: direct ? root : h('div', { class: 'col' }, h('label', { text: 'Available formats:' }), lb, G.f('&Language:', lang), auto), buttons: [
      { label: '&Default...', onClick: () => { const { pic, autoUpdate } = current(); if (direct) G.reactDateTimeAction('default', pic, autoUpdate); else G.reactDateTimeAction('default', pic, autoUpdate); return false; } },
      { label: 'OK', primary: true, onClick: () => { const { pic, autoUpdate } = current(); G.reactDateTimeAction('insert', pic, autoUpdate); } },
      { label: 'Cancel' }], onOpen: (el) => el.querySelector('.dlg-foot .btn').classList.add('left') });
    const ok = dlg.buttons[1];
  };

  /* ================= Field ================= */
  const FIELD_CATS = {
    [L.N('Date and Time')]: ['CREATEDATE', 'DATE', 'EDITTIME', 'PRINTDATE', 'SAVEDATE', 'TIME'],
    [L.N('Document Information')]: ['AUTHOR', 'COMMENTS', 'DOCPROPERTY', 'FILENAME', 'FILESIZE', 'INFO', 'KEYWORDS', 'LASTSAVEDBY', 'NUMCHARS', 'NUMPAGES', 'NUMWORDS', 'SUBJECT', 'TEMPLATE', 'TITLE'],
    [L.N('Equations and Formulas')]: ['=', 'ADVANCE', 'EQ', 'SYMBOL'],
    [L.N('Index and Tables')]: ['INDEX', 'RD', 'TA', 'TC', 'TOA', 'TOC', 'XE'],
    [L.N('Links and References')]: ['AUTOTEXT', 'HYPERLINK', 'INCLUDEPICTURE', 'INCLUDETEXT', 'NOTEREF', 'PAGEREF', 'QUOTE', 'REF', 'STYLEREF'],
    [L.N('Mail Merge')]: ['ASK', 'COMPARE', 'DATABASE', 'FILLIN', 'IF', 'MERGEFIELD', 'MERGEREC', 'MERGESEQ', 'NEXT', 'NEXTIF', 'SET', 'SKIPIF'],
    [L.N('Numbering')]: ['AUTONUM', 'AUTONUMLGL', 'AUTONUMOUT', 'BARCODE', 'LISTNUM', 'PAGE', 'REVNUM', 'SECTION', 'SECTIONPAGES', 'SEQ'],
    [L.N('User Information')]: ['USERADDRESS', 'USERINITIALS', 'USERNAME'],
  };
  const FIELD_DESC = { DATE: "Today's date", TIME: 'The current time', CREATEDATE: 'The date the document was created', SAVEDATE: 'The date the document was last saved', PRINTDATE: 'The date the document was last printed', EDITTIME: 'The total document editing time', AUTHOR: 'The name of the document author from Document Properties', TITLE: "The document's title from Document Properties", SUBJECT: "The document's subject", KEYWORDS: "The document's keywords", COMMENTS: 'The comments from Document Properties', FILENAME: "The document's name", NUMPAGES: 'The number of pages in the document', NUMWORDS: 'The number of words in the document', NUMCHARS: 'The number of characters in the document', PAGE: 'The number of the current page', SECTION: 'The number of the current section', SECTIONPAGES: 'The number of pages in the section', SEQ: 'Inserts an automatic sequence number', REF: 'Inserts the text marked by a bookmark', PAGEREF: 'Inserts the number of the page containing a bookmark', NOTEREF: 'Inserts the mark of a footnote or endnote', STYLEREF: 'Inserts the text from a like-style paragraph', TOC: 'Builds a table of contents', INDEX: 'Builds an index', XE: 'Marks an index entry', TC: 'Marks a table of contents entry', '=': 'Calculates the result of an expression', MERGEFIELD: 'Inserts a mail merge field', USERNAME: 'Your name from User Information', USERINITIALS: 'Your initials from User Information', USERADDRESS: 'Your postal address from User Information', LASTSAVEDBY: 'Name of the person who last saved the document', HYPERLINK: 'Opens and jumps to the specified file', QUOTE: 'Inserts literal text', SYMBOL: 'Inserts a symbol', IF: 'Evaluates arguments conditionally', REVNUM: "The document's revision number", TEMPLATE: "The document's template name", DOCPROPERTY: 'The value of a document property', MACROBUTTON: 'Runs a macro' };
  G.field = function (fieldRef, opts) {
    opts = opts || {};
    const editing = !!fieldRef;
    const direct = !!(window.VO && VO.reactDirectHost);
    if (direct) {
      const root = h('div', { id: 'react-wordx-field-root' });
      const initialCode = editing ? fieldRef.it.instr.trim().replace(/\s*\\\*\s*MERGEFORMAT/i, '') : '';
      const initialType = editing ? F().parse(fieldRef.it.instr).type : '';
      const fullCode = (code, format, datePicture, keep) => {
        let value = code.trim();
        if (format) value += ' \\* ' + format;
        if (datePicture && /^(DATE|TIME|CREATEDATE|SAVEDATE|PRINTDATE)\b/i.test(value) && !/\\@/.test(value)) value += ' \\@ "' + datePicture + '"';
        if (keep && !/MERGEFORMAT/i.test(value) && !/^(TOC|INDEX|=)/i.test(value)) value += ' \\* MERGEFORMAT';
        return value;
      };
      G.reactFieldState = () => ({ code: initialCode, type: initialType, editing, format: '', datePicture: '', keep: true });
      G.reactFieldCatalog = (category = 'all') => ({
        categories: [{ value: 'all', label: L.t('(All)') }, ...Object.keys(FIELD_CATS).map((value) => ({ value, label: value }))],
        names: (category === 'all' ? Array.from(new Set(Object.values(FIELD_CATS).flat())).sort() : FIELD_CATS[category] || []).map((value) => ({ value, label: value })),
        formats: [['', '(none)'], ['Upper', 'Uppercase'], ['Lower', 'Lowercase'], ['FirstCap', 'First capital'], ['Caps', 'Title case'], ['Arabic', '1, 2, 3, ...'], ['roman', 'i, ii, iii, ...'], ['ROMAN', 'I, II, III, ...'], ['alphabetic', 'a, b, c, ...'], ['ALPHABETIC', 'A, B, C, ...'], ['Ordinal', '1st, 2nd, 3rd ...'], ['CardText', 'One, Two, Three'], ['OrdText', 'First, Second, Third']].map(([value, label]) => ({ value, label: L.t(label) })),
        datePictures: ['', ...F().DATE_PICTURES].map((value) => ({ value, label: value || L.t('(default)') })),
      });
      G.reactFieldDescription = (name) => L.t(FIELD_DESC[name] || '');
      G.reactFieldPreview = (code, format, datePicture, keep) => {
        try { const result = F().evaluate(fullCode(code, format, datePicture, keep), E.sel ? E.sel.f : null); return result == null ? L.t('(the result is produced when the field is inserted or updated)') : Array.isArray(result) ? L.t('(generated paragraphs)') : String(result); }
        catch (e) { return L.t('Error!'); }
      };
      G.reactFieldAction = (code, format, datePicture, keep) => {
        const value = fullCode(code, format, datePicture, keep);
        if (!value) return false;
        const type = F().parse(value).type;
        if (editing) F().setInstr(fieldRef, value);
        else if (type === 'TOC' || type === 'INDEX') F().insertGenerated(value);
        else A().insertFieldNow(value);
      };
      ui.dialog({ title: 'Field', width: 560, focus: opts.codes ? root : null, body: root, buttons: [{ label: 'OK', primary: true, onClick: () => {
        const get = (id) => root.querySelector('#' + id);
        return G.reactFieldAction(get('react-wordx-field-code')?.value || '', get('react-wordx-field-format')?.value || '', get('react-wordx-field-date-picture')?.value || '', !!get('react-wordx-field-keep')?.checked);
      } }, { label: 'Cancel' }] });
      return;
    }
    const cat = ui.select([['all', '(All)']].concat(Object.keys(FIELD_CATS).map((k) => [k, k])), 'all', () => fillNames(), { id: 'fl-cat' });
    const names = G.listBox([], null, { height: 200, translate: 'no', onChange: (v) => { desc.textContent = L.t(FIELD_DESC[v] || ''); code.value = v === '=' ? '= ' : v + ' '; preview(); } });
    const fillNames = () => { const list = cat.value === 'all' ? Array.from(new Set(Object.values(FIELD_CATS).flat())).sort() : FIELD_CATS[cat.value]; names.setItems(list); };
    fillNames();
    const code = h('input', { type: 'text', id: 'fl-code', style: 'width:100%' });
    const desc = h('div', { class: 'tp-note', style: 'min-height:28px' });
    const res = h('div', { class: 'tp-box', style: 'min-height:22px', translate: 'no' });   // the field's result: the document's text
    const fmtSel = ui.select([['', '(none)'], ['Upper', 'Uppercase'], ['Lower', 'Lowercase'], ['FirstCap', 'First capital'], ['Caps', 'Title case'], ['Arabic', '1, 2, 3, ...'], ['roman', 'i, ii, iii, ...'], ['ROMAN', 'I, II, III, ...'], ['alphabetic', 'a, b, c, ...'], ['ALPHABETIC', 'A, B, C, ...'], ['Ordinal', '1st, 2nd, 3rd ...'], ['CardText', 'One, Two, Three'], ['OrdText', 'First, Second, Third']], '', () => preview(), { id: 'fl-fmt' });
    const datePic = ui.select([['', L.t('(default)')]].concat(F().DATE_PICTURES.map((p) => [p, p])), '', () => preview(), { id: 'fl-date', translate: 'no' });
    const keep = ui.check('Preserve formatting during updates', true, null);
    const fullCode = () => { let c = code.value.trim(); if (fmtSel.value) c += ' \\* ' + fmtSel.value; if (datePic.value && /^(DATE|TIME|CREATEDATE|SAVEDATE|PRINTDATE)\b/i.test(c) && !/\\@/.test(c)) c += ` \\@ "${datePic.value}"`; if (keep.input.checked && !/MERGEFORMAT/i.test(c) && !/^(TOC|INDEX|=)/i.test(c)) c += ' \\* MERGEFORMAT'; return c; };
    const preview = () => { try { const r = F().evaluate(fullCode(), E.sel ? E.sel.f : null); res.textContent = r == null ? L.t('(the result is produced when the field is inserted or updated)') : Array.isArray(r) ? L.t('(generated paragraphs)') : String(r); } catch (e) { res.textContent = L.t('Error!'); } };
    if (editing) { code.value = fieldRef.it.instr.trim().replace(/\s*\\\*\s*MERGEFORMAT/i, ''); const t = F().parse(fieldRef.it.instr).type; names.set(t); desc.textContent = L.t(FIELD_DESC[t] || ''); }
    else if (opts.codes) code.value = '';
    code.addEventListener('input', preview);
    preview();
    ui.dialog({
      title: 'Field', width: 560, focus: opts.codes ? code : null,
      body: G.row(h('div', { class: 'col', style: 'width:190px' }, G.f('&Categories:', cat), h('label', { text: 'Field names:' }), names), h('div', { class: 'col', style: 'flex:1;min-width:220px' }, h('div', { class: 'cd-cap', text: 'Field codes:' }), code, desc, ui.group('Field options', G.f('&Format:', fmtSel), G.f('Date &picture:', datePic)), h('div', { class: 'cd-cap', text: 'Result preview' }), res, keep)),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const c = fullCode();
        if (!c) return false;
        const t = F().parse(c).type;
        if (editing) { F().setInstr(fieldRef, c); return; }
        if (t === 'TOC' || t === 'INDEX') { F().insertGenerated(c); return; }
        A().insertFieldNow(c);
      } }, { label: 'Cancel' }],
    });
  };

  /* ================= Symbol ================= */
  const SPECIALS = [['—', L.N('Em Dash'), 'Alt+Ctrl+Num -'], ['–', L.N('En Dash'), 'Ctrl+Num -'], ['‑', L.N('Nonbreaking Hyphen'), 'Ctrl+_'], ['­', L.N('Optional Hyphen'), 'Ctrl+-'], [' ', L.N('Em Space'), ''], [' ', L.N('En Space'), ''], [' ', L.N('1/4 Em Space'), ''], [' ', L.N('Nonbreaking Space'), 'Ctrl+Shift+Space'], ['©', L.N('Copyright'), 'Alt+Ctrl+C'], ['®', L.N('Registered'), 'Alt+Ctrl+R'], ['™', L.N('Trademark'), 'Alt+Ctrl+T'], ['§', L.N('Section'), ''], ['¶', L.N('Paragraph'), ''], ['…', L.N('Ellipsis'), 'Alt+Ctrl+.'], ['‘', L.N('Single Opening Quote'), "Ctrl+`,`"], ['’', L.N('Single Closing Quote'), "Ctrl+','"], ['“', L.N('Double Opening Quote'), 'Ctrl+`,"'], ['”', L.N('Double Closing Quote'), "Ctrl+',\""], ['​', L.N('No-Width Optional Break'), ''], ['⁠', L.N('No-Width Non Break'), '']];
  const SUBSETS = [[L.N('Basic Latin'), 0x20, 0x7e], [L.N('Latin-1 Supplement'), 0xa0, 0xff], [L.N('Latin Extended-A'), 0x100, 0x17f], [L.N('Greek and Coptic'), 0x370, 0x3ff], [L.N('Cyrillic'), 0x400, 0x4ff], [L.N('General Punctuation'), 0x2010, 0x205e], [L.N('Currency Symbols'), 0x20a0, 0x20bf], [L.N('Letterlike Symbols'), 0x2100, 0x214f], [L.N('Number Forms'), 0x2150, 0x218b], [L.N('Arrows'), 0x2190, 0x21ff], [L.N('Mathematical Operators'), 0x2200, 0x22ff], [L.N('Miscellaneous Technical'), 0x2300, 0x23ff], [L.N('Box Drawing'), 0x2500, 0x257f], [L.N('Block Elements'), 0x2580, 0x259f], [L.N('Geometric Shapes'), 0x25a0, 0x25ff], [L.N('Miscellaneous Symbols'), 0x2600, 0x26ff], [L.N('Dingbats'), 0x2700, 0x27bf]];
  G.symbol = function (opts) {
    opts = opts || {};
    const direct = !!(window.VO && VO.reactDirectHost);
    let cur = null;
    const recent = L.store.get('recentSymbols', ['€', '£', '¥', '©', '®', '™', '±', '≠', '≤', '≥', '÷', '×', '∞', 'µ', 'α', 'β']);
    const font = ui.select([['', L.t('(normal text)')], ['Symbol', 'Symbol'], ['Wingdings', 'Wingdings'], ['Webdings', 'Webdings']].concat(G.FONTS().map((f) => [f, f])), '', () => fill(), { id: 'sy-font', translate: 'no' });
    const subset = ui.select(SUBSETS.map((s, i) => [i, s[0]]), 1, () => fill(), { id: 'sy-sub' });
    const grid = h('div', { class: 'sym-grid' });
    const codeIn = h('input', { type: 'text', id: 'sy-code', style: 'width:70px' });
    const name = h('div', { class: 'tp-note', style: 'min-height:16px' });
    const fill = () => {
      L.clear(grid);
      const f = font.value;
      const chars = [];
      if (f === 'Symbol' || f === 'Wingdings' || f === 'Webdings') for (let c = 0x21; c <= 0xff; c++) { if (c >= 0x7f && c < 0xa1) continue; const u = L.mapSymbolChar(String.fromCharCode(c), f); if (u) chars.push({ ch: u, raw: String.fromCharCode(c), font: f }); }
      else { const [, a, b] = SUBSETS[+subset.value]; for (let c = a; c <= b; c++) chars.push({ ch: String.fromCodePoint(c), font: f }); }
      for (const c of chars) {
        const b = h('button', { type: 'button', text: c.ch, style: c.font && !/Symbol|Wingdings|Webdings/.test(c.font) ? `font-family:${L.fontStack(c.font)}` : '' });
        b.addEventListener('click', () => { L.$$('button.on', grid).forEach((x) => x.classList.remove('on')); b.classList.add('on'); cur = c; codeIn.value = c.ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); name.textContent = L.t('Unicode character {0}', codeIn.value); });
        b.addEventListener('dblclick', () => { cur = c; insertBtn.click(); });
        grid.appendChild(b);
      }
    };
    fill();
    codeIn.addEventListener('change', () => { const n = parseInt(codeIn.value, 16); if (!isNaN(n)) { cur = { ch: String.fromCodePoint(n), font: '' }; name.textContent = L.t('Unicode character {0}', codeIn.value.toUpperCase()); } });
    const rec = h('div', { class: 'sym-recent' });
    for (const c of recent.slice(0, 16)) { const b = h('button', { type: 'button', text: c }); b.addEventListener('click', () => { cur = { ch: c, font: '' }; codeIn.value = c.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); }); b.addEventListener('dblclick', () => { cur = { ch: c, font: '' }; insertBtn.click(); }); rec.appendChild(b); }
    const spec = G.listBox(SPECIALS.map(([c, n, k]) => ({ value: c, label: L.t(n) + (k ? '   ' + ui.keyText(k) : '') })), SPECIALS[0][0], { height: 220, onDbl: () => insertBtn.click() });
    const tabs = ui.tabs([{ label: '&Symbols', body: h('div', { class: 'col' }, G.row(G.f('&Font:', font), G.f('Subset:', subset)), grid, h('div', { class: 'cd-cap', text: 'Recently used symbols:' }), rec, G.row(name, G.f('Character code:', codeIn))) }, { label: 'S&pecial Characters', body: spec }]);
    let tabIdx = 0;
    tabs.querySelectorAll('.tab').forEach((t, i) => t.addEventListener('click', () => { tabIdx = i; }));
    const doInsert = () => {
      const c = tabIdx === 1 ? { ch: spec.get(), font: '' } : cur;
      if (!c) return;
      if (opts.pick) { opts.pick(c.ch, c.font); return; }
      const list = [c.ch].concat(recent.filter((x) => x !== c.ch)).slice(0, 16);
      L.store.set('recentSymbols', list);
      if (c.raw && c.font) E.insertItem(D.item('sym', { font: c.font, char: c.raw, code: c.raw.charCodeAt(0).toString(16).toUpperCase() }, O.inheritRPr(E.sel.f)), 'Insert Symbol');
      else E.insertTextAt(c.ch, 'Insert Symbol');
    };
    let symbolState = { font: '', subset: 1, tab: 0, selected: null, code: '', recent: recent.slice(0, 16), special: SPECIALS[0][0] };
    const symbolCatalog = () => {
      const chars = [];
      if (symbolState.font === 'Symbol' || symbolState.font === 'Wingdings' || symbolState.font === 'Webdings') for (let c = 0x21; c <= 0xff; c++) { if (c >= 0x7f && c < 0xa1) continue; const raw = String.fromCharCode(c), ch = L.mapSymbolChar(raw, symbolState.font); if (ch) chars.push({ ch, raw, font: symbolState.font }); }
      else { const [, a, b] = SUBSETS[symbolState.subset]; for (let c = a; c <= b; c++) chars.push({ ch: String.fromCodePoint(c), font: symbolState.font }); }
      return { ...symbolState, fonts: ['', 'Symbol', 'Wingdings', 'Webdings', ...G.FONTS()], subsets: SUBSETS.map((s, i) => ({ value: i, label: s[0] })), chars: chars.map((c) => ({ ...c, code: c.ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0') })), specials: SPECIALS.map(([ch, label, key]) => ({ value: ch, label: L.t(label) + (key ? '   ' + ui.keyText(key) : '') })) };
    };
    G.reactSymbolCatalog = symbolCatalog;
    G.reactSymbolAction = (action, value) => {
      if (action === 'font') { symbolState.font = value; symbolState.selected = null; symbolState.code = ''; }
      else if (action === 'subset') { symbolState.subset = Number(value); symbolState.selected = null; symbolState.code = ''; }
      else if (action === 'tab') symbolState.tab = Number(value);
      else if (action === 'char') { cur = value; symbolState.selected = value.ch; symbolState.code = value.code; }
      else if (action === 'recent') { cur = { ch: value, font: '' }; symbolState.selected = value; symbolState.code = value.codePointAt(0).toString(16).toUpperCase().padStart(4, '0'); }
      else if (action === 'special') symbolState.special = value;
      else if (action === 'code') { symbolState.code = value; const n = parseInt(value, 16); if (!isNaN(n) && n <= 0x10ffff) { cur = { ch: String.fromCodePoint(n), font: '' }; symbolState.selected = cur.ch; } }
      else if (action === 'insert') { tabIdx = symbolState.tab; if (tabIdx === 1) spec.set(symbolState.special); doInsert(); if (!opts.pick) { symbolState.recent = L.store.get('recentSymbols', symbolState.recent); L.bus.emit('wordx-symbol-refresh'); } }
      return symbolCatalog();
    };
    const reactRoot = h('div', { id: 'react-wordx-symbol-root' });
    const dlg = ui.dialog({ title: 'Symbol', width: 520, body: direct ? reactRoot : tabs, buttons: [{ label: '&Insert', primary: true, onClick: () => { if (direct) G.reactSymbolAction('insert'); else doInsert(); if (opts.pick) return; return false; } }, { label: opts.pick ? 'Cancel' : 'Close' }] });
    const insertBtn = dlg.buttons[0];
  };

  /* ================= Footnote and Endnote ================= */
  G.footnote = function () {
    const d = doc();
    const direct = !!(window.VO && VO.reactDirectHost);
    if (direct) {
      const root = h('div', { id: 'react-wordx-footnote-root' });
      const state = () => ({
        kind: 'fn',
        fnPos: (d.settings.fnPr && d.settings.fnPr.pos) || 'pageBottom',
        enPos: (d.settings.enPr && d.settings.enPr.pos) || 'docEnd',
        fnFmt: (d.settings.fnPr && d.settings.fnPr.fmt) || 'decimal',
        enFmt: (d.settings.enPr && d.settings.enPr.fmt) || 'lowerRoman',
        fnStart: (d.settings.fnPr && d.settings.fnPr.start) || 1,
        enStart: (d.settings.enPr && d.settings.enPr.start) || 1,
        fnRestart: (d.settings.fnPr && d.settings.fnPr.restart) || 'continuous',
        enRestart: (d.settings.enPr && d.settings.enPr.restart) || 'continuous',
        custom: '',
      });
      G.reactFootnoteState = state;
      G.reactFootnoteCatalog = () => ({
        locations: { fn: [['pageBottom', 'Bottom of page'], ['beneathText', 'Below text']], en: [['docEnd', 'End of document'], ['sectEnd', 'End of section']] },
        formats: [['decimal', '1, 2, 3, ...'], ['lowerLetter', 'a, b, c, ...'], ['upperLetter', 'A, B, C, ...'], ['lowerRoman', 'i, ii, iii, ...'], ['upperRoman', 'I, II, III, ...'], ['chicago', '*, †, ‡, §, ...']],
        restarts: [['continuous', 'Continuous'], ['eachSect', 'Restart each section'], ['eachPage', 'Restart each page']],
      });
      G.reactFootnoteAction = (action, values) => {
        if (action === 'symbol') {
          G.symbol({ pick: (character) => { const input = root.querySelector('#react-wordx-footnote-custom'); if (input) { input.value = character; input.dispatchEvent(new Event('input', { bubbles: true })); } } });
          return;
        }
        if (action === 'convert') { convert(); return; }
        const kind = values.kind;
        const config = kind === 'fn' ? { fmt: values.fnFmt, start: values.fnStart, restart: values.fnRestart, pos: values.fnPos } : { fmt: values.enFmt, start: values.enStart, restart: values.enRestart, pos: values.enPos };
        L.notes.setOptions(kind, config);
        if (action === 'insert') L.notes.insert(kind, { mark: values.custom.trim() || null });
        else return false;
      };
      const values = () => Object.fromEntries(['kind', 'fnPos', 'enPos', 'fnFmt', 'enFmt', 'fnStart', 'enStart', 'fnRestart', 'enRestart', 'custom'].map((key) => [key, (key === 'kind' ? root.querySelector('[name=react-wordx-footnote-kind]:checked') : root.querySelector('#react-wordx-footnote-' + key))?.value || (key.endsWith('Start') ? 1 : '')]).map(([key, value]) => [key, key.endsWith('Start') ? Math.max(1, Number(value)) : value]));
      ui.dialog({ title: 'Footnote and Endnote', width: 380, body: root, buttons: [
        { label: '&Insert', primary: true, onClick: () => G.reactFootnoteAction('insert', values()) },
        { label: 'Cancel' },
        { label: '&Apply', onClick: () => G.reactFootnoteAction('apply', values()) },
      ] });
      return;
    }
    const kind = G.radios('fn-k', [['fn', '&Footnotes:'], ['en', '&Endnotes:']], 'fn', () => sync());
    const fnLoc = ui.select([['pageBottom', 'Bottom of page'], ['beneathText', 'Below text']], (d.settings.fnPr && d.settings.fnPr.pos) || 'pageBottom', null, { id: 'fn-floc' });
    const enLoc = ui.select([['docEnd', 'End of document'], ['sectEnd', 'End of section']], (d.settings.enPr && d.settings.enPr.pos) || 'docEnd', null, { id: 'fn-eloc' });
    const fmt = ui.select([['decimal', '1, 2, 3, ...'], ['lowerLetter', 'a, b, c, ...'], ['upperLetter', 'A, B, C, ...'], ['lowerRoman', 'i, ii, iii, ...'], ['upperRoman', 'I, II, III, ...'], ['chicago', '*, †, ‡, §, ...']], 'decimal', null, { id: 'fn-fmt' });
    const custom = h('input', { type: 'text', id: 'fn-custom', style: 'width:60px' });
    const startAt = G.num(1, { min: 1, id: 'fn-start' });
    const numbering = ui.select([['continuous', 'Continuous'], ['eachSect', 'Restart each section'], ['eachPage', 'Restart each page']], 'continuous', null, { id: 'fn-num' });
    const sync = () => { const k = kind.get(); const pr = (k === 'fn' ? d.settings.fnPr : d.settings.enPr) || {}; fmt.value = pr.fmt || (k === 'fn' ? 'decimal' : 'lowerRoman'); startAt.set(pr.start || 1); numbering.value = pr.restart || 'continuous'; };
    sync();
    const settings = () => ({ fmt: fmt.value, start: startAt.get(), restart: numbering.value, pos: kind.get() === 'fn' ? fnLoc.value : enLoc.value });
    const kEls = kind.querySelectorAll('.ck');
    ui.dialog({
      title: 'Footnote and Endnote', width: 380,
      body: h('div', { class: 'col' }, ui.group('Location', G.row(kEls[0], fnLoc), G.row(kEls[1], enLoc), ui.button('C&onvert...', () => convert(), { class: 'btn small' })), ui.group('Format', G.f('Number f&ormat:', fmt, 'wide'), G.row(G.f('Custom &mark:', custom, 'wide'), ui.button('S&ymbol...', () => G.symbol({ pick: (ch) => { custom.value = ch; } }), { class: 'btn small' })), G.f('&Start at:', startAt, 'wide'), G.f('Numbering:', numbering, 'wide'))),
      buttons: [{ label: '&Insert', primary: true, onClick: () => { L.notes.setOptions(kind.get(), settings()); L.notes.insert(kind.get(), { mark: custom.value.trim() || null }); } }, { label: 'Cancel' }, { label: '&Apply', onClick: () => { L.notes.setOptions(kind.get(), settings()); return false; } }],
    });
    function convert() {
      if (direct) {
        const convertRoot = h('div', { id: 'react-wordx-convert-notes-root' });
        ui.dialog({ title: 'Convert Notes', width: 300, body: convertRoot, buttons: [{ label: 'OK', primary: true, onClick: () => L.notes.convertAll(convertRoot.querySelector('[name=react-wordx-convert-notes]:checked')?.value || 'fn') }, { label: 'Cancel' }] });
        return;
      }
      const r = G.radios('fn-conv', [['fn', 'Convert all &footnotes to endnotes'], ['en', 'Convert all &endnotes to footnotes']], 'fn');
      ui.dialog({ title: 'Convert Notes', width: 300, body: r, buttons: [{ label: 'OK', primary: true, onClick: () => L.notes.convertAll(r.get()) }, { label: 'Cancel' }] });
    }
  };

  /* ================= Caption ================= */
  G.caption = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const labels = Array.from(new Set(F().CAPTION_LABELS.concat(L.store.get('captionLabels', []))));
    const ci = L.tables && L.tables.inTable();
    const lab = ui.select(labels.map((x) => [x, x]), ci ? 'Table' : 'Figure', () => upd(), { id: 'cp-lab', translate: 'no' });   // labels go into the document as they are
    const txt = h('input', { type: 'text', id: 'cp-text', style: 'width:100%' });
    const pos = ui.select([['below', 'Below selected item'], ['above', 'Above selected item']], ci ? 'above' : 'below', null, { id: 'cp-pos' });
    const excl = ui.check('&Exclude label from caption', false, () => upd());
    let num = { fmt: 'ARABIC', chapter: false, chapterStyle: 1, sep: '-' };
    let reactCaption = { labels, label: lab.value, position: pos.value, exclude: false, text: '', numbering: { ...num } };
    const captionPreview = (label, excluded, numbering) => { const n = L.D.fmtNum(1, { ARABIC: 'decimal', alphabetic: 'lowerLetter', ALPHABETIC: 'upperLetter', roman: 'lowerRoman', ROMAN: 'upperRoman' }[numbering.fmt]); return (excluded ? '' : label + ' ') + (numbering.chapter ? '1' + numbering.sep : '') + n; };
    const upd = () => { if (direct) { txt.value = captionPreview(reactCaption.label, reactCaption.exclude, reactCaption.numbering); return; } txt.value = captionPreview(lab.value, excl.input.checked, num); };
    upd();
    reactCaption.text = txt.value;
    G.reactCaptionCatalog = () => ({ ...reactCaption, labels: reactCaption.labels.slice() });
    G.reactCaptionAction = async (action, value) => {
      if (action === 'change') { reactCaption = { ...reactCaption, ...(value || {}) }; num = { ...reactCaption.numbering }; if (!direct) { lab.value = reactCaption.label; pos.value = reactCaption.position; excl.input.checked = reactCaption.exclude; txt.value = reactCaption.text; } }
      else if (action === 'newLabel') { const name = await ui.prompt('Label:', '', 'New Label'); if (!name) return; const list = L.store.get('captionLabels', []); if (!list.includes(name)) { list.push(name); L.store.set('captionLabels', list); } if (!direct) lab.value = name; const text = captionPreview(name, reactCaption.exclude, reactCaption.numbering); if (!direct) { upd(); } reactCaption = { ...reactCaption, labels: Array.from(new Set(F().CAPTION_LABELS.concat(list))), label: name, text: direct ? text : txt.value }; L.bus.emit('wordx-caption-refresh'); }
      else if (action === 'deleteLabel') { if (F().CAPTION_LABELS.includes(reactCaption.label)) return; const list = L.store.get('captionLabels', []).filter((item) => item !== reactCaption.label); L.store.set('captionLabels', list); const label = F().CAPTION_LABELS[0] || 'Figure'; if (!direct) lab.value = label; const text = captionPreview(label, reactCaption.exclude, reactCaption.numbering); if (!direct) { upd(); } reactCaption = { ...reactCaption, labels: Array.from(new Set(F().CAPTION_LABELS.concat(list))), label, text: direct ? text : txt.value }; L.bus.emit('wordx-caption-refresh'); }
      else if (action === 'numbering') numbering();
    };
    ui.dialog({
      title: 'Caption', width: 360, focus: direct ? null : txt,
      body: direct ? h('div', { id: 'react-wordx-caption-root' }) : h('div', { class: 'col' }, h('label', { for: 'cp-text', html: ui.ampHTML('&Caption:') }), txt, ui.group('Options', G.f('&Label:', lab), G.f('&Position:', pos), excl), h('div', { class: 'tp-btns' }, ui.button('&New Label...', async () => { const n = await ui.prompt('Label:', '', 'New Label'); if (!n) return; const list = L.store.get('captionLabels', []); if (!list.includes(n)) { list.push(n); L.store.set('captionLabels', list); } lab.appendChild(G.raw('option', { value: n }, n)); lab.value = n; upd(); }, { class: 'btn small' }), ui.button('&Delete Label', () => { if (F().CAPTION_LABELS.includes(lab.value)) return; L.store.set('captionLabels', L.store.get('captionLabels', []).filter((x) => x !== lab.value)); lab.querySelector(`option[value="${lab.value}"]`).remove(); upd(); }, { class: 'btn small' }), ui.button('N&umbering...', () => numbering(), { class: 'btn small' }))),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const label = direct ? reactCaption.label : lab.value;
        const excluded = direct ? reactCaption.exclude : excl.input.checked;
        const position = direct ? reactCaption.position : pos.value;
        const currentNum = direct ? reactCaption.numbering : num;
        const prefix = (excluded ? '' : label + ' ');
        let rest = direct ? reactCaption.text : txt.value;
        const m = new RegExp('^' + prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:[\\dA-Za-z]+' + (currentNum.chapter ? '[-.:—–]?[\\dA-Za-z]+' : '') + ')').exec(rest);
        rest = m ? rest.slice(m[0].length) : '';
        F().insertCaption({ label, position, text: rest, fmt: currentNum.fmt, chapter: currentNum.chapter, chapterStyle: currentNum.chapterStyle, chapterSep: currentNum.sep, noLabel: excluded });
      } }, { label: 'Cancel' }],
    });
    function numbering() {
      if (direct) { num = { ...reactCaption.numbering }; }
      let reactNum = { ...num };
      G.reactCaptionNumberingCatalog = () => ({ ...reactNum });
      G.reactCaptionNumberingAction = (value) => { reactNum = { ...reactNum, ...(value || {}) }; };
      const fmt = ui.select([['ARABIC', '1, 2, 3, ...'], ['alphabetic', 'a, b, c, ...'], ['ALPHABETIC', 'A, B, C, ...'], ['roman', 'i, ii, iii, ...'], ['ROMAN', 'I, II, III, ...']], num.fmt, null, { id: 'cn-fmt' });
      const chap = ui.check('Include &chapter number', num.chapter, null);
      const style = ui.select([1, 2, 3, 4, 5, 6, 7, 8, 9].map((i) => [i, 'Heading ' + i]), num.chapterStyle, null, { id: 'cn-st', translate: 'no' });   // style names
      const sep = ui.select([['-', '- (hyphen)'], ['.', '. (period)'], [':', ': (colon)'], ['—', '— (em dash)'], ['–', '– (en dash)']], num.sep, null, { id: 'cn-sep' });
      const reactRoot = h('div', { id: 'react-wordx-caption-numbering-root' });
      ui.dialog({ title: 'Caption Numbering', width: 320, body: direct ? reactRoot : h('div', { class: 'col' }, G.f('F&ormat:', fmt), chap, G.f('Chapter starts with style:', style, 'wide'), G.f('Use separator:', sep, 'wide')), buttons: [{ label: 'OK', primary: true, onClick: () => { num = direct ? reactNum : { fmt: fmt.value, chapter: chap.input.checked, chapterStyle: +style.value, sep: sep.value }; if (direct) { reactCaption = { ...reactCaption, numbering: { ...num }, text: captionPreview(reactCaption.label, reactCaption.exclude, num) }; L.bus.emit('wordx-caption-refresh'); } else upd(); } }, { label: 'Cancel' }] });
    }
  };

  /* ================= Cross-reference ================= */
  G.crossRef = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const kinds = [L.N('Numbered item'), L.N('Heading'), L.N('Bookmark'), L.N('Footnote'), L.N('Endnote')];
    const types = kinds.concat(['Equation', 'Figure', 'Table'], L.store.get('captionLabels', []));   // then the caption labels, as the document has them
    const what = { 'Numbered item': [['number', 'Paragraph number'], ['numberFull', 'Paragraph number (full context)'], ['text', 'Paragraph text'], ['page', 'Page number'], ['abovebelow', L.N('Above/below')]], Heading: [['text', 'Heading text'], ['page', 'Page number'], ['number', 'Heading number'], ['abovebelow', L.N('Above/below')]], Bookmark: [['text', 'Bookmark text'], ['page', 'Page number'], ['number', 'Paragraph number'], ['abovebelow', L.N('Above/below')]], Footnote: [['notenum', 'Footnote number'], ['page', 'Page number'], ['abovebelow', L.N('Above/below')]], Endnote: [['notenum', 'Endnote number'], ['page', 'Page number'], ['abovebelow', L.N('Above/below')]] };
    const capWhat = [['text', 'Entire caption'], ['labelnum', 'Only label and number'], ['page', 'Page number'], ['abovebelow', L.N('Above/below')]];
    const type = direct ? null : ui.select(types.map((t) => [t, kinds.includes(t) ? L.t(t) : t]), 'Heading', () => refill(), { id: 'xr-type', translate: 'no' });
    const ref = direct ? null : ui.select([], '', null, { id: 'xr-ref' });
    const link = direct ? null : ui.check('Insert as &hyperlink', true, null);
    const ab = direct ? null : ui.check('Include above/&below', false, null);
    const list = direct ? null : G.listBox([], null, { height: 200, translate: 'no' });
    const root = h('div', { id: 'react-wordx-cross-reference-root' });
    let items = [];
    const refill = (selectedType = type.value) => {
      L.clear(ref);
      for (const [v, l] of what[selectedType] || capWhat) ref.appendChild(h('option', { value: v, text: l }));
      items = F().xrefItems(selectedType);
      list.setItems(items.map((x, i) => ({ value: String(i), label: x.text || L.t('(empty)') })));
      if (items.length) list.set('0');
    };
    if (direct) {
      G.reactCrossRefCatalog = (selectedType = 'Heading') => ({
        types: types.map((value) => ({ value, label: kinds.includes(value) ? L.t(value) : value })),
        references: (what[selectedType] || capWhat).map(([value, label]) => ({ value, label: L.t(label) })),
        items: F().xrefItems(selectedType).map((item, index) => ({ value: String(index), label: item.text || L.t('(empty)') })),
      });
      G.reactCrossRefAction = (selectedType, itemIndex, reference, hyperlink, aboveBelow) => {
        const candidates = F().xrefItems(selectedType);
        const item = candidates[itemIndex];
        if (!item) return;
        F().insertXref(selectedType, item, reference, { link: hyperlink, abovebelow: aboveBelow });
      };
    } else refill();
    const dlg = ui.dialog({
      title: 'Cross-reference', width: 460, modeless: true,
      body: direct ? root : h('div', { class: 'col' }, G.row(G.f('Reference type:', type), G.f('Insert &reference to:', ref)), G.row(link, ab), h('div', { class: 'cd-cap', text: 'For which item:' }), list),
      buttons: [{ label: '&Insert', primary: true, onClick: () => {
        if (direct) { const selectedType = root.querySelector('#react-wordx-cross-reference-type')?.value || 'Heading'; G.reactCrossRefAction(selectedType, +(root.querySelector('#react-wordx-cross-reference-items')?.value || 0), root.querySelector('#react-wordx-cross-reference-reference')?.value || 'text', !!root.querySelector('#react-wordx-cross-reference-link')?.checked, !!root.querySelector('#react-wordx-cross-reference-above-below')?.checked); return false; }
        const it = items[+list.get()]; if (!it) return false; F().insertXref(type.value, it, ref.value, { link: link.input.checked, abovebelow: ab.input.checked }); return false;
      } }, { label: 'Close' }],
    });
    void dlg;
  };

  /* ================= Index and Tables ================= */
  G.indexTables = function (tab) {
    const direct = !!(window.VO && VO.reactDirectHost);
    if (direct) {
      const initialTab = tab === 'tof' ? 'tof' : tab === 'index' ? 'index' : 'toc';
      const root = h('div', { id: 'react-wordx-index-tables-root' });
      const styles = () => Object.keys(doc().styles).filter((id) => doc().styles[id].type === 'paragraph').sort((a,b)=>D.styleDisplayName(doc().styles[a]).localeCompare(D.styleDisplayName(doc().styles[b]))).map((id)=>({ id, name: D.styleDisplayName(doc().styles[id]), level: D.styleProps(doc(),id).pPr.outline }));
      G.reactIndexTablesCatalog = () => ({ initialTab, styles: styles(), labels: ['Figure','Table','Equation'].concat(L.store.get('captionLabels', [])), leaders: [['dot','.......'],['hyphen','-------'],['underscore','_______'],['none',L.N('(none)')]] });
      G.reactIndexTablesAction = (action, value) => {
        if (action === 'modify') { G.style('TOC1'); return; }
        if (action === 'mark') { G.markEntry(); return; }
        if (action !== 'insert') return;
        const { tab: current, toc, figures, index } = value;
        if (current === 'toc') {
          let instr = `TOC \\o "1-${toc.levels}"`;
          if (toc.styles) { const st = Object.entries(toc.styles).filter(([id])=>!id.startsWith('__')&&!/^Heading\\d$/.test(id)).map(([id,level])=>`${D.styleDisplayName(doc().styles[id])},${level}`).join(','); if (!toc.styles.__outline) instr='TOC'; if (st) instr+=` \\t "${st}"`; if (toc.styles.__fields) instr+=' \\f'; }
          if (toc.links) instr+=' \\h'; instr+=' \\z \\u'; if (!toc.show) instr+=' \\n'; else if (!toc.right) instr+=' \\p " "';
          F().insertGenerated(instr); if (toc.format!=='template') applyTocFormat(toc.format,toc.levels);
        } else if (current === 'tof') {
          let instr=`TOC \\h \\z \\c "${figures.label}"`; if (!figures.include) instr+=` \\a "${figures.label}"`; if (!figures.show) instr+=' \\n'; F().insertGenerated(instr);
        } else if (current === 'index') {
          let instr=`INDEX \\c "${index.columns}"`; if (index.type==='runin') instr+=' \\r'; if (index.headings) instr+=' \\h "A"'; if (index.right) instr+=' \\e "\\t"'; F().insertGenerated(instr,{columns:index.columns});
        }
      };
      ui.dialog({ title:'Index and Tables', width:520, body:root, buttons:[{label:'OK',primary:true,onClick:()=>{const body=root.querySelector('#react-wordx-index-tables-body');if(!body)return;const current=body.querySelector('[role=tab][aria-selected=true]')?.dataset.indexTab||initialTab;const read=(id)=>body.querySelector('#'+id);const tocStyles={};body.querySelectorAll('[data-toc-style]').forEach((input)=>{if(input.value.trim())tocStyles[input.dataset.tocStyle]=Number(input.value)});if(body.querySelector('#react-wordx-toc-outline'))tocStyles.__outline=body.querySelector('#react-wordx-toc-outline').checked;if(body.querySelector('#react-wordx-toc-fields'))tocStyles.__fields=body.querySelector('#react-wordx-toc-fields').checked;G.reactIndexTablesAction('insert',{tab:current,toc:{show:read('react-wordx-toc-show')?.checked,right:read('react-wordx-toc-right')?.checked,links:read('react-wordx-toc-links')?.checked,leader:read('react-wordx-toc-leader')?.value,levels:Number(read('react-wordx-toc-levels')?.value)||3,format:read('react-wordx-toc-format')?.value||'template',styles:Object.keys(tocStyles).length?tocStyles:null},figures:{label:read('react-wordx-tof-label')?.value||'Figure',include:read('react-wordx-tof-include')?.checked,show:read('react-wordx-tof-show')?.checked},index:{type:body.querySelector('[name=react-wordx-index-type]:checked')?.value||'indent',columns:Number(read('react-wordx-index-columns')?.value)||2,right:read('react-wordx-index-right')?.checked,headings:read('react-wordx-index-headings')?.checked}})}},{label:'Cancel'}]});
      return;
    }
    const leaders = [['dot', '.......'], ['hyphen', '-------'], ['underscore', '_______'], ['none', L.N('(none)')]];
    /* TOC */
    const tShow = ui.check('&Show page numbers', true, () => pvToc());
    const tRight = ui.check('&Right align page numbers', true, () => pvToc());
    const tLead = ui.select(leaders, 'dot', () => pvToc(), { id: 'it-tl' });
    const tLinks = ui.check('&Use hyperlinks instead of page numbers (Web)', true, null);
    const tLevels = G.num(3, { min: 1, max: 9, id: 'it-lv', onChange: () => pvToc() });
    const tFormat = ui.select([['template', 'From template'], ['classic', 'Classic'], ['distinctive', 'Distinctive'], ['formal', 'Formal'], ['simple', 'Simple']], 'template', null, { id: 'it-tf' });
    let tStyles = null;
    const tPv = h('div', { class: 'preview-box', style: 'height:150px;flex-direction:column;align-items:stretch;justify-content:flex-start;padding:6px;font:11px Times New Roman,Tinos,serif' });
    const pvToc = () => { L.clear(tPv); for (let i = 1; i <= Math.min(3, tLevels.get()); i++) tPv.appendChild(h('div', { style: `display:flex;padding-left:${(i - 1) * 12}px;${i === 1 ? 'font-weight:bold' : ''}` }, h('span', { text: L.t('Heading {0}', i) }), h('span', { style: `flex:1;${tRight.input.checked && tLead.value !== 'none' ? 'border-bottom:1px dotted #000;margin:0 2px 3px' : ''}` }), h('span', { text: tShow.input.checked ? String(i * 2 - 1) : '' }))); };
    pvToc();
    const tocOptions = () => {
      const d = doc();
      const rows = h('div', { style: 'max-height:240px;overflow:auto;border:1px solid #919b9c;background:#fff' });
      const ids = Object.keys(d.styles).filter((k) => d.styles[k].type === 'paragraph').sort((a, b) => D.styleDisplayName(d.styles[a]).localeCompare(D.styleDisplayName(d.styles[b])));
      const inputs = {};
      for (const k of ids) { const lv = D.styleProps(d, k).pPr.outline; const inp = h('input', { type: 'text', style: 'width:36px', value: tStyles ? tStyles[k] || '' : lv != null && lv < tLevels.get() && /^Heading\d$/.test(k) ? String(lv + 1) : '' }); inputs[k] = inp; rows.appendChild(h('div', { class: 'row', style: 'justify-content:space-between;padding:1px 4px' }, G.raw('span', null, D.styleDisplayName(d.styles[k])), inp)); }
      const outl = ui.check('&Outline levels', true, null), fields = ui.check('Table &entry fields', false, null);
      ui.dialog({ title: 'Table of Contents Options', width: 340, body: h('div', { class: 'col' }, h('div', { text: 'Build table of contents from: Styles — TOC level:' }), rows, outl, fields), buttons: [{ label: 'OK', primary: true, onClick: () => { tStyles = {}; for (const k in inputs) if (inputs[k].value.trim()) tStyles[k] = +inputs[k].value; tStyles.__outline = outl.input.checked; tStyles.__fields = fields.input.checked; } }, { label: 'Cancel' }] });
    };
    const tocBody = h('div', { class: 'col' }, G.row(h('div', { class: 'col', style: 'flex:1' }, h('div', { class: 'cd-cap', text: 'Print Preview' }), tPv, tShow, tRight, G.f('Ta&b leader:', tLead)), h('div', { class: 'col', style: 'flex:1' }, h('div', { class: 'cd-cap', text: 'Web Preview' }), h('div', { class: 'preview-box', style: 'height:150px;flex-direction:column;align-items:flex-start;justify-content:flex-start;padding:6px;gap:3px' }, h('u', { text: 'Heading 1', style: 'color:#00f' }), h('u', { text: 'Heading 2', style: 'color:#00f;margin-left:10px' }), h('u', { text: 'Heading 3', style: 'color:#00f;margin-left:20px' })), tLinks)), ui.group('General', G.row(G.f('Forma&ts:', tFormat), G.f('Show &levels:', tLevels))), h('div', { class: 'tp-btns' }, ui.button('&Options...', tocOptions, { class: 'btn small' }), ui.button('&Modify...', () => G.style('TOC1'), { class: 'btn small' })));
    /* Table of figures */
    const fLabel = ui.select(['Figure', 'Table', 'Equation'].concat(L.store.get('captionLabels', [])).map((x) => [x, x]), 'Figure', null, { id: 'it-fl', translate: 'no' });
    const fIncl = ui.check('Include label and &number', true, null);
    const fShow = ui.check('&Show page numbers', true, null);
    const fRight = ui.check('&Right align page numbers', true, null);
    const fLead = ui.select(leaders, 'dot', null, { id: 'it-fle' });
    const tofBody = h('div', { class: 'col' }, G.f('Caption &label:', fLabel), fIncl, fShow, fRight, G.f('Ta&b leader:', fLead), h('div', { class: 'tp-note', text: 'Captions inserted with Insert ▸ Reference ▸ Caption are collected here.' }));
    /* Index */
    const iType = G.radios('it-it', [['indent', 'Inde&nted'], ['runin', 'Ru&n-in']], 'indent');
    const iCols = G.num(2, { min: 1, max: 4, id: 'it-ic' });
    const iRight = ui.check('&Right align page numbers', false, null);
    const iLead = ui.select(leaders, 'dot', null, { id: 'it-il' });
    const iHead = ui.check('Show letter &headings', false, null);
    const idxBody = h('div', { class: 'col' }, G.row(ui.group('Type', iType), h('div', { class: 'col' }, G.f('C&olumns:', iCols), iRight, G.f('Ta&b leader:', iLead), iHead)), h('div', { class: 'tp-btns' }, ui.button('Mar&k Entry...', () => G.markEntry(), { class: 'btn small' })), h('div', { class: 'tp-note', text: 'Mark entries first (Alt+Shift+X), then insert the index where it should appear.' }));
    const toaBody = h('div', { class: 'tp-note', text: 'Tables of authorities need citations marked with TA fields. Use Insert ▸ Field ▸ TA to mark citations, then Field ▸ TOA to build the table.' });
    let cur = tab === 'toc' || tab == null ? 1 : tab === 'tof' ? 2 : tab === 'index' ? 0 : 1;
    const tabs = ui.tabs([{ label: 'Inde&x', body: idxBody, onShow: () => { cur = 0; } }, { label: 'Table of &Contents', body: tocBody, onShow: () => { cur = 1; } }, { label: 'Table of &Figures', body: tofBody, onShow: () => { cur = 2; } }, { label: 'Table of &Authorities', body: toaBody, onShow: () => { cur = 3; } }], cur);
    ui.dialog({
      title: 'Index and Tables', width: 520, body: tabs,
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        if (cur === 1) {
          let instr = `TOC \\o "1-${tLevels.get()}"`;
          if (tStyles) { const st = Object.entries(tStyles).filter(([k]) => !k.startsWith('__') && !/^Heading\d$/.test(k)).map(([k, v]) => `${D.styleDisplayName(doc().styles[k])},${v}`).join(','); if (!tStyles.__outline) instr = 'TOC'; if (st) instr += ` \\t "${st}"`; if (tStyles.__fields) instr += ' \\f'; }
          if (tLinks.input.checked) instr += ' \\h';
          instr += ' \\z \\u';
          if (!tShow.input.checked) instr += ' \\n';
          else if (!tRight.input.checked) instr += ' \\p " "';
          F().insertGenerated(instr);
          if (tFormat.value !== 'template') applyTocFormat(tFormat.value, tLevels.get());
        } else if (cur === 2) {
          let instr = `TOC \\h \\z \\c "${fLabel.value}"`;
          if (!fIncl.input.checked) instr += ` \\a "${fLabel.value}"`;
          if (!fShow.input.checked) instr += ' \\n';
          F().insertGenerated(instr);
        } else if (cur === 0) {
          let instr = `INDEX \\c "${iCols.get()}"`;
          if (iType.get() === 'runin') instr += ' \\r';
          if (iHead.input.checked) instr += ' \\h "A"';
          if (iRight.input.checked) instr += ' \\e "\t"';
          F().insertGenerated(instr, { columns: iCols.get() });
        }
      } }, { label: 'Cancel' }],
    });
  };
  function applyTocFormat(kind, levels) {
    const d = doc();
    const look = { classic: [{ b: true, caps: true }, { smallCaps: true }, { i: true }], distinctive: [{ b: true, font: 'Arial', sz: 12 }, { font: 'Arial', sz: 11 }, { sz: 10 }], formal: [{ caps: true, b: true }, { caps: true }, { i: true }], simple: [{}, {}, {}] }[kind];
    E.edit('TOC Format', () => { D.touchKey(d, 'styles'); for (let i = 1; i <= levels; i++) { const id = D.useStyle(d, 'TOC' + i); d.styles[id].rPr = Object.assign({}, look[Math.min(i - 1, 2)]); } D.stylesChanged(); return E.sel; });
  }
  G.markEntry = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const sel = E.sel && !E.collapsed() ? O.textRange(...E.range(), ' ').trim() : '';
    const main = direct ? null : h('input', { type: 'text', id: 'me-main', value: sel, style: 'width:100%' });
    const sub = direct ? null : h('input', { type: 'text', id: 'me-sub', style: 'width:100%' });
    const opt = direct ? null : G.radios('me-o', [['cur', '&Current page'], ['xref', 'Cross-&reference:']], 'cur');
    const xref = direct ? null : h('input', { type: 'text', id: 'me-x', value: 'See ', style: 'width:100%' });
    const bold = direct ? null : ui.check('&Bold', false, null), ital = direct ? null : ui.check('&Italic', false, null);
    const root = h('div', { id: 'react-wordx-mark-entry-root' });
    const mark = (all, values) => {
      const m = (direct ? values.main : main.value).trim();
      if (!m) return;
      const subentry = direct ? values.sub : sub.value;
      const entry = (m + (subentry.trim() ? ':' + subentry.trim() : '')).replace(/"/g, '\\"');
      let instr = 'XE "' + entry + '"';
      if ((direct ? values.option : opt.get()) === 'xref') instr += ' \\t "' + (direct ? values.xref : xref.value).replace(/"/g, '\\"') + '"';
      if (direct ? values.bold : bold.input.checked) instr += ' \\b';
      if (direct ? values.italic : ital.input.checked) instr += ' \\i';
      if (!all) { E.edit('Mark Index Entry', () => { const [, b] = E.range(); return O.insertField(b, instr, '', {}); }); return; }
      /* Mark All: every occurrence of the text in the main story */
      const d = doc();
      D.reindex(d);
      let n = 0;
      E.edit('Mark Index Entries', () => {
        for (const p of d.main.paras.slice()) {
          const t = D.ptext(p);
          let i = t.toLowerCase().indexOf(m.toLowerCase());
          const hits = [];
          while (i >= 0) { hits.push(i + m.length); i = t.toLowerCase().indexOf(m.toLowerCase(), i + m.length); }
          for (const at of hits.reverse()) { O.insertField(D.pos(p, at), instr, '', {}); n++; }
        }
        return E.sel;
      });
      A().status(n === 1 ? L.t('{0} entry marked.', n) : L.t('{0} entries marked.', n));
    };
    G.reactMarkEntryState = () => ({ main: sel, sub: '', option: 'cur', xref: 'See ', bold: false, italic: false });
    G.reactMarkEntryAction = (all, values) => mark(all, values);
    const values = () => ({ main: root.querySelector('#react-wordx-mark-main')?.value || '', sub: root.querySelector('#react-wordx-mark-sub')?.value || '', option: root.querySelector('[name=react-wordx-mark-option]:checked')?.value || 'cur', xref: root.querySelector('#react-wordx-mark-xref')?.value || '', bold: !!root.querySelector('#react-wordx-mark-bold')?.checked, italic: !!root.querySelector('#react-wordx-mark-italic')?.checked });
    ui.dialog({ title: 'Mark Index Entry', width: 360, modeless: true, body: direct ? root : h('div', { class: 'col' }, G.f('Main &entry:', main), G.f('&Subentry:', sub), ui.group('Options', opt, xref), ui.group('Page number format', bold, ital)), buttons: [{ label: '&Mark', primary: true, onClick: () => { if (direct) G.reactMarkEntryAction(false, values()); else mark(false); return false; } }, { label: 'Mark &All', onClick: () => { if (direct) G.reactMarkEntryAction(true, values()); else mark(true); return false; } }, { label: 'Close' }] });
  };

  /* ================= Bookmark ================= */
  G.bookmark = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    let reactName = '', reactSort = 'name', reactHidden = false, reactNames = [];
    const sortBy = direct ? { get: () => reactSort } : G.radios('bm-s', [['name', '&Name'], ['loc', '&Location']], 'name', () => fill());
    const hidden = direct ? null : ui.check('&Hidden bookmarks', false, () => fill());
    const name = h('input', { type: 'text', id: 'bm-name', style: 'width:100%' });
    const list = direct ? null : G.listBox([], null, { height: 160, translate: 'no', onChange: (v) => { name.value = v; }, onDbl: () => go() });
    const getName = () => direct ? reactName : name.value;
    const fill = () => {
      const all = Array.from(F().bookmarks().values()).filter((b) => (direct ? reactHidden : hidden.input.checked) || !b.name.startsWith('_'));
      if (sortBy.get() === 'name') all.sort((a, b) => a.name.localeCompare(b.name)); else all.sort((a, b) => (a.p._o || 0) - (b.p._o || 0) || a.o - b.o);
      if (direct) { reactNames = all.map((b) => b.name); L.bus.emit('wordx-bookmark-refresh'); }
      else list.setItems(all.map((b) => b.name));
    };
    fill();
    if (E.sel && !E.collapsed()) { const t = O.textRange(...E.range(), ' ').trim().replace(/[^\p{L}\p{N}_]/gu, '_').replace(/^[^\p{L}]+/u, '').slice(0, 40); name.value = t; reactName = t; }
    const valid = (n) => /^[\p{L}][\p{L}\p{N}_]{0,39}$/u.test(n);
    const go = () => { const n = getName().trim(); if (n) A().gotoBookmark(n); };
    const reactRoot = h('div', { id: 'react-wordx-bookmark-root' });
    G.reactBookmarkCatalog = () => ({ name: reactName, sort: reactSort, hidden: reactHidden, names: reactNames.slice() });
    G.reactBookmarkAction = (action, value) => {
      if (action === 'name') { reactName = value; name.value = value; }
      else if (action === 'select') { reactName = value; name.value = value; }
      else if (action === 'sort') { reactSort = value; fill(); }
      else if (action === 'hidden') { reactHidden = value === 'true'; fill(); }
      else if (action === 'add') { const n = getName().trim(); if (!valid(n)) { ui.msg('Bookmark names must begin with a letter and contain only letters, numbers and underscores.', { icon: 'warn' }); return false; } const [a, b] = E.range(); E.edit(L.N('Bookmark'), () => { F().addBookmark(n, a, b); return E.sel; }); fill(); }
      else if (action === 'delete') { const bm = F().bookmarks().get(getName().trim()); if (bm) E.edit('Delete Bookmark', () => { F().removeBookmarkItems(bm); return E.sel; }); reactName = ''; name.value = ''; fill(); }
      else if (action === 'goto') go();
    };
    ui.dialog({
      title: 'Bookmark', width: 320, focus: direct ? undefined : name,
      body: direct ? reactRoot : h('div', { class: 'col' }, h('label', { for: 'bm-name', html: ui.ampHTML('&Bookmark name:') }), name, list, ui.group('Sort by:', sortBy), hidden),
      buttons: [{ label: '&Add', primary: true, onClick: () => { const n = getName().trim(); if (!valid(n)) { ui.msg('Bookmark names must begin with a letter and contain only letters, numbers and underscores.', { icon: 'warn' }); return false; } const [a, b] = E.range(); E.edit(L.N('Bookmark'), () => { F().addBookmark(n, a, b); return E.sel; }); fill(); if (direct) return false; } }, { label: '&Delete', onClick: () => { G.reactBookmarkAction('delete'); fill(); return false; } }, { label: '&Go To', onClick: () => { go(); return false; } }, { label: 'Cancel' }],
    });
  };

  /* ================= Hyperlink ================= */
  G.hyperlink = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const d = doc();
    let cur = null;
    if (E.sel) {
      const it = D.itemAfter(E.sel.f.p, E.sel.f.o) || D.itemBefore(E.sel.f.p, E.sel.f.o);
      if (it && it.rPr && it.rPr.link) cur = it.rPr.link;
    }
    /* the whole link text when the caret sits in a link */
    let range = E.sel ? E.range() : null;
    if (cur && E.collapsed()) {
      const p = E.sel.f.p; let pos = 0, s = -1, e = -1;
      for (const it of p.runs) { const n = D.ilen(it); if (it.rPr && it.rPr.link && JSON.stringify(it.rPr.link) === JSON.stringify(cur) && pos <= E.sel.f.o + 1) { if (s < 0) s = pos; e = pos + n; } else if (s >= 0 && pos >= E.sel.f.o) break; pos += n; }
      if (s >= 0) range = [D.pos(p, s), D.pos(p, e)];
    }
    const selText = range && !D.eqPos(range[0], range[1]) ? O.textRange(range[0], range[1], ' ') : '';
    const disp = h('input', { type: 'text', id: 'hl-disp', value: selText || '<<Selection in Document>>', style: 'width:100%' });
    if (!selText) disp.value = '';
    const addr = h('input', { type: 'text', id: 'hl-addr', value: cur && cur.url && !/^mailto:/.test(cur.url) ? cur.url : '', style: 'width:100%' });
    const tip = cur && cur.tip ? cur.tip : '';
    const place = G.listBox([], cur && cur.anchor ? cur.anchor : null, { height: 170, translate: 'no' });   // headings and bookmarks: the document's own
    const places = [{ value: '', label: L.t('Top of the Document') }];
    D.reindex(d);
    places.push({ value: '', label: L.t('Headings'), group: true });
    for (const p of d.main.paras) { const lv = D.headingLevel(d, p); if (lv) places.push({ value: '§' + p.id, label: '    '.repeat(lv - 1) + D.plainText(p).trim().slice(0, 60) }); }
    places.push({ value: '', label: L.t('Bookmarks'), group: true });
    for (const [n] of F().bookmarks()) if (!n.startsWith('_')) places.push({ value: n, label: n });
    place.setItems(places);
    const mail = h('input', { type: 'text', id: 'hl-mail', value: cur && /^mailto:/.test(cur.url || '') ? cur.url.slice(7).split('?')[0] : '', style: 'width:100%' });
    const subj = h('input', { type: 'text', id: 'hl-subj', style: 'width:100%' });
    let kind = cur && cur.anchor && !cur.url ? 1 : cur && /^mailto:/.test(cur.url || '') ? 2 : 0;
    const tabs = ui.tabs([
      { label: 'E&xisting File or Web Page', body: h('div', { class: 'col' }, h('label', { for: 'hl-addr', html: ui.ampHTML('Addr&ess:') }), addr, h('div', { class: 'tp-note', text: L.t('Type a web address (https://…) or the name of a file.') })), onShow: () => { kind = 0; } },
      { label: 'Place in This &Document', body: h('div', { class: 'col' }, h('div', { text: 'Select a place in this document:' }), place), onShow: () => { kind = 1; } },
      { label: 'E-&mail Address', body: h('div', { class: 'col' }, h('label', { for: 'hl-mail', text: 'E-mail address:' }), mail, h('label', { for: 'hl-subj', text: 'Subject:' }), subj), onShow: () => { kind = 2; } },
    ], kind);
    let screenTip = tip;
    let reactHyperlink = { kind, display: disp.value, address: addr.value, place: cur && cur.anchor ? cur.anchor : '', email: mail.value, subject: subj.value, screenTip, places };
    G.reactHyperlinkCatalog = () => ({ ...reactHyperlink, places: reactHyperlink.places.slice() });
    G.reactHyperlinkAction = async (action, value) => {
      if (action === 'change') { reactHyperlink = { ...reactHyperlink, ...(value || {}) }; kind = reactHyperlink.kind; }
      else if (action === 'screenTip') { const result = await ui.prompt('ScreenTip text:', reactHyperlink.screenTip, 'Set Hyperlink ScreenTip'); if (result != null) { screenTip = result; reactHyperlink = { ...reactHyperlink, screenTip: result }; L.bus.emit('wordx-hyperlink-refresh'); } }
      else if (action === 'remove') A().removeHyperlink();
    };
    const reactRoot = h('div', { id: 'react-wordx-hyperlink-root' });
    ui.dialog({
      title: cur ? 'Edit Hyperlink' : 'Insert Hyperlink', width: 520, focus: kind === 0 ? addr : null,
      body: direct ? reactRoot : h('div', { class: 'col' }, G.row(h('label', { for: 'hl-disp', html: ui.ampHTML('&Text to display:'), style: 'width:110px' }), h('div', { style: 'flex:1' }, disp), ui.button('ScreenTi&p...', async () => { const t = await ui.prompt('ScreenTip text:', screenTip, 'Set Hyperlink ScreenTip'); if (t != null) screenTip = t; }, { class: 'btn small' })), tabs),
      buttons: (cur ? [{ label: '&Remove Link', onClick: () => { A().removeHyperlink(); } }] : []).concat([{ label: 'OK', primary: true, onClick: () => {
        let link;
        const form = direct ? reactHyperlink : { kind, address: addr.value, place: place.get(), email: mail.value, subject: subj.value, display: disp.value, screenTip };
        if (form.kind === 0) { let u = form.address.trim(); if (!u) return false; if (/^www\./i.test(u)) u = 'http://' + u; link = { url: u }; }
        else if (form.kind === 1) {
          let a = form.place;
          if (a && a.startsWith('§')) { const p = D.byId(d, +a.slice(1)); E.edit(L.N('Bookmark'), () => { a = F().ensureParaBookmark(p, '_Toc'); return E.sel; }); }
          link = { anchor: a || '_top' };
        } else { const m = form.email.trim(); if (!m) return false; link = { url: 'mailto:' + m + (form.subject ? '?subject=' + encodeURIComponent(form.subject) : '') }; }
        if (form.screenTip) link.tip = form.screenTip;
        if (cur?.target) link.target = cur.target; // the frame it opens in, which this dialog doesn't show
        const text = form.display || (form.kind === 0 ? form.address : form.kind === 2 ? form.email : form.place || 'Top of the Document');
        E.edit(cur ? 'Edit Hyperlink' : 'Insert Hyperlink', () => {
          const style = D.useStyle(d, 'Hyperlink');
          if (range && !D.eqPos(range[0], range[1]) && text === selText) { O.formatRange(range[0], range[1], (r) => Object.assign(r, { link, style })); return { a: range[0], f: range[1] }; }
          E.sel = { a: range[0], f: range[1] };
          const pos = E.deleteSelection();
          const base = O.inheritRPr(pos);
          delete base.link;
          return O.insertText(pos, text, Object.assign(base, { link, style }));
        });
        E.pending = { link: undefined, style: undefined };
      } }, { label: 'Cancel' }]),
    });
  };

  /* ================= Insert Object ================= */
  G.insertObject = function () {
    const types = [['chart', L.N('Chart')], ['equation', L.N('Equation')], ['wordart', L.N('LettersArt')], ['excel', L.N('Spreadsheet (table)')], ['document', L.N('Document (inserted as text)')], ['picture', L.N('Bitmap Image')]];
    const direct = !!(window.VO && VO.reactDirectHost);
    let selected = 'chart', tab = 0, dlg = null;
    const lb = direct ? { get: () => selected } : G.listBox(types.map(([v, l]) => ({ value: v, label: l })), 'chart', { height: 160, onDbl: () => ok.click() });
    const root = h('div', { id: 'react-wordx-object-root' });
    G.reactObjectCatalog = () => types.map(([value, label]) => ({ value, label: L.t(label) }));
    G.reactObjectAction = (action, value) => {
      if (action === 'select') selected = value;
      else if (action === 'tab') tab = value === 'file' ? 1 : 0;
      else if (action === 'browse') { dlg.close(null); A().insertFileDialog(); }
      else if (action === 'insert') insert(selected);
    };
    const insert = (v) => {
      if (v === 'chart') A().insertChart();
      else if (v === 'equation') G.equation();
      else if (v === 'wordart') G.wordart();
      else if (v === 'excel') L.tables.insert(4, 5, { worksheet: true });
      else if (v === 'document') A().insertFileDialog();
      else A().insertPictureDialog();
    };
    dlg = ui.dialog({ title: 'Object', width: 360, body: direct ? root : ui.tabs([{ label: '&Create New', body: h('div', { class: 'col' }, h('label', { text: 'Object type:' }), lb) }, { label: 'Create from &File', body: h('div', { class: 'col' }, h('div', { class: 'tp-note', text: 'Insert the contents of a .docx, .txt or .htm file, or a picture.' }), ui.button('&Browse...', () => { dlg.close(null); A().insertFileDialog(); })) }]), buttons: [{ label: 'OK', primary: true, onClick: () => {
      const v = lb.get();
      insert(v);
    } }, { label: 'Cancel' }] });
    const ok = dlg.buttons[0];
  };
  /** a simple equation editor: linear format with Unicode math symbols */
  G.equation = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const inp = direct ? null : h('textarea', { rows: 3, id: 'eq-in', style: 'width:100%;font:16px "Cambria Math",Caladea,serif' });
    const pal = h('div', { class: 'sym-recent', style: 'grid-template-columns:repeat(16,26px)' });
    if (!direct) for (const c of '±×÷≠≈≤≥∞√∛∑∏∫∂∆∇πθαβγδλμσφωΩ→⇒∈∉⊂⊃∪∩∀∃¬∧∨²³⁴ⁿ₀₁₂ᵢ½'.split('')) { const b = h('button', { type: 'button', text: c }); b.addEventListener('click', () => { const s = inp.selectionStart; inp.value = inp.value.slice(0, s) + c + inp.value.slice(inp.selectionEnd); inp.focus(); inp.setSelectionRange(s + 1, s + 1); }); pal.appendChild(b); }
    const root = h('div', { id: 'react-wordx-equation-root' });
    G.reactEquationAction = (action, value) => {
      if (action !== 'insert' || !String(value || '').trim()) return;
      E.edit('Insert Equation', () => { const pos = E.deleteSelection(); return O.insertText(pos, String(value).trim(), Object.assign(O.inheritRPr(pos), { font: 'Cambria Math', i: true })); });
    };
    ui.dialog({ title: 'Equation', width: 470, focus: inp || root, body: direct ? root : h('div', { class: 'col' }, pal, inp, h('div', { class: 'tp-note', text: 'Type the equation; it is inserted in Cambria Math italics.' })), buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (direct) G.reactEquationAction('insert', root.querySelector('textarea')?.value || '');
      else if (inp.value.trim()) E.edit('Insert Equation', () => { const pos = E.deleteSelection(); return O.insertText(pos, inp.value.trim(), Object.assign(O.inheritRPr(pos), { font: 'Cambria Math', i: true })); });
    } }, { label: 'Cancel' }] });
  };

  /* ================= Diagram Gallery ================= */
  G.diagram = function () {
    const kinds = [['org', 'Organization Chart'], ['cycle', 'Cycle Diagram'], ['radial', 'Radial Diagram'], ['pyramid', 'Pyramid Diagram'], ['venn', 'Venn Diagram'], ['target', 'Target Diagram']];
    const direct = !!(window.VO && VO.reactDirectHost);
    if (direct) {
      const root = h('div', { id: 'react-wordx-diagram-root' });
      G.reactDiagramCatalog = () => kinds.map(([value, label]) => ({ value, label: L.t(label) }));
      G.reactDiagramAction = (value) => insertDiagram(value);
      ui.dialog({ title: 'Diagram Gallery', width: 340, body: root, buttons: [{ label: 'OK', primary: true, onClick: () => G.reactDiagramAction(root.querySelector('[data-wordx-diagram].on')?.getAttribute('data-wordx-diagram') || 'org') }, { label: 'Cancel' }] });
      return;
    }
    const svg = { org: '<svg width="56" height="44"><rect x="20" y="2" width="16" height="10" fill="#fdd" stroke="#333"/><path d="M28 12v6M10 18h36M10 18v4M28 18v4M46 18v4" stroke="#333" fill="none"/><rect x="2" y="22" width="16" height="10" fill="#dfd" stroke="#333"/><rect x="20" y="22" width="16" height="10" fill="#dfd" stroke="#333"/><rect x="38" y="22" width="16" height="10" fill="#dfd" stroke="#333"/></svg>', cycle: '<svg width="56" height="44"><circle cx="28" cy="22" r="16" fill="none" stroke="#888" stroke-dasharray="3 2"/><circle cx="28" cy="6" r="5" fill="#9cf" stroke="#333"/><circle cx="42" cy="30" r="5" fill="#fc9" stroke="#333"/><circle cx="14" cy="30" r="5" fill="#cf9" stroke="#333"/></svg>', radial: '<svg width="56" height="44"><path d="M28 22L28 6M28 22L44 32M28 22L12 32" stroke="#333"/><circle cx="28" cy="22" r="7" fill="#fc9" stroke="#333"/><circle cx="28" cy="6" r="5" fill="#9cf" stroke="#333"/><circle cx="44" cy="32" r="5" fill="#9cf" stroke="#333"/><circle cx="12" cy="32" r="5" fill="#9cf" stroke="#333"/></svg>', pyramid: '<svg width="56" height="44"><path d="M28 2L50 42H6z" fill="#fc9" stroke="#333"/><path d="M19 18h18M12 30h32" stroke="#333"/></svg>', venn: '<svg width="56" height="44"><circle cx="22" cy="18" r="12" fill="#9cf" fill-opacity=".6" stroke="#333"/><circle cx="34" cy="18" r="12" fill="#fc9" fill-opacity=".6" stroke="#333"/><circle cx="28" cy="28" r="12" fill="#cf9" fill-opacity=".6" stroke="#333"/></svg>', target: '<svg width="56" height="44"><circle cx="24" cy="24" r="18" fill="#9cf" stroke="#333"/><circle cx="24" cy="24" r="12" fill="#fc9" stroke="#333"/><circle cx="24" cy="24" r="6" fill="#f99" stroke="#333"/></svg>' };
    let pick = 'org';
    const grid = h('div', { class: 'preset-grid' });
    for (const [k, l] of kinds) { const b = h('button', { type: 'button', class: k === pick ? 'on' : '', 'data-tip': l, 'aria-label': l, html: svg[k] }); b.addEventListener('click', () => { pick = k; L.$$('button', grid).forEach((x) => x.classList.toggle('on', x === b)); name.textContent = L.t(l); }); b.addEventListener('dblclick', () => ok.click()); grid.appendChild(b); }
    const name = h('div', { class: 'cd-cap', text: kinds[0][1] });
    const dlg = ui.dialog({ title: 'Diagram Gallery', width: 340, body: h('div', { class: 'col' }, h('div', { text: 'Select a diagram type:' }), grid, name), buttons: [{ label: 'OK', primary: true, onClick: () => insertDiagram(pick) }, { label: 'Cancel' }] });
    const ok = dlg.buttons[0];
  };
  function insertDiagram(kind) {
    const S = L.drawing.newShape;
    const tb = (sh, text) => { sh.tb = { kind: 'tb', blocks: [D.para(text ? [D.text(text, { sz: 10 })] : [], { jc: 'center' })] }; sh.ins = { l: 3.6, t: 3.6, r: 3.6, b: 3.6 }; sh.anchor = 'ctr'; return sh; };
    const kid = (sh, x, y) => Object.assign(sh, { x, y });
    const W = 324, H = 216;
    const kids = [];
    const fills = ['#99CCFF', '#FFCC99', '#CCFF99', '#FF9999', '#CC99FF', '#FFFF99'];
    if (kind === 'org') {
      kids.push(kid(tb(S('rect', 90, 40), 'Manager'), 117, 8));
      for (let i = 0; i < 3; i++) kids.push(kid(tb(S('rect', 90, 40), 'Employee ' + (i + 1)), 9 + i * 108, 120));
      kids.push(kid(S('line', 0, 40), 162, 48), kid(S('line', 216, 0), 54, 88));
      for (let i = 0; i < 3; i++) kids.push(kid(S('line', 0, 32), 54 + i * 108, 88));
    } else if (kind === 'pyramid') {
      for (let i = 0; i < 3; i++) { const sh = S('trapezoid', 120 + i * 80, 56); sh.fill = { t: 'solid', c: fills[i], a: 1 }; tb(sh, 'Level ' + (i + 1)); kids.push(kid(sh, (W - (120 + i * 80)) / 2, 16 + i * 62)); }
      kids[0].geom = 'triangle'; kids[0].w = 120;
    } else {
      const n = kind === 'target' ? 3 : kind === 'venn' ? 3 : 4;
      for (let i = 0; i < n; i++) {
        const a = (i / n) * Math.PI * 2 - Math.PI / 2;
        let sh;
        if (kind === 'target') { const r = 100 - i * 30; sh = S('ellipse', r * 2, r * 2); kids.push(kid(Object.assign(sh, { fill: { t: 'solid', c: fills[i], a: 1 } }), W / 2 - r, H / 2 - r)); continue; }
        if (kind === 'venn') { sh = S('ellipse', 120, 120); sh.fill = { t: 'solid', c: fills[i], a: 0.5 }; kids.push(kid(sh, W / 2 - 60 + Math.cos(a) * 40, H / 2 - 60 + Math.sin(a) * 40)); continue; }
        sh = tb(S('ellipse', 72, 52), kind === 'cycle' ? 'Step ' + (i + 1) : 'Item ' + (i + 1));
        sh.fill = { t: 'solid', c: fills[i], a: 1 };
        kids.push(kid(sh, W / 2 - 36 + Math.cos(a) * 110, H / 2 - 26 + Math.sin(a) * 75));
      }
      if (kind === 'radial') { const c = tb(S('ellipse', 80, 56), 'Center'); c.fill = { t: 'solid', c: '#FFCC99', a: 1 }; kids.unshift(kid(c, W / 2 - 40, H / 2 - 28)); }
    }
    const g = { t: 'group', w: W, h: H, kids, name: 'Diagram ' + D.nid() };
    E.edit('Insert Diagram', () => O.insertItem(E.deleteSelection(), g));
  }

  /* ================= WordArt ================= */
  const WA_STYLES = [
    { fill: { t: 'solid', c: '#000000' }, warp: 'textPlain' }, { fill: { t: 'grad', ang: 90, stops: [{ p: 0, c: '#FFFFFF' }, { p: 1, c: '#999999' }] }, line: { c: '#000000', w: 0.75 }, warp: 'textPlain' }, { fill: { t: 'solid', c: '#C0C0C0' }, line: { c: '#000000', w: 0.75 }, warp: 'textPlain', i: true }, { fill: { t: 'solid', c: '#000000' }, warp: 'textSlantUp' }, { fill: { t: 'grad', ang: 0, stops: [{ p: 0, c: '#FF9900' }, { p: 1, c: '#FFFF00' }] }, warp: 'textArchUp' },
    { fill: { t: 'solid', c: '#3366FF' }, line: { c: '#000080', w: 0.75 }, warp: 'textPlain', shadow: { dx: 3, dy: 3 } }, { fill: { t: 'grad', ang: 90, stops: [{ p: 0, c: '#FF0000' }, { p: 0.5, c: '#FFFF00' }, { p: 1, c: '#00CC00' }] }, warp: 'textWave1' }, { fill: { t: 'solid', c: '#FFFFFF' }, line: { c: '#3366CC', w: 1.5 }, warp: 'textPlain', b: true }, { fill: { t: 'solid', c: '#993300' }, warp: 'textChevron' }, { fill: { t: 'grad', ang: 90, stops: [{ p: 0, c: '#C0C0C0' }, { p: 1, c: '#333333' }] }, warp: 'textCanUp' },
    { fill: { t: 'solid', c: '#CC0000' }, line: { c: '#660000', w: 0.75 }, warp: 'textArchDown' }, { fill: { t: 'solid', c: '#336633' }, warp: 'textTriangle' }, { fill: { t: 'grad', ang: 45, stops: [{ p: 0, c: '#0000FF' }, { p: 1, c: '#00FFFF' }] }, warp: 'textPlain', b: true }, { fill: { t: 'solid', c: '#FFCC00' }, line: { c: '#996600', w: 0.75 }, warp: 'textSlantDown', shadow: { dx: 2, dy: 2 } }, { fill: { t: 'solid', c: '#660099' }, warp: 'textCircle' },
  ];
  G.wordart = function (mode) {
    const direct = !!(window.VO && VO.reactDirectHost);
    const it = mode && L.drawing.selectedItem();
    if (mode === true && it && it.wordart) { editText(it, (wa) => L.drawing.modifySelected('Edit LettersArt', (x) => { x.wordart = Object.assign({}, x.wordart, wa); }), direct); return; }
    let pick = 0;
    const grid = h('div', { class: 'wa-grid' });
    WA_STYLES.forEach((s, i) => {
      const b = h('button', { type: 'button', class: i === 0 ? 'on' : '', 'aria-label': L.t('LettersArt style {0}', i + 1) });
      const sh = { t: 'shape', w: 66, h: 40, fill: Object.assign({ a: 1 }, s.fill), line: s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }, warp: s.warp, shadow: s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined, wordart: { text: 'LettersArt', font: 'Arial Black', b: s.b, i: s.i } };
      b.appendChild(L.drawing.wordartSVG(sh));
      b.addEventListener('click', () => { pick = i; L.$$('button', grid).forEach((x) => x.classList.toggle('on', x === b)); });
      b.addEventListener('dblclick', () => ok.click());
      grid.appendChild(b);
    });
    const reactRoot = h('div', { id: 'react-wordx-wordart-root' });
    const wordartPreview = (s) => {
      const shape = { t: 'shape', w: 66, h: 40, fill: Object.assign({ a: 1 }, s.fill), line: s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }, warp: s.warp, shadow: s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined, wordart: { text: 'LettersArt', font: 'Arial Black', b: s.b, i: s.i } };
      return L.drawing.wordartSVG(shape).outerHTML;
    };
    G.reactWordartCatalog = () => ({ selected: pick, previews: WA_STYLES.map(wordartPreview) });
    G.reactWordartAction = (action, value) => {
      if (action === 'select') { pick = Number(value); return G.reactWordartCatalog(); }
      if (action !== 'apply') return G.reactWordartCatalog();
      const s = WA_STYLES[pick];
      if (mode === 'gallery' && it) { L.drawing.modifySelected('LettersArt Gallery', (x) => { x.fill = Object.assign({ a: 1 }, s.fill); x.line = s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }; x.warp = s.warp; x.shadow = s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined; }); return; }
      const init = E.sel && !E.collapsed() ? O.textRange(...E.range(), ' ').trim().slice(0, 200) : 'Your Text Here';
      setTimeout(() => editText({ wordart: { text: init, font: 'Arial Black', sz: 36, b: s.b, i: s.i } }, (wa) => {
        const lines = wa.text.split('\n');
        const longest = Math.max(...lines.map((l) => l.length), 1);
        const w = Math.min(468, longest * wa.sz * 0.62 + 10), hh = lines.length * wa.sz * 1.15 + 6;
        const shape = { t: 'shape', geom: 'rect', w: L.round(w, 2), h: L.round(hh, 2), fill: Object.assign({ a: 1 }, s.fill), line: s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }, warp: s.warp, shadow: s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined, wordart: wa, name: 'LettersArt ' + D.nid() };
        E.edit('Insert LettersArt', () => O.insertItem(E.deleteSelection(), shape));
      }, direct), 0);
    };
    const dlg = ui.dialog({ title: 'LettersArt Gallery', width: 470, body: direct ? reactRoot : h('div', { class: 'col' }, h('div', { text: 'Select a LettersArt style:' }), grid), buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (direct) { G.reactWordartAction('apply'); return; }
      const s = WA_STYLES[pick];
      if (mode === 'gallery' && it) { L.drawing.modifySelected('LettersArt Gallery', (x) => { x.fill = Object.assign({ a: 1 }, s.fill); x.line = s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }; x.warp = s.warp; x.shadow = s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined; }); return; }
      const init = E.sel && !E.collapsed() ? O.textRange(...E.range(), ' ').trim().slice(0, 200) : 'Your Text Here';
      setTimeout(() => editText({ wordart: { text: init, font: 'Arial Black', sz: 36, b: s.b, i: s.i } }, (wa) => {
        const lines = wa.text.split('\n');
        const longest = Math.max(...lines.map((l) => l.length), 1);
        const w = Math.min(468, longest * wa.sz * 0.62 + 10), hh = lines.length * wa.sz * 1.15 + 6;
        const shape = { t: 'shape', geom: 'rect', w: L.round(w, 2), h: L.round(hh, 2), fill: Object.assign({ a: 1 }, s.fill), line: s.line ? Object.assign({ dash: 'solid' }, s.line) : { t: 'none' }, warp: s.warp, shadow: s.shadow ? Object.assign({ c: '#808080', a: 0.6 }, s.shadow) : undefined, wordart: wa, name: 'LettersArt ' + D.nid() };
        E.edit('Insert LettersArt', () => O.insertItem(E.deleteSelection(), shape));
      }, direct), 0);
    } }, { label: 'Cancel' }] });
    const ok = dlg.buttons[0];
  };
  function editText(it, done, reactHost = !!(window.VO && VO.reactDirectHost)) {
    const direct = !!reactHost;
    const wa = Object.assign({ text: 'Your Text Here', font: 'Arial Black', sz: 36 }, it.wordart || {});
    const font = direct ? null : ui.select(G.FONTS().map((f) => [f, f]), wa.font, () => upd(), { id: 'wa-font', translate: 'no' });
    const size = direct ? null : ui.select([8, 10, 12, 14, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96].map((x) => [x, String(x)]), wa.sz || 36, null, { id: 'wa-size' });
    const b = direct ? null : ui.check('&Bold', !!wa.b, () => upd()), i = direct ? null : ui.check('&Italic', !!wa.i, () => upd());
    const ta = direct ? null : h('textarea', { rows: 5, id: 'wa-text', style: 'width:100%;font-size:20px' });
    if (ta) ta.value = wa.text;
    const upd = () => { if (ta && font && b && i) { ta.style.fontFamily = L.fontStack(font.value); ta.style.fontWeight = b.input.checked ? 'bold' : 'normal'; ta.style.fontStyle = i.input.checked ? 'italic' : 'normal'; } };
    if (!direct) upd();
    let reactValue = { text: wa.text, font: wa.font, sz: wa.sz || 36, b: !!wa.b, i: !!wa.i };
    const sizes = [8, 10, 12, 14, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
    const reactRoot = h('div', { id: 'react-wordx-lettersart-text-root' });
    G.reactLettersArtTextCatalog = () => ({ ...reactValue, fonts: G.FONTS(), sizes });
    G.reactLettersArtTextAction = (action, value) => {
      if (action === 'change') { reactValue = { ...reactValue, ...value }; return G.reactLettersArtTextCatalog(); }
      if (action === 'accept') done({ text: reactValue.text || ' ', font: reactValue.font, sz: +reactValue.sz, b: reactValue.b || undefined, i: reactValue.i || undefined });
      return G.reactLettersArtTextCatalog();
    };
    ui.dialog({ title: 'Edit LettersArt Text', width: 440, focus: direct ? null : ta, body: direct ? reactRoot : h('div', { class: 'col' }, G.row(G.f('&Font:', font), G.f('&Size:', size), b, i), h('label', { for: 'wa-text', html: ui.ampHTML('&Text:') }), ta), buttons: [{ label: 'OK', primary: true, onClick: () => { if (direct) G.reactLettersArtTextAction('accept'); else done({ text: ta.value || ' ', font: font.value, sz: +size.value, b: b.input.checked || undefined, i: i.input.checked || undefined }); } }, { label: 'Cancel' }] });
  }

  /* ================= Table dialogs ================= */
  G.insertTable = function () {
    const last = L.store.get('tableDims', { c: 5, r: 2 });
    const direct = !!(window.VO && VO.reactDirectHost);
    const cols = G.num(last.c, { min: 1, max: 63, id: 'it-c' }), rows = G.num(last.r, { min: 1, max: 32767, id: 'it-r' });
    const fit = G.radios('it-fit', [['fixed', 'Fi&xed column width:'], ['contents', 'Auto&Fit to contents'], ['window', 'AutoFit to win&dow']], 'fixed');
    const cw = G.len(null, { id: 'it-cw', autoLabel: L.N('Auto'), min: 9 });
    let style = null, look = null;
    const styleNote = h('div', { class: 'tp-note', text: L.t('Table style: {0}', 'Table Grid') });
    const remember = ui.check('Remember dime&nsions for new tables', false, null);
    const fEls = fit.querySelectorAll('.ck');
    const root = h('div', { id: 'react-wordx-insert-table-root' });
    G.reactInsertTableState = () => ({ columns: last.c, rows: last.r });
    G.reactInsertTableAction = (action, options) => {
      if (action === 'autoFormat') {
        G.tableAutoFormat({ pick: (name, selectedLook) => {
          style = name; look = selectedLook;
          window.dispatchEvent(new CustomEvent('wordx-insert-table-style', { detail: { name } }));
        } });
        return;
      }
      if (action !== 'insert') return;
      if (options.remember) L.store.set('tableDims', { c: options.columns, r: options.rows });
      cw.input.value = options.width || '';
      const insertOptions = { autofit: options.fit };
      const width = cw.get();
      if (options.fit === 'fixed' && width) insertOptions.colWidth = width;
      if (style) { insertOptions.style = L.tables.styleId(style); insertOptions.look = look; }
      L.tables.insert(options.rows, options.columns, insertOptions);
      if (style) L.tables.applyStyle(style, look);
    };
    ui.dialog({
      title: 'Insert Table', width: 330, focus: direct ? root : cols.input,
      body: direct ? root : h('div', { class: 'col' }, ui.group('Table size', G.f('Number of &columns:', cols, 'wide'), G.f('Number of &rows:', rows, 'wide')), ui.group('AutoFit behavior', G.row(fEls[0], cw), fEls[1], fEls[2]), G.row(styleNote, ui.button('&AutoFormat...', () => G.tableAutoFormat({ pick: (s, lk) => { style = s; look = lk; styleNote.textContent = L.t('Table style: {0}', s); } }), { class: 'btn small' })), remember),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        if (direct) {
          const get = (selector, fallback) => root.querySelector(selector)?.value || fallback;
          G.reactInsertTableAction('insert', { columns: +get('#react-wordx-table-columns', '5'), rows: +get('#react-wordx-table-rows', '2'), fit: root.querySelector('[data-table-fit]:checked')?.getAttribute('data-table-fit') || 'fixed', width: get('#react-wordx-table-width', ''), remember: !!root.querySelector('input[type=checkbox]')?.checked });
          return;
        }
        if (remember.input.checked) L.store.set('tableDims', { c: cols.get(), r: rows.get() });
        const o = { autofit: fit.get() };
        const w = cw.get();
        if (fit.get() === 'fixed' && w) o.colWidth = w;
        if (style) { o.style = L.tables.styleId(style); o.look = look; }
        L.tables.insert(rows.get(), cols.get(), o);
        if (style) L.tables.applyStyle(style, look);
      } }, { label: 'Cancel' }],
    });
  };
  G.insertCells = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const items = [['right', 'Shift cells right'], ['down', 'Shift cells down'], ['row', 'Insert entire row'], ['col', 'Insert entire column']];
    const deleteItems = [['left', 'Shift cells left'], ['up', 'Shift cells up'], ['row', 'Delete entire row'], ['col', 'Delete entire column']];
    const r = direct ? null : G.radios('ic', [['right', 'Shift cells r&ight'], ['down', 'Shift cells &down'], ['row', 'Insert entire &row'], ['col', 'Insert entire &column']], 'down');
    const root = h('div', { id: 'react-wordx-cell-operation-root' });
    G.reactCellDialogCatalog = (kind) => kind === 'delete' ? deleteItems : items;
    G.reactCellDialogKind = 'insert';
    ui.dialog({ title: 'Insert Cells', width: 240, body: direct ? root : r, buttons: [{ label: 'OK', primary: true, onClick: () => L.tables.insertCells(direct ? root.querySelector('[data-cell-operation]:checked')?.getAttribute('data-cell-operation') || 'down' : r.get()) }, { label: 'Cancel' }] });
  };
  G.deleteCells = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    G.reactCellDialogCatalog = () => [['left', 'Shift cells left'], ['up', 'Shift cells up'], ['row', 'Delete entire row'], ['col', 'Delete entire column']];
    G.reactCellDialogKind = 'delete';
    const r = direct ? null : G.radios('dc', [['left', 'Shift cells &left'], ['up', 'Shift cells &up'], ['row', 'Delete entire &row'], ['col', 'Delete entire &column']], 'left');
    const root = h('div', { id: 'react-wordx-cell-operation-root' });
    G.reactCellDialogKind = 'delete';
    ui.dialog({ title: 'Delete Cells', width: 240, body: direct ? root : r, buttons: [{ label: 'OK', primary: true, onClick: () => L.tables.deleteCells(direct ? root.querySelector('[data-cell-operation]:checked')?.getAttribute('data-cell-operation') || 'left' : r.get()) }, { label: 'Cancel' }] });
  };
  G.splitCells = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const c = G.num(2, { min: 1, max: 63, id: 'sc-c' }), r = G.num(1, { min: 1, max: 100, id: 'sc-r' });
    const merge = ui.check('&Merge cells before split', !!E.cellSel, null);
    const root = h('div', { id: 'react-wordx-split-cells-root' });
    G.reactSplitCellsState = () => ({ columns: 2, rows: 1, merge: !!E.cellSel });
    ui.dialog({ title: 'Split Cells', width: 260, focus: direct ? root : c.input, body: direct ? root : h('div', { class: 'col' }, G.f('Number of &columns:', c, 'wide'), G.f('Number of &rows:', r, 'wide'), merge), buttons: [{ label: 'OK', primary: true, onClick: () => L.tables.split(direct ? +(root.querySelector('#react-wordx-split-rows')?.value || 1) : r.get(), direct ? +(root.querySelector('#react-wordx-split-columns')?.value || 2) : c.get(), direct ? !!root.querySelector('input[type=checkbox]')?.checked : merge.input.checked) }, { label: 'Cancel' }] });
  };
  G.tableAutoFormat = function (opts) {
    opts = opts || {};
    const names = ['Table Normal', 'Table Grid'].concat(Object.keys(L.tables.STYLES).filter((n) => n !== 'Table Grid'));
    const ci = L.tables.cur();
    const curStyle = ci ? (doc().styles[ci.tbl.tblPr.style] || {}).name || 'Table Grid' : 'Table Grid';
    const look = Object.assign({ firstRow: true, firstCol: true, lastRow: false, lastCol: false }, ci ? ci.tbl.tblPr.look || {} : {});
    let reactName = names.includes(curStyle) ? curStyle : 'Table Grid';
    let reactLook = { ...look };
    const direct = !!(window.VO && VO.reactDirectHost);
    const lb = G.listBox(names, reactName, { height: 200, translate: 'no', onChange: () => draw() });
    const pv = h('div', { class: 'preview-box', style: 'width:220px;height:170px' });
    const chk = (k, l) => ui.check(l, !!look[k], (v) => { look[k] = v; draw(); });
    const checks = [chk('firstRow', '&Heading rows'), chk('firstCol', 'Fir&st column'), chk('lastRow', '&Last row'), chk('lastCol', 'Last co&lumn')];
    function previewModel(name, selectedLook) {
      const def = L.tables.STYLES[name];
      const grid = [['', L.N('Jan'), L.N('Feb'), L.N('Mar'), L.N('Total')], [L.N('East'), '7', '7', '5', '19'], [L.N('West'), '6', '4', '7', '17'], [L.N('South'), '8', '7', '9', '24'], [L.N('Total'), '21', '18', '21', '60']];
      const bd = (b) => (b && b.val !== 'nil' ? `${Math.max(1, b.sz || 0.5)}px ${b.val === 'double' ? 'double' : b.val === 'dotted' ? 'dotted' : 'solid'} ${b.color && b.color !== 'auto' ? '#' + b.color : '#000'}` : 'none');
      return grid.map((row, ri) => row.map((v, cix) => {
          const conds = R.cellConds({ tblPr: {} }, ri, cix, grid.length, row.length, Object.assign({ noVBand: !/Columns/.test(name), noHBand: !/List|Contemporary|Subtle|3D/.test(name) }, selectedLook));
          let rr = Object.assign({}, def && def.rPr), tc = Object.assign({}, def && def.tcPr), tb = Object.assign({}, def && def.tblPr && def.tblPr.borders);
          if (name === 'Table Grid') tb = { top: { val: 'single' }, bottom: { val: 'single' }, left: { val: 'single' }, right: { val: 'single' }, insideH: { val: 'single' }, insideV: { val: 'single' } };
          let cb = {};
          for (const c of conds) { const cf = def && def.cond && def.cond[c]; if (!cf) continue; if (cf.rPr) rr = Object.assign(rr, cf.rPr); if (cf.tcPr) { tc = Object.assign(tc, cf.tcPr); if (cf.tcPr.borders) cb = Object.assign(cb, cf.tcPr.borders); } }
          const side = (k, outer, inner, isOuter) => cb[k] || (isOuter ? tb[outer] : tb[inner]);
          const css = [`border-top:${bd(side('top', 'top', 'insideH', ri === 0))}`, `border-bottom:${bd(side('bottom', 'bottom', 'insideH', ri === grid.length - 1))}`, `border-left:${bd(side('left', 'left', 'insideV', cix === 0))}`, `border-right:${bd(side('right', 'right', 'insideV', cix === row.length - 1))}`, 'padding:2px 4px'];
          if (tc.shd) css.push('background:' + R.shadeCSS(tc.shd));
          if (rr.b) css.push('font-weight:bold'); if (rr.i) css.push('font-style:italic'); if (rr.color) css.push('color:#' + rr.color); if (rr.caps) css.push('text-transform:uppercase');
          return { text: v, style: css.join(';') };
      }));
    }
    G.reactTableAutoFormatCatalog = () => ({ names: names.slice(), selected: reactName, look: { ...reactLook } });
    G.reactTableAutoFormatPreview = (name, selectedLook) => previewModel(name || reactName, selectedLook || reactLook);
    G.reactTableAutoFormatAction = (action, next) => {
      if (action === 'change' && next) { reactName = next.selected; reactLook = { ...next.look }; return; }
      if (action !== 'apply') return;
      if (opts.pick) { opts.pick(reactName, Object.assign({}, reactLook, { noVBand: true })); return; }
      L.tables.applyStyle(reactName, Object.assign({}, reactLook, { noVBand: !/Columns/.test(reactName) }), { clearDirect: true });
    };
    function draw() {
      L.clear(pv);
      const t = h('table', { style: 'border-collapse:collapse;font:10px Arial,Arimo,sans-serif;width:200px' });
      previewModel(lb.get(), look).forEach((row) => {
        const tr = h('tr');
        row.forEach((cell) => tr.appendChild(h('td', { text: cell.text, style: cell.style })));
        t.appendChild(tr);
      });
      pv.appendChild(t);
    }
    draw();
    const reactRoot = h('div', { id: 'react-wordx-table-autoformat-root' });
    ui.dialog({
      title: 'Table AutoFormat', width: 480, body: direct ? reactRoot : h('div', { class: 'col' }, h('div', { text: 'Category: All table styles' }), G.row(h('div', { class: 'col', style: 'width:200px' }, h('label', { text: 'Table styles:' }), lb), h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: 'Preview' }), pv)), ui.group('Apply special formats to', G.row(...checks))),
      buttons: [{ label: '&Apply', primary: true, onClick: () => { if (direct) { G.reactTableAutoFormatAction('apply'); return; } if (opts.pick) { opts.pick(lb.get(), Object.assign({}, look, { noVBand: true })); return; } L.tables.applyStyle(lb.get(), Object.assign({}, look, { noVBand: !/Columns/.test(lb.get()) }), { clearDirect: true }); } }, { label: 'Cancel' }],
    });
  };
  G.tableProps = function (tab) {
    const ci = L.tables.cur();
    if (!ci) return;
    const t = ci.tbl;
    const d = doc();
    const { tp } = R.tblProps(d, t);
    const rg = L.tables.range();
    const total = t.grid.reduce((a, b) => a + b, 0);
    const direct = !!(window.VO && VO.reactDirectHost);
    if (direct) {
      const row = t.rows[rg.r0], cell = ci.cell, cellMap = rg.map[ci.ri].find((m) => m.cell === cell);
      const cellWidth = t.grid.slice(cellMap.c0, cellMap.c0 + cellMap.span).reduce((a, b) => a + b, 0);
      let state = {
        tab: tab || 0, tableWidthEnabled: !!(t.tblPr.w && t.tblPr.w.type !== 'auto' && t.tblPr.w.v), tableWidth: t.tblPr.w && t.tblPr.w.type === 'dxa' ? t.tblPr.w.v : total,
        tableWidthUnit: t.tblPr.w && t.tblPr.w.type === 'pct' ? 'pct' : 'dxa', alignment: tp.jc === 'center' ? 'center' : tp.jc === 'right' || tp.jc === 'end' ? 'right' : 'left', indent: tp.ind || 0,
        wrap: t.tblPr.float ? 'around' : 'none', margins: { ...tp.cellMar }, spacing: t.tblPr.cellSpacing || 0, autoResize: t.tblPr.layout !== 'fixed',
        rowHeightEnabled: !!row.trPr.h, rowHeight: row.trPr.h || 14.4, rowHeightRule: row.trPr.hRule === 'exact' ? 'exact' : 'atLeast', allowRowBreak: !row.trPr.cantSplit, repeatHeader: !!row.trPr.header,
        columnWidthEnabled: true, columnWidth: t.grid.slice(rg.c0, rg.c1 + 1).reduce((a, b) => a + b, 0) / (rg.c1 - rg.c0 + 1),
        cellWidthEnabled: true, cellWidth, verticalAlignment: cell.tcPr.vAlign || 'top', sameCellMargins: !cell.tcPr.mar,
        cellMargins: { t: cell.tcPr.mar?.t ?? tp.cellMar.t ?? 0, b: cell.tcPr.mar?.b ?? tp.cellMar.b ?? 0, l: cell.tcPr.mar?.l ?? tp.cellMar.l ?? 0, r: cell.tcPr.mar?.r ?? tp.cellMar.r ?? 0 },
        wrapText: !cell.tcPr.noWrap, fitText: !!cell.tcPr.fitText,
        rowLabel: rg.r1 > rg.r0 ? L.t('Rows {0}-{1}:', rg.r0 + 1, rg.r1 + 1) : L.t('Row {0}:', rg.r0 + 1),
        columnLabel: rg.c1 > rg.c0 ? L.t('Columns {0}-{1}:', rg.c0 + 1, rg.c1 + 1) : L.t('Column {0}:', rg.c0 + 1),
      };
      const root = h('div', { id: 'react-wordx-table-properties-root' });
      G.reactTablePropsCatalog = () => ({ ...state, margins: { ...state.margins }, cellMargins: { ...state.cellMargins } });
      G.reactTablePropsAction = (action, value) => {
        if (action === 'change' && value) { state = { ...state, ...value, margins: { ...state.margins, ...(value.margins || {}) }, cellMargins: { ...state.cellMargins, ...(value.cellMargins || {}) } }; return G.reactTablePropsCatalog(); }
        if (action === 'borders') { G.borders(); return G.reactTablePropsCatalog(); }
        if (action !== 'apply') return G.reactTablePropsCatalog();
        E.edit('Table Properties', () => {
          D.touchTbl(t);
          const tpp = t.tblPr;
          if (state.tableWidthEnabled) {
            if (state.tableWidthUnit === 'pct') tpp.w = { type: 'pct', v: 5000 };
            else { const w = state.tableWidth; tpp.w = { type: 'dxa', v: w }; const tot = t.grid.reduce((a, b) => a + b, 0) || 1; t.grid = t.grid.map((g) => L.round((g * w) / tot, 2)); }
          } else tpp.w = { type: 'auto', v: 0 };
          tpp.jc = state.alignment === 'left' ? undefined : state.alignment;
          if (!tpp.jc) delete tpp.jc;
          if (state.alignment === 'left' && state.indent) tpp.ind = state.indent; else delete tpp.ind;
          // The legacy dialog exposes this choice but does not write tblpPr.
          // Preserve that behavior until the editor has a complete positioning model.
          tpp.cellMar = { ...state.margins };
          if (state.spacing) tpp.cellSpacing = state.spacing; else delete tpp.cellSpacing;
          tpp.layout = state.autoResize ? 'autofit' : 'fixed';
          for (let r = rg.r0; r <= rg.r1; r++) {
            const tr = t.rows[r].trPr;
            if (state.rowHeightEnabled) { tr.h = state.rowHeight; tr.hRule = state.rowHeightRule; } else { delete tr.h; delete tr.hRule; }
            if (state.allowRowBreak) delete tr.cantSplit; else tr.cantSplit = true;
          }
          if (state.repeatHeader) for (let r = 0; r <= rg.r1; r++) t.rows[r].trPr.header = true; else for (let r = rg.r0; r <= rg.r1; r++) delete t.rows[r].trPr.header;
          if (state.columnWidthEnabled) for (let i = rg.c0; i <= rg.c1; i++) t.grid[i] = L.round(state.columnWidth, 2);
          for (const c of rg.cells) {
            c.tcPr.vAlign = state.verticalAlignment === 'top' ? undefined : state.verticalAlignment;
            if (!c.tcPr.vAlign) delete c.tcPr.vAlign;
            if (state.sameCellMargins) delete c.tcPr.mar; else c.tcPr.mar = { ...state.cellMargins };
            if (state.wrapText) delete c.tcPr.noWrap; else c.tcPr.noWrap = true;
            if (state.fitText) c.tcPr.fitText = true; else delete c.tcPr.fitText;
          }
          if (state.cellWidthEnabled && rg.cells.length === 1 && !state.columnWidthEnabled) {
            const want = state.cellWidth, have = cellWidth, k = want / (have || 1);
            for (let i = cellMap.c0; i < cellMap.c0 + cellMap.span; i++) t.grid[i] = L.round(t.grid[i] * k, 2);
          }
          L.tables.fixWidths(t);
          return E.sel;
        });
      };
      ui.dialog({ title: 'Table Properties', width: 430, body: root, buttons: [{ label: 'OK', primary: true, onClick: () => G.reactTablePropsAction('apply') }, { label: 'Cancel' }] });
      return;
    }
    /* Table tab */
    const tW = ui.check('&Preferred width:', !!(t.tblPr.w && t.tblPr.w.type !== 'auto' && t.tblPr.w.v), null);
    const tWv = G.len(t.tblPr.w && t.tblPr.w.type === 'dxa' ? t.tblPr.w.v : total, { id: 'tp-w', min: 9 });
    const tWu = ui.select([['dxa', 'Inches/cm'], ['pct', 'Percent']], t.tblPr.w && t.tblPr.w.type === 'pct' ? 'pct' : 'dxa', null, { id: 'tp-wu' });
    const align = G.radios('tp-al', [['left', '&Left'], ['center', '&Center'], ['right', 'Rig&ht']], tp.jc === 'center' ? 'center' : tp.jc === 'right' || tp.jc === 'end' ? 'right' : 'left');
    const indent = G.len(tp.ind || 0, { id: 'tp-ind' });
    const wrap = G.radios('tp-wr', [['none', '&None'], ['around', 'Aro&und']], t.tblPr.float ? 'around' : 'none');
    const mar = { t: G.len(tp.cellMar.t || 0, { id: 'tp-mt', min: 0 }), b: G.len(tp.cellMar.b || 0, { id: 'tp-mb', min: 0 }), l: G.len(tp.cellMar.l || 0, { id: 'tp-ml', min: 0 }), r: G.len(tp.cellMar.r || 0, { id: 'tp-mr', min: 0 }) };
    const spacing = G.len(t.tblPr.cellSpacing || 0, { id: 'tp-sp', min: 0 });
    const autoResize = ui.check('Automatically resi&ze to fit contents', t.tblPr.layout !== 'fixed', null);
    const optsBtn = ui.button('&Options...', () => ui.dialog({ title: 'Table Options', width: 320, body: h('div', { class: 'col' }, ui.group('Default cell margins', G.row(G.f('&Top:', mar.t, 'narrow'), G.f('&Left:', mar.l, 'narrow')), G.row(G.f('&Bottom:', mar.b, 'narrow'), G.f('&Right:', mar.r, 'narrow'))), ui.group('Default cell spacing', G.f('Allow spacing between cells', spacing, 'wide')), ui.group('Options', autoResize)), buttons: [{ label: 'OK', primary: true }] }), { class: 'btn small' });
    /* Row tab */
    const row = t.rows[rg.r0];
    const rH = ui.check('&Specify height:', !!row.trPr.h, null);
    const rHv = G.len(row.trPr.h || 14.4, { id: 'tp-rh', min: 1 });
    const rRule = ui.select([['atLeast', 'At least'], ['exact', 'Exactly']], row.trPr.hRule === 'exact' ? 'exact' : 'atLeast', null, { id: 'tp-rr' });
    const rSplit = ui.check('Allow row to brea&k across pages', !row.trPr.cantSplit, null);
    const rHead = ui.check('Repeat as &header row at the top of each page', !!row.trPr.header, null);
    /* Column tab */
    const cW = ui.check('Preferred &width:', true, null);
    const cWv = G.len(t.grid.slice(rg.c0, rg.c1 + 1).reduce((a, b) => a + b, 0) / (rg.c1 - rg.c0 + 1), { id: 'tp-cw', min: 6 });
    /* Cell tab */
    const cell = ci.cell;
    const ceW = ui.check('Preferred &width:', true, null);
    const ceWv = G.len(t.grid.slice(rg.map[ci.ri].find((m) => m.cell === cell).c0, rg.map[ci.ri].find((m) => m.cell === cell).c0 + (cell.tcPr.span || 1)).reduce((a, b) => a + b, 0), { id: 'tp-cew', min: 6 });
    const va = G.radios('tp-va', [['top', 'To&p'], ['center', '&Center'], ['bottom', '&Bottom']], cell.tcPr.vAlign || 'top');
    const cm = cell.tcPr.mar || {};
    const cmSame = ui.check('Same as the whole &table', !cell.tcPr.mar, null);
    const cmar = { t: G.len(cm.t != null ? cm.t : tp.cellMar.t || 0, { id: 'tp-cmt', min: 0 }), b: G.len(cm.b != null ? cm.b : tp.cellMar.b || 0, { id: 'tp-cmb', min: 0 }), l: G.len(cm.l != null ? cm.l : tp.cellMar.l || 0, { id: 'tp-cml', min: 0 }), r: G.len(cm.r != null ? cm.r : tp.cellMar.r || 0, { id: 'tp-cmr', min: 0 }) };
    const noWrap = ui.check('&Wrap text', !cell.tcPr.noWrap, null);
    const fitText = ui.check('&Fit text', !!cell.tcPr.fitText, null);
    const cellOpts = ui.button('&Options...', () => ui.dialog({ title: 'Cell Options', width: 320, body: h('div', { class: 'col' }, ui.group('Cell margins', cmSame, G.row(G.f('Top:', cmar.t, 'narrow'), G.f('Left:', cmar.l, 'narrow')), G.row(G.f('Bottom:', cmar.b, 'narrow'), G.f('Right:', cmar.r, 'narrow'))), ui.group('Options', noWrap, fitText)), buttons: [{ label: 'OK', primary: true }] }), { class: 'btn small' });
    const tabs = ui.tabs([
      { label: '&Table', body: h('div', { class: 'col' }, ui.group('Size', G.row(tW, tWv, G.f('Measure in:', tWu))), ui.group('Alignment', G.row(align, G.f('&Indent from left:', indent, 'wide'))), ui.group('Text wrapping', wrap), h('div', { class: 'tp-btns' }, ui.button('&Borders and Shading...', () => G.borders(), { class: 'btn small' }), optsBtn)) },
      { label: '&Row', body: h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: rg.r1 > rg.r0 ? L.t('Rows {0}-{1}:', rg.r0 + 1, rg.r1 + 1) : L.t('Row {0}:', rg.r0 + 1) }), ui.group('Size', G.row(rH, rHv, G.f('Row height &is:', rRule))), ui.group('Options', rSplit, rHead)) },
      { label: 'Col&umn', body: h('div', { class: 'col' }, h('div', { class: 'cd-cap', text: rg.c1 > rg.c0 ? L.t('Columns {0}-{1}:', rg.c0 + 1, rg.c1 + 1) : L.t('Column {0}:', rg.c0 + 1) }), ui.group('Size', G.row(cW, cWv))) },
      { label: 'C&ell', body: h('div', { class: 'col' }, ui.group('Size', G.row(ceW, ceWv)), ui.group('Vertical alignment', va), h('div', { class: 'tp-btns' }, cellOpts)) },
    ], tab || 0);
    ui.dialog({
      title: 'Table Properties', width: 430, body: tabs,
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        E.edit('Table Properties', () => {
          D.touchTbl(t);
          const tpp = t.tblPr;
          if (tW.input.checked) {
            if (tWu.value === 'pct') tpp.w = { type: 'pct', v: 5000 };
            else { const w = tWv.get(); tpp.w = { type: 'dxa', v: w }; const tot = t.grid.reduce((a, b) => a + b, 0) || 1; t.grid = t.grid.map((g) => L.round((g * w) / tot, 2)); }
          } else tpp.w = { type: 'auto', v: 0 };
          tpp.jc = align.get() === 'left' ? undefined : align.get();
          if (!tpp.jc) delete tpp.jc;
          const ind = indent.get();
          if (align.get() === 'left' && ind) tpp.ind = ind; else delete tpp.ind;
          tpp.cellMar = { t: mar.t.get(), b: mar.b.get(), l: mar.l.get(), r: mar.r.get() };
          const sp = spacing.get(); if (sp) tpp.cellSpacing = sp; else delete tpp.cellSpacing;
          tpp.layout = autoResize.input.checked ? 'autofit' : 'fixed';
          for (let r = rg.r0; r <= rg.r1; r++) {
            const tr = t.rows[r].trPr;
            if (rH.input.checked) { tr.h = rHv.get(); tr.hRule = rRule.value; } else { delete tr.h; delete tr.hRule; }
            if (rSplit.input.checked) delete tr.cantSplit; else tr.cantSplit = true;
          }
          if (rHead.input.checked) for (let r = 0; r <= rg.r1; r++) t.rows[r].trPr.header = true; else for (let r = rg.r0; r <= rg.r1; r++) delete t.rows[r].trPr.header;
          if (cW.input.checked) { const w = cWv.get(); for (let i = rg.c0; i <= rg.c1; i++) t.grid[i] = L.round(w, 2); }
          for (const c of rg.cells) {
            c.tcPr.vAlign = va.get() === 'top' ? undefined : va.get();
            if (!c.tcPr.vAlign) delete c.tcPr.vAlign;
            if (cmSame.input.checked) delete c.tcPr.mar; else c.tcPr.mar = { t: cmar.t.get(), b: cmar.b.get(), l: cmar.l.get(), r: cmar.r.get() };
            if (noWrap.input.checked) delete c.tcPr.noWrap; else c.tcPr.noWrap = true;
            if (fitText.input.checked) c.tcPr.fitText = true; else delete c.tcPr.fitText;
          }
          if (ceW.input.checked && rg.cells.length === 1 && !cW.input.checked) {
            const m = rg.map[ci.ri].find((x) => x.cell === cell);
            const want = ceWv.get(), have = t.grid.slice(m.c0, m.c0 + m.span).reduce((a, b) => a + b, 0);
            const k = want / (have || 1);
            for (let i = m.c0; i < m.c0 + m.span; i++) t.grid[i] = L.round(t.grid[i] * k, 2);
          }
          L.tables.fixWidths(t);
          return E.sel;
        });
      } }, { label: 'Cancel' }],
    });
  };
  G.sort = function () {
    const inTbl = L.tables.inTable();
    const ci = inTbl ? L.tables.cur() : null;
    const fields = inTbl ? ci.tbl.grid.map((_, i) => ['' + i, L.t('Column {0}', i + 1)]) : [['0', L.t('Paragraphs')]].concat([1, 2, 3, 4, 5].map((i) => ['' + i, L.t('Field {0}', i)]));
    const defaults = { inTable: inTbl, fields, keys: [{ by: fields[0][0], type: 'text', dir: 'asc' }, { by: '', type: 'text', dir: 'asc' }, { by: '', type: 'text', dir: 'asc' }], header: inTbl && ci.tbl.rows[0].trPr.header ? 'yes' : 'no', separator: '\t', caseSensitive: false };
    let reactState = defaults;
    G.reactSortCatalog = () => ({ ...reactState, fields: reactState.fields.map((field) => field.slice()), keys: reactState.keys.map((key) => ({ ...key })) });
    G.reactSortAction = (action, next) => {
      if (action === 'change' && next) { reactState = next; return; }
      if (action !== 'apply') return;
      const keys = reactState.keys.filter((key) => key.by !== '').map((key) => ({ col: +key.by, field: +key.by, type: key.type, desc: key.dir === 'desc' }));
      if (inTbl) L.tables.sortTable(keys, { header: reactState.header === 'yes', caseSensitive: reactState.caseSensitive });
      else L.tables.sortParagraphs(keys, { header: reactState.header === 'yes', sep: reactState.separator === 'other' ? ';' : reactState.separator, caseSensitive: reactState.caseSensitive });
    };
    const mk = (n, def) => ({ by: ui.select(n ? [['', '(none)']].concat(fields) : fields, def, null, { id: 'so-b' + n }), type: ui.select([['text', 'Text'], ['number', 'Number'], ['date', 'Date']], 'text', null, { id: 'so-t' + n }), dir: G.radios('so-d' + n, [['asc', '&Ascending'], ['desc', '&Descending']], 'asc') });
    const k1 = mk(0, fields[0][0]), k2 = mk(1, ''), k3 = mk(2, '');
    const hdr = G.radios('so-h', [['yes', 'Header &row'], ['no', 'No header ro&w']], inTbl && ci.tbl.rows[0].trPr.header ? 'yes' : 'no');
    const sepSel = ui.select([['\t', 'Tabs'], [',', 'Commas'], ['other', 'Other']], '\t', null, { id: 'so-sep' });
    const caseS = ui.check('Case &sensitive', false, null);
    const row = (lbl, k) => ui.group(lbl, G.row(k.by, G.f('Type:', k.type), k.dir));
    const reactRoot = h('div', { id: 'react-wordx-sort-root' });
    ui.dialog({
      title: inTbl ? 'Sort' : 'Sort Text', width: 470,
      body: window.VO && VO.reactDirectHost ? reactRoot : h('div', { class: 'col' }, row('&Sort by', k1), row('&Then by', k2), row('Then &by', k3), ui.group('My list has', hdr), G.row(G.f('Separate fields at:', sepSel), caseS)),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        if (window.VO && VO.reactDirectHost) { G.reactSortAction('apply'); return; }
        const keys = [k1, k2, k3].filter((k) => k.by.value !== '').map((k) => ({ col: +k.by.value, field: +k.by.value, type: k.type.value, desc: k.dir.get() === 'desc' }));
        if (inTbl) L.tables.sortTable(keys, { header: hdr.get() === 'yes', caseSensitive: caseS.input.checked });
        else L.tables.sortParagraphs(keys, { header: hdr.get() === 'yes', sep: sepSel.value === 'other' ? ';' : sepSel.value, caseSensitive: caseS.input.checked });
      } }, { label: 'Cancel' }],
    });
  };
  G.formula = function () {
    let guess = '=SUM(ABOVE)';
    try { if (L.tables.inTable()) { const v = L.tables.evalFormula('SUM(ABOVE)', E.sel.f.p); const l = L.tables.evalFormula('SUM(LEFT)', E.sel.f.p); if (!v && l) guess = '=SUM(LEFT)'; } else guess = '='; } catch (e) { /* default */ }
    const direct = !!(window.VO && VO.reactDirectHost);
    const inp = direct ? null : h('input', { type: 'text', id: 'fo-f', value: guess, style: 'width:100%' });
    const fmt = h('input', { type: 'text', id: 'fo-fmt', list: 'fo-fmts', style: 'width:100%' });
    const dl = h('datalist', { id: 'fo-fmts' }, ...['#,##0', '#,##0.00', '$#,##0.00;($#,##0.00)', '0', '0%', '0.00', '0.00%'].map((x) => h('option', { value: x })));
    const functions = ['ABS', 'AND', 'AVERAGE', 'COUNT', 'DEFINED', 'FALSE', 'IF', 'INT', 'MAX', 'MIN', 'MOD', 'NOT', 'OR', 'PRODUCT', 'ROUND', 'SIGN', 'SUM', 'TRUE'];
    const fn = direct ? null : ui.select([['', L.t('Paste function')]].concat(functions.map((x) => [x, x])), '', (v) => { if (!v) return; inp.value += v + '()'; inp.focus(); inp.setSelectionRange(inp.value.length - 1, inp.value.length - 1); fn.value = ''; }, { id: 'fo-fn', translate: 'no' });
    const bms = direct ? null : ui.select([['', L.t('Paste bookmark')]].concat(Array.from(F().bookmarks().keys()).filter((n) => !n.startsWith('_')).map((n) => [n, n])), '', (v) => { if (v) { inp.value += v; bms.value = ''; } }, { id: 'fo-bm', translate: 'no' });
    const root = h('div', { id: 'react-wordx-formula-root' });
    G.reactFormulaCatalog = () => ({ guess, functions, bookmarks: Array.from(F().bookmarks().keys()).filter((n) => !n.startsWith('_')), formats: ['#,##0', '#,##0.00', '$#,##0.00;($#,##0.00)', '0', '0%', '0.00', '0.00%'] });
    G.reactFormulaAction = (formula, numberFormat) => {
      let c = formula.trim();
      if (!c.startsWith('=')) c = '=' + c;
      if (numberFormat.trim()) c += ` \\# "${numberFormat.trim()}"`;
      A().insertFieldNow(c, L.N('Formula'));
    };
    ui.dialog({ title: 'Formula', width: 340, focus: direct ? root : inp, body: direct ? root : h('div', { class: 'col' }, h('label', { for: 'fo-f', html: ui.ampHTML('&Formula:') }), inp, dl, h('label', { for: 'fo-fmt', html: ui.ampHTML('&Number format:') }), fmt, G.row(fn, bms)), buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (direct) G.reactFormulaAction(root.querySelector('#react-wordx-formula-input')?.value || '', root.querySelector('#react-wordx-formula-format')?.value || '');
      else G.reactFormulaAction(inp.value, fmt.value);
    } }, { label: 'Cancel' }] });
  };
  G.textToTable = function () {
    const paras = E.selectedParas();
    const direct = !!(window.VO && VO.reactDirectHost);
    const initialSeparator = paras.some((p) => p.runs.some((it) => it.t === 'tab')) ? '\t' : paras.some((p) => D.plainText(p).includes(',')) ? ',' : 'para';
    const initialColumns = initialSeparator === 'para' ? 1 : Math.max(1, ...paras.map((p) => initialSeparator === '\t' ? p.runs.filter((it) => it.t === 'tab').length + 1 : D.plainText(p).split(initialSeparator).length));
    const sepR = direct ? null : G.radios('tt-s', [['para', '&Paragraphs'], [',', '&Commas'], ['\t', '&Tabs'], ['other', 'O&ther:']], initialSeparator, () => upd());
    const other = h('input', { type: 'text', maxlength: 1, value: '-', style: 'width:30px' });
    const cols = G.num(1, { min: 1, max: 63, id: 'tt-c', onChange: () => upd(true) });
    const rowsN = h('span');
    const fit = G.radios('tt-fit', [['fixed', 'Fi&xed column width'], ['contents', 'Auto&Fit to contents'], ['window', 'AutoFit to win&dow']], 'fixed');
    const root = h('div', { id: 'react-wordx-text-to-table-root' });
    G.reactTextToTableState = () => ({ initialSeparator, columns: direct ? initialColumns : cols.get(), paragraphs: paras.map((p) => ({ text: D.plainText(p), tabs: p.runs.filter((it) => it.t === 'tab').length })) });
    G.reactTextToTableAction = (options) => L.tables.textToTable(options);
    const upd = (fromCols) => {
      const sep = sepR.get() === 'other' ? other.value || '-' : sepR.get();
      if (!fromCols) cols.set(sep === 'para' ? 1 : Math.max(...paras.map((p) => (sep === '\t' ? p.runs.filter((it) => it.t === 'tab').length : (D.plainText(p).split(sep).length - 1)) + 1)));
      rowsN.textContent = String(sep === 'para' ? Math.ceil(paras.length / cols.get()) : paras.length);
    };
    if (!direct) upd();
    const sEls = sepR?.querySelectorAll('.ck');
    ui.dialog({ title: 'Convert Text to Table', width: 330, body: direct ? root : h('div', { class: 'col' }, ui.group('Table size', G.f('Number of &columns:', cols, 'wide'), G.f('Number of rows:', rowsN, 'wide')), ui.group('AutoFit behavior', fit), ui.group('Separate text at', G.row(sEls[0], sEls[1]), G.row(sEls[2], sEls[3], other))), buttons: [{ label: 'OK', primary: true, onClick: () => {
      if (direct) {
        const mode = root.querySelector('[data-text-separator]:checked')?.getAttribute('data-text-separator') || initialSeparator;
        const sep = mode === 'other' ? root.querySelector('#react-wordx-text-separator')?.value || '-' : mode === 'paragraphs' ? 'para' : mode === 'tabs' ? '\t' : ',';
        G.reactTextToTableAction({ sep, cols: +(root.querySelector('#react-wordx-text-table-columns')?.value || 1), autofit: root.querySelector('[data-text-fit]:checked')?.getAttribute('data-text-fit') || 'fixed' });
      } else L.tables.textToTable({ sep: sepR.get() === 'other' ? other.value || '-' : sepR.get(), cols: cols.get(), autofit: fit.get() });
    } }, { label: 'Cancel' }] });
  };
  G.tableToText = function () {
    const direct = !!(window.VO && VO.reactDirectHost);
    const r = G.radios('tx-s', [['para', '&Paragraph marks'], ['\t', '&Tabs'], [',', '&Commas'], ['other', '&Other:']], '\t');
    const other = h('input', { type: 'text', maxlength: 1, value: '-', style: 'width:30px' });
    const root = h('div', { id: 'react-wordx-table-to-text-root' });
    ui.dialog({ title: 'Convert Table to Text', width: 280, body: direct ? root : ui.group('Separate text with', r, other), buttons: [{ label: 'OK', primary: true, onClick: () => {
      const mode = direct ? root.querySelector('[data-table-text-separator]:checked')?.getAttribute('data-table-text-separator') || 'tabs' : r.get();
      const sep = direct ? mode === 'para' ? 'para' : mode === 'tabs' ? '\t' : mode === 'commas' ? ',' : 'other' : mode;
      const custom = direct ? root.querySelector('#react-wordx-table-text-separator')?.value || '-' : other.value || '-';
      L.tables.tableToText(sep === 'other' ? custom : sep);
    } }, { label: 'Cancel' }] });
  };
})();
