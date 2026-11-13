/* Excelx — dialogs, part 3 (formulas and printing): Insert Function, Function Arguments, names, Evaluate Formula,
 * Watch Window, Error Checking, Spelling, Page Setup with Header/Footer, Print. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, O = L.ops, F = L.formula, NF = L.numfmt, LY = L.layout, C = L.calc;
  const { h } = L;
  const ui = L.ui;
  const D = (L.dlg = L.dlg || {});
  let reactFunction = null, reactFunctionAction = null, reactFunctionArgs = null, reactFunctionArgsAction = null;
  D.reactFunctionCatalog = () => reactFunction;
  D.reactFunctionAction = (action, value) => reactFunctionAction ? reactFunctionAction(action, value) : false;
  D.reactFunctionArgsCatalog = () => reactFunctionArgs;
  D.reactFunctionArgsAction = (action, index, value) => reactFunctionArgsAction ? reactFunctionArgsAction(action, index, value) : false;
  let reactCustomHF = null, reactCustomHFDone = null;
  D.reactCustomHFCatalog = () => reactCustomHF;
  D.reactCustomHFApply = (sections) => {
    if (!reactCustomHF || !reactCustomHFDone || !sections) return false;
    const code = ['l', 'c', 'r'].map((key) => sections[key] ? '&' + key.toUpperCase() + sections[key] : '').join('');
    const done = reactCustomHFDone; reactCustomHF = reactCustomHFDone = null; L.bus.emit('excelx-custom-hf-dialog'); done(code); return true;
  };
  D.reactCustomHFClose = () => { reactCustomHF = reactCustomHFDone = null; L.bus.emit('excelx-custom-hf-dialog'); };
  const G = () => L.grid;
  const A = () => L.app;
  const E = () => L.editor;
  const sh = () => G().sheet();
  const wb = () => G().wb;
  const tryRun = (fn) => { try { fn(); } catch (e) { A().error(e); } G().paint(); G().syncObjects(); ui.refresh(); };
  /** a value as Excel shows it in dialogs: "text", 12.5, TRUE, {1,2;3,4} */
  function show(v, depth) {
    if (v instanceof C.Ref) {
      if (v.single) return show(C.cellValue(v.sheet, v.r1, v.c1));
      const rows = [];
      for (let r = v.r1; r <= Math.min(v.r2, v.r1 + 5); r++) { const row = []; for (let c = v.c1; c <= Math.min(v.c2, v.c1 + 5); c++) row.push(show(C.cellValue(v.sheet, r, c), 1)); rows.push(row.join(',') + (v.c2 > v.c1 + 5 ? ',...' : '')); }
      return '{' + rows.join(';') + (v.r2 > v.r1 + 5 ? ';...' : '') + '}';
    }
    if (v instanceof C.RefList) return v.list.map((x) => show(x)).join(',');
    if (v && v.rows && Array.isArray(v.rows)) { const rows = v.rows.slice(0, 6).map((r) => r.slice(0, 6).map((x) => show(x, 1)).join(',')); return '{' + rows.join(';') + '}'; }
    if (v == null) return depth ? '0' : '';
    if (M.isErr(v)) return v.e;
    if (typeof v === 'string') return '"' + v + '"';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return NF.general(v, 15, true);
    return String(v);
  }
  D.showValue = show;
  const evalText = (text) => {
    const s = sh(), sel = G().sel();
    try { const ast = F.parse(String(text).replace(/^=/, '')); return C.ev(ast, C.ctx(wb(), s, sel.r, sel.c)); } catch (e) { return null; }
  };

  /* ================================================================ Insert Function */
  const MRU = () => L.store.get('fnMRU', ['SUM', 'AVERAGE', 'IF', 'HYPERLINK', 'COUNT', 'MAX', 'SIN', 'SUMIF', 'PMT', 'STDEV']);
  D.insertFunction = function (prefer) {
    const FI = L.fninfo;
    if (root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected) {
      let category = prefer ? 'All' : 'Most Recently Used', selected = prefer || null, search = '';
      const namesFor = (cat, query) => {
        if (query) {
          const words = query.toLowerCase().split(/[^\p{L}\p{N}_]+/u).filter((word) => word.length > 2 || /[^\x00-\x7f]/.test(word));
          if (!words.length) return [];
          return FI.all().map((fn) => { const text = (fn.name + ' ' + fn.desc + ' ' + L.t(fn.desc)).toLowerCase(); let score = 0; for (const word of words) { if (fn.name.toLowerCase() === word) score += 10; if (text.includes(word)) score += 2; if (text.includes(word.replace(/s$/, ''))) score += 1; } return [fn.name, score]; }).filter((item) => item[1] > 0).sort((a, b) => b[1] - a[1]).slice(0, 30).map((item) => item[0]);
        }
        if (cat === 'Most Recently Used') return MRU().filter((name) => FI.get(name));
        return FI.all().filter((fn) => cat === 'All' || fn.cat === cat).map((fn) => fn.name).sort();
      };
      const publish = () => {
        let names = namesFor(category, search);
        if (!names.includes(selected)) selected = names[0] || null;
        const fn = selected && FI.get(selected);
        reactFunction = { categories: FI.categories, category, search, names, selected, syntax: fn ? FI.syntax(fn) : '', description: fn ? fn.desc : '' };
        L.bus.emit('excelx-insert-function-dialog');
      };
      reactFunctionAction = (action, value) => {
        if (!reactFunction) return false;
        if (action === 'close') { reactFunction = reactFunctionAction = null; L.bus.emit('excelx-insert-function-dialog'); return true; }
        if (action === 'search' && typeof value === 'string') { search = value; if (search) category = 'Recommended'; publish(); return true; }
        if (action === 'category' && (FI.categories.includes(value) || value === 'Recommended')) { category = value; if (category !== 'Recommended') search = ''; publish(); return true; }
        if (action === 'select' && typeof value === 'string' && reactFunction.names.includes(value)) { selected = value; publish(); return true; }
        if (action === 'open' && selected) { const name = selected; const recent = MRU().filter((item) => item !== name); recent.unshift(name); L.store.set('fnMRU', recent.slice(0, 10)); reactFunction = reactFunctionAction = null; L.bus.emit('excelx-insert-function-dialog'); D.functionArgs(name, true); return true; }
        return false;
      };
      publish();
      return;
    }
    let cat = prefer ? 'All' : 'Most Recently Used';
    let pick = prefer || null;
    const search = h('input', { type: 'text', placeholder: 'Type a brief description of what you want to do and then click Go', style: 'flex:1;min-width:0' });
    const listBox = h('div');
    const syntax = h('div', { class: 'if-syntax', translate: 'no' }), desc = h('div', { class: 'if-desc' });
    const describe = (name) => { const f = FI.get(name); syntax.textContent = f ? FI.syntax(f) : ''; desc.textContent = f ? L.t(f.desc) : ''; };
    let lb;
    const draw = (names) => {
      listBox.textContent = '';
      lb = D.list(names.map((n) => [n, n]), names.includes(pick) ? pick : names[0], (n) => { pick = n; describe(n); }, { class: 'lbox fnlist', translate: 'no' });
      lb.ondbl = (n) => { pick = n; d.close(1); go(); };
      listBox.appendChild(lb);
      if (!names.includes(pick)) pick = names[0] || null;
      describe(pick);
    };
    const byCat = () => {
      if (cat === 'Most Recently Used') return MRU().filter((n) => FI.get(n));
      const all = FI.all().filter((f) => cat === 'All' || f.cat === cat).map((f) => f.name);
      return all.sort();
    };
    const catSel = ui.select(FI.categories.map((c) => [c, c]), cat, (v) => { cat = v; draw(byCat()); });
    const doSearch = () => {
      const words = search.value.toLowerCase().split(/[^\p{L}\p{N}_]+/u).filter((w) => w.length > 2 || /[^\x00-\x7f]/.test(w));
      if (!words.length) return;
      const scored = FI.all().map((f) => { const t = (f.name + ' ' + f.desc + ' ' + L.t(f.desc)).toLowerCase(); let sc = 0; for (const w of words) { if (f.name.toLowerCase() === w) sc += 10; if (t.includes(w)) sc += 2; if (t.includes(w.replace(/s$/, ''))) sc += 1; } return [f.name, sc]; }).filter((x) => x[1] > 0).sort((a, b) => b[1] - a[1]).slice(0, 30).map((x) => x[0]);
      catSel.value = 'Recommended';
      draw(scored.length ? scored : []);
    };
    search.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); doSearch(); } });
    draw(byCat());
    const go = () => {
      if (!pick) return;
      const list = MRU().filter((n) => n !== pick); list.unshift(pick); L.store.set('fnMRU', list.slice(0, 10));
      D.functionArgs(pick, true);
    };
    const d = ui.dialog({ title: 'Insert Function', width: 460, body: h('div', { class: 'col' }, h('label', { text: 'Search for a function:' }), h('div', { class: 'row', style: 'gap:6px' }, search, ui.button('&Go', doSearch)), ui.field('Or select a &category:', catSel), h('label', { text: 'Select a functio&n:' }), listBox, syntax, desc),
      buttons: [{ label: 'OK', primary: true, onClick: () => go() }, { label: 'Cancel' }] });
  };

  /* ================================================================ Function Arguments */
  D.functionArgs = function (name, fresh) {
    const FI = L.fninfo;
    const ed = E();
    let existing = null;
    if (!name && ed.active) { const cf = ed.currentFunction(); if (cf) name = cf.name; }
    if (!name) { const s = sh(), sel = G().sel(); const cl = s.get(sel.r, sel.c); if (cl && cl.f != null) { const m = /^\s*(?:_xlfn\.)?([A-Z][A-Z0-9.]*)\(/i.exec(cl.f); if (m) { name = m[1].toUpperCase(); existing = F.display(cl.f); } } }
    if (!name) { D.insertFunction(); return; }
    const f = FI.get(name) || { name, args: ['value1', '[value2]', '...'], desc: '' };
    /* existing arguments of a formula consisting of this one call */
    let preset = [];
    if (existing && !fresh) { try { const ast = F.parse(existing); if (ast.t === 'fn') preset = ast.args.map((a) => F.toText(a)); } catch (e) { preset = []; } }
    const base = f.args.filter((a) => a !== '...');
    const repeating = f.args.includes('...');
    if (root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected) {
      let args = Array.from({ length: Math.max(base.length, preset.length, 1) }, (_, index) => preset[index] || '');
      let activeHelp = L.t(f.desc);
      const argName = (index) => { if (index < base.length) return base[index]; const last = base[base.length - 1] || 'value1'; const match = /^(\[?)([a-z_]+?)(\d+)(\]?)$/.exec(last); return match ? '[' + match[2] + (+match[3] + index - base.length + 1) + ']' : last; };
      const callText = () => { const values = args.map((value) => value.trim()); while (values.length && !values[values.length - 1] && values.length > base.filter((arg) => !arg.startsWith('[')).length) values.pop(); return f.name + '(' + values.join(',') + ')'; };
      const snapshot = () => {
        const call = callText(), evaluated = evalText(call);
        reactFunctionArgs = { name: f.name, description: f.desc, help: activeHelp, repeating, args: args.map((value, index) => { const name = argName(index), text = value.trim(); return { name, value, optional: name.startsWith('['), help: L.t(FI.argHelp(name)) || '', preview: text ? '= ' + show(evalText(text)) : name.startsWith('[') ? '' : '= ' }; }), call, result: '= ' + (evaluated == null ? '' : show(evaluated)), total: L.t('Formula result = {0}', evaluated == null ? '' : show(evaluated)) };
        L.bus.emit('excelx-function-arguments-dialog');
      };
      const commit = () => {
        const call = callText();
        if (ed.active) {
          const text = ed.text(), caret = G().input.selectionStart != null ? G().input.selectionStart : text.length;
          const before = text.slice(0, caret), after = text.slice(caret);
          const pattern = new RegExp('(?:_xlfn\\.)?' + f.name.replace('.', '\\.') + '\\($', 'i');
          if (pattern.test(before)) ed.setText(before.replace(pattern, '') + call + after.replace(/^\)/, ''), before.replace(pattern, '').length + call.length);
          else ed.insertText(call);
          return;
        }
        if (G().guardProtect(null, true)) return;
        ed.begin('enter', '=' + call); ed.commit();
      };
      reactFunctionArgsAction = (action, index, value) => {
        if (!reactFunctionArgs) return false;
        if (action === 'close') { reactFunctionArgs = reactFunctionArgsAction = null; L.bus.emit('excelx-function-arguments-dialog'); return true; }
        if (action === 'set' && Number.isInteger(index) && index >= 0 && index < args.length && typeof value === 'string') { args[index] = value; activeHelp = L.t(FI.argHelp(argName(index))) || L.t(f.desc); snapshot(); return true; }
        if (action === 'focus' && Number.isInteger(index) && index >= 0 && index < args.length) { activeHelp = L.t(FI.argHelp(argName(index))) || L.t(f.desc); if (repeating && index === args.length - 1) args = [...args, '']; snapshot(); return true; }
        if (action === 'apply') { commit(); reactFunctionArgs = reactFunctionArgsAction = null; L.bus.emit('excelx-function-arguments-dialog'); return true; }
        return false;
      };
      snapshot();
      return;
    }
    const rows = h('div', { class: 'fa-rows' });
    const inputs = [];
    const help = h('div', { class: 'fa-help' });
    const result = h('div', { class: 'fa-result' });
    const total = h('div', { class: 'fa-total' });
    const argName = (i) => { if (i < base.length) return base[i]; const last = base[base.length - 1] || 'value1'; const m = /^(\[?)([a-z_]+?)(\d+)(\]?)$/.exec(last); return m ? '[' + m[2] + (+m[3] + i - base.length + 1) + ']' : last; };
    const addRow = (i, value) => {
      const nm = argName(i);
      const inp = h('input', { type: 'text', value: value || '', class: 'fa-in' });
      const out = h('span', { class: 'fa-val' });
      const upd = () => { const t = inp.value.trim(); out.textContent = t ? '= ' + show(evalText(t)) : nm.startsWith('[') ? '' : '= '; refresh(); };
      inp.addEventListener('input', upd);
      inp.addEventListener('focus', () => { help.textContent = ''; help.append(h('b', { translate: 'no', text: nm.replace(/[[\]]/g, '').replace(/^./, (c) => c.toUpperCase()) + ' ' }), document.createTextNode(L.t(FI.argHelp(nm)) || '')); if (repeating && i === inputs.length - 1) addRow(i + 1, ''); });
      const label = h('label', { class: nm.startsWith('[') ? 'opt' : 'req', translate: 'no', text: nm.replace(/[[\]]/g, '').replace(/^./, (c) => c.toUpperCase()) });
      const rb = D.refBox ? null : null;
      void rb;
      rows.appendChild(h('div', { class: 'fa-row' }, label, inp, out));
      inputs.push(inp);
      upd();
    };
    const callText = () => { const args = inputs.map((i) => i.value.trim()); while (args.length && !args[args.length - 1] && args.length > base.filter((a) => !a.startsWith('[')).length) args.pop(); return f.name + '(' + args.join(',') + ')'; };
    const n0 = Math.max(base.length, preset.length, 1);
    for (let i = 0; i < n0; i++) addRow(i, preset[i]);
    function refresh() { if (!inputs.length) return; const v = evalText(callText()); total.textContent = L.t('Formula result = {0}', v == null ? '' : show(v)); result.textContent = '= ' + (v == null ? '' : show(v)); }
    refresh();
    help.textContent = L.t(f.desc);
    ui.dialog({ title: 'Function Arguments', width: 520, body: h('div', { class: 'col' }, h('b', { translate: 'no', text: f.name }), rows, result, h('div', { class: 'fa-desc', text: f.desc }), help, total),
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        const call = callText();
        if (ed.active) {
          /* replace a partly typed name before the caret, then insert the call */
          const t = ed.text(), surf = document.activeElement;
          void surf;
          const m = new RegExp('(?:_xlfn\\.)?' + f.name.replace('.', '\\.') + '\\($', 'i');
          const caret = (G().input.selectionStart != null ? G().input.selectionStart : t.length);
          const before = t.slice(0, caret), after = t.slice(caret);
          if (m.test(before)) ed.setText(before.replace(m, '') + call + after.replace(/^\)/, ''), before.replace(m, '').length + call.length);
          else ed.insertText(call);
          return;
        }
        const s = sh(), sel = G().sel();
        if (G().guardProtect(null, true)) return;
        ed.begin('enter', '=' + (existing && !fresh ? call : call));
        ed.commit();
        void s; void sel;
      } }, { label: 'Cancel' }] });
  };

  /* ================================================================ names */
  const visibleNames = (w) => w.names.filter((n) => !n.hidden && !/^_xlnm\./i.test(n.name));
  const qualified = (w, n) => (n.scope != null && w.sheets[n.scope] ? F.quoteSheet(w.sheets[n.scope].name) + '!' : '') + n.name;
  D.validName = function (name) {
    if (!/^[A-Za-z_\\À-￿][\w.\\À-￿]*$/.test(name)) return 'The name that you entered is not valid.';
    if (/^[A-Za-z]{1,3}\d+$/.test(name) || /^R\d*C\d*$/i.test(name) || /^[rRcC]$/.test(name)) return 'The name that you entered is not valid.\n\nReasons for this can include:\n• The name does not begin with a letter or an underscore\n• The name contains a space or other invalid characters\n• The name conflicts with an Excel built-in name or the name of another object in the workbook';
    if (name.length > 255) return 'The name is too long.';
    return null;
  };
  D.defineName = function () {
    const w = wb(), s = sh();
    let names = w.names.map((n) => Object.assign({}, n));
    const sel = G().sel();
    const left = s.val(sel.r, sel.c - 1), above = s.val(sel.r - 1, sel.c);
    const proposal = typeof left === 'string' ? left : typeof above === 'string' ? above : '';
    const proposedName = proposal.trim().replace(/\s+/g, '_').replace(/[^\w.]/g, '');
    const initialRef = '=' + F.quoteSheet(s.name) + '!' + F.absRangeName(G().range());
    const direct = root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected;
    if (direct) {
      let fields = { name: proposedName, ref: initialRef, comment: '' }, error = '';
      const snapshot = () => ({ names: names.filter((n) => !n.hidden && !/^_xlnm\./i.test(n.name)).map((n) => ({ key: qualified(w, n), name: qualified(w, n), ref: '=' + String(n.ref).replace(/^=/, ''), comment: n.comment || '' })), ...fields, error });
      D.reactDefinedNamesCatalog = () => direct && fields ? snapshot() : null;
      D.reactDefinedNamesAction = (action, value) => {
        if (!fields) return false;
        if (action === 'change') { fields = Object.assign({}, fields, value); error = ''; }
        else if (action === 'select') { const item = names.find((n) => qualified(w, n) === value); if (!item) return false; fields = { name: qualified(w, item), ref: '=' + String(item.ref).replace(/^=/, ''), comment: item.comment || '' }; error = ''; }
        else if (action === 'add' || action === 'commit') {
          let nm = fields.name.trim(), scope = null;
          if (nm) {
            const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(nm);
            if (m) { const sheetName = (m[1] || m[2]).replace(/''/g, "'"); scope = w.sheets.findIndex((item) => item.name.toLowerCase() === sheetName.toLowerCase()); if (scope < 0) error = 'Reference is not valid.'; else nm = m[3]; }
            if (!error) error = D.validName(nm) || '';
            const refText = fields.ref.trim().replace(/^=/, '');
            if (!error) { try { F.parse(refText); } catch (_) { error = 'The formula you typed contains an error.'; } }
            if (!error) {
              const exists = names.some((item) => qualified(w, item) === fields.name.trim());
              if (action === 'add' || !exists) names = names.filter((item) => !(item.name.toLowerCase() === nm.toLowerCase() && (item.scope == null ? null : item.scope) === scope)).concat([{ name: nm, ref: F.toStore(refText), scope, comment: fields.comment || undefined }]);
            }
          } else if (action === 'add') error = 'The name that you entered is not valid.';
          if (error) { L.bus.emit('excelx-defined-names'); return false; }
          if (action === 'commit') { O.tx(w, 'Define Name', () => O.setNames(w, names)); fields = null; }
          else fields = { name: '', ref: initialRef, comment: '' };
        } else if (action === 'delete') { const key = fields.name.trim(); names = names.filter((item) => qualified(w, item) !== key && !(item.scope == null && item.name === key)); fields = Object.assign({}, fields, { name: '' }); error = ''; }
        else if (action === 'close') { fields = null; }
        else return false;
        L.bus.emit('excelx-defined-names'); return true;
      };
      L.bus.emit('excelx-defined-names'); return;
    }
    const nameIn = h('input', { type: 'text', style: 'width:100%' });
    const refIn = D.refBox(F.quoteSheet(s.name) + '!' + F.absRangeName(G().range()), 300);
    const comment = h('input', { type: 'text', style: 'width:100%' });
    const box = h('div');
    const draw = () => { box.textContent = ''; const lb = D.list(names.filter((n) => !n.hidden && !/^_xlnm\./i.test(n.name)).map((n) => [qualified(w, n), qualified(w, n)]), null, (q) => { const n = names.find((x) => qualified(w, x) === q); nameIn.value = q; refIn.input.value = '=' + String(n.ref).replace(/^=/, ''); comment.value = n.comment || ''; }, { class: 'lbox short', translate: 'no' }); box.appendChild(lb); };
    draw();
    /* a name for the active cell's label, as Excel proposes */
    nameIn.value = proposedName;
    const add = () => {
      let nm = nameIn.value.trim();
      if (!nm) return false;
      let scope = null;
      const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(nm);
      if (m) { const sname = (m[1] || m[2]).replace(/''/g, "'"); const idx = w.sheets.findIndex((x) => x.name.toLowerCase() === sname.toLowerCase()); if (idx < 0) { ui.msg('Reference is not valid.', { icon: 'warn' }); return false; } scope = idx; nm = m[3]; }
      const err = D.validName(nm);
      if (err) { ui.msg(err, { icon: 'warn' }); return false; }
      let ref = refIn.input.value.trim().replace(/^=/, '');
      try { F.parse(ref); } catch (e) { ui.msg('The formula you typed contains an error.', { icon: 'warn' }); return false; }
      ref = F.toStore(ref);
      names = names.filter((x) => !(x.name.toLowerCase() === nm.toLowerCase() && (x.scope == null ? null : x.scope) === scope));
      names.push({ name: nm, ref, scope, comment: comment.value || undefined });
      draw();
      return true;
    };
    ui.dialog({ title: 'Define Name', width: 380, body: h('div', { class: 'col' }, h('label', { text: 'Names in &workbook:' }), nameIn, box, ui.field('Co&mment:', comment), ui.field('&Refers to:', refIn)),
      buttons: [{ label: '&Add', onClick: () => { add(); return false; } }, { label: '&Delete', onClick: () => { const q = nameIn.value.trim(); names = names.filter((x) => qualified(w, x) !== q && !(x.scope == null && x.name === q)); nameIn.value = ''; draw(); return false; } },
        { label: 'OK', primary: true, onClick: () => { if (nameIn.value.trim() && !names.some((x) => qualified(w, x) === nameIn.value.trim())) { if (add() === false) return false; } tryRun(() => { O.tx(w, 'Define Name', () => O.setNames(w, names)); }); } }, { label: 'Close' }] });
  };
  D.pasteName = function () {
    const w = wb();
    const names = visibleNames(w);
    if (!names.length) { ui.msg('There are no names defined in this workbook.', { icon: 'info' }); return; }
    let pick = qualified(w, names[0]);
    const direct = root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected;
    if (direct) {
      let state = { names: names.map((n) => ({ key: qualified(w, n), name: qualified(w, n) })), pick };
      D.reactPasteNameCatalog = () => state;
      D.reactPasteNameAction = (action, value) => {
        if (!state) return false;
        if (action === 'select') { if (!state.names.some((n) => n.key === value)) return false; state = Object.assign({}, state, { pick: value }); }
        else if (action === 'insert') { const ed = E(); if (ed.active) ed.insertText(state.pick); else ed.begin('enter', '=' + state.pick); state = null; }
        else if (action === 'paste-list') {
          const s = sh(), sel = G().sel();
          tryRun(() => O.tx(w, 'Paste List', () => names.forEach((n, i) => { O.put(s, sel.r + i, sel.c, { v: qualified(w, n) }); O.put(s, sel.r + i, sel.c + 1, { v: '=' + F.display(String(n.ref).replace(/^=/, '')) }); })));
          state = null;
        } else if (action === 'close') state = null;
        else return false;
        L.bus.emit('excelx-paste-name'); return true;
      };
      L.bus.emit('excelx-paste-name'); return;
    }
    const lb = D.list(names.map((n) => [qualified(w, n), qualified(w, n)]), pick, (v) => { pick = v; }, { translate: 'no' });
    const insert = () => { const ed = E(); if (ed.active) ed.insertText(pick); else { ed.begin('enter', '=' + pick); } };
    const d = ui.dialog({ title: 'Paste Name', width: 300, body: h('div', { class: 'col' }, h('label', { text: 'Paste &name' }), lb),
      buttons: [{ label: 'Paste &List', onClick: () => { tryRun(() => { const s = sh(), sel = G().sel(); O.tx(w, 'Paste List', () => names.forEach((n, i) => { O.put(s, sel.r + i, sel.c, { v: qualified(w, n) }); O.put(s, sel.r + i, sel.c + 1, { v: '=' + F.display(String(n.ref).replace(/^=/, '')) }); })); }); } }, { label: 'OK', primary: true, onClick: () => insert() }, { label: 'Cancel' }] });
    lb.ondbl = () => { d.close(1); insert(); };
  };
  D.createNames = function () {
    const s = sh(), w = wb();
    const rg = O.clip(s, G().range());
    const st = { top: typeof s.val(rg.r1, rg.c1 + (rg.c2 > rg.c1 ? 1 : 0)) === 'string', left: typeof s.val(rg.r1 + (rg.r2 > rg.r1 ? 1 : 0), rg.c1) === 'string', bottom: false, right: false };
    const direct = root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected;
    if (direct) {
      let state = Object.assign({}, st);
      D.reactCreateNamesCatalog = () => state;
      D.reactCreateNamesAction = (action, value) => {
        if (!state) return false;
        if (action === 'change') state = Object.assign({}, state, value);
        else if (action === 'commit') {
          let names = w.names.slice();
          const mk = (label, ref) => { let nm = String(label == null ? '' : label).trim().replace(/\s+/g, '_').replace(/[^\w.\\]/g, ''); if (!nm) return; if (/^\d/.test(nm) || D.validName(nm)) nm = '_' + nm; names = names.filter((x) => !(x.name.toLowerCase() === nm.toLowerCase() && x.scope == null)); names.push({ name: nm, ref: F.quoteSheet(s.name) + '!' + F.absRangeName(ref), scope: null }); };
          const r1 = rg.r1 + (state.top ? 1 : 0), r2 = rg.r2 - (state.bottom ? 1 : 0), c1 = rg.c1 + (state.left ? 1 : 0), c2 = rg.c2 - (state.right ? 1 : 0);
          if (state.top) for (let c = c1; c <= c2; c++) mk(s.val(rg.r1, c), { r1, c1: c, r2, c2: c });
          if (state.bottom) for (let c = c1; c <= c2; c++) mk(s.val(rg.r2, c), { r1, c1: c, r2, c2: c });
          if (state.left) for (let r = r1; r <= r2; r++) mk(s.val(r, rg.c1), { r1: r, c1, r2: r, c2 });
          if (state.right) for (let r = r1; r <= r2; r++) mk(s.val(r, rg.c2), { r1: r, c1, r2: r, c2 });
          O.tx(w, 'Create Names', () => O.setNames(w, names)); state = null;
        } else if (action === 'close') state = null;
        else return false;
        L.bus.emit('excelx-create-names'); return true;
      };
      L.bus.emit('excelx-create-names'); return;
    }
    ui.dialog({ title: 'Create Names', width: 260, body: ui.group('Create names in', ui.check('&Top row', st.top, (v) => { st.top = v; }), ui.check('&Left column', st.left, (v) => { st.left = v; }), ui.check('&Bottom row', false, (v) => { st.bottom = v; }), ui.check('&Right column', false, (v) => { st.right = v; })),
      buttons: [{ label: 'OK', primary: true, onClick: () => tryRun(() => {
        let names = w.names.slice();
        const mk = (label, ref) => { let nm = String(label == null ? '' : label).trim().replace(/\s+/g, '_').replace(/[^\w.\\]/g, ''); if (!nm) return; if (/^\d/.test(nm) || D.validName(nm)) nm = '_' + nm; names = names.filter((x) => !(x.name.toLowerCase() === nm.toLowerCase() && x.scope == null)); names.push({ name: nm, ref: F.quoteSheet(s.name) + '!' + F.absRangeName(ref), scope: null }); };
        const r1 = rg.r1 + (st.top ? 1 : 0), r2 = rg.r2 - (st.bottom ? 1 : 0), c1 = rg.c1 + (st.left ? 1 : 0), c2 = rg.c2 - (st.right ? 1 : 0);
        if (st.top) for (let c = c1; c <= c2; c++) mk(s.val(rg.r1, c), { r1, c1: c, r2, c2: c });
        if (st.bottom) for (let c = c1; c <= c2; c++) mk(s.val(rg.r2, c), { r1, c1: c, r2, c2: c });
        if (st.left) for (let r = r1; r <= r2; r++) mk(s.val(r, rg.c1), { r1: r, c1, r2: r, c2 });
        if (st.right) for (let r = r1; r <= r2; r++) mk(s.val(r, rg.c2), { r1: r, c1, r2: r, c2 });
        O.tx(w, 'Create Names', () => O.setNames(w, names));
      }) }, { label: 'Cancel' }] });
  };
  D.applyNames = function () {
    const s = sh(), w = wb();
    const names = visibleNames(w).filter((n) => { try { const a = F.parse(String(n.ref)); return a.t === 'ref' || a.t === 'area'; } catch (e) { return false; } });
    if (!names.length) { ui.msg('There are no names that refer to cells.', { icon: 'info' }); return; }
    const direct = root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected;
    if (direct) {
      let state = { names: names.map((n) => n.name), selected: names.map((n) => n.name) };
      D.reactApplyNamesCatalog = () => state;
      D.reactApplyNamesAction = (action, name) => {
        if (!state) return false;
        if (action === 'toggle') state = Object.assign({}, state, { selected: state.selected.includes(name) ? state.selected.filter((item) => item !== name) : state.selected.concat(name) });
        else if (action === 'close') state = null;
        else if (action === 'commit') {
          const targets = names.filter((n) => state.selected.includes(n.name)).map((n) => { const a = F.parse(String(n.ref)); return { name: n.name, a, sheet: (a.sheet || s.name).toLowerCase() }; });
          let rg = O.clip(s, G().range());
          if (rg.r1 === rg.r2 && rg.c1 === rg.c2) rg = { r1: 0, c1: 0, r2: s.maxR, c2: s.maxC };
          O.tx(w, 'Apply Names', () => {
            s.each(rg.r1, rg.c1, rg.r2, rg.c2, (cl, r, c) => {
              if (cl.f == null) return;
              let changed = false;
              const ast = F.map(F.parse(cl.f), (n) => {
                if (n.t !== 'ref' && n.t !== 'area') return null;
                const sheet = (n.sheet || s.name).toLowerCase();
                const hit = targets.find((t) => t.sheet === sheet && t.a.t === n.t && (n.t === 'ref' ? t.a.r === n.r && t.a.c === n.c : t.a.r1 === n.r1 && t.a.c1 === n.c1 && t.a.r2 === n.r2 && t.a.c2 === n.c2));
                if (!hit) return null;
                changed = true;
                return { t: 'name', name: hit.name };
              });
              if (changed) O.setFormula(s, r, c, F.toText(ast, { store: true }));
            });
          });
          state = null;
        } else return false;
        L.bus.emit('excelx-apply-names'); return true;
      };
      L.bus.emit('excelx-apply-names'); return;
    }
    const picked = new Set(names.map((n) => n.name));
    const box = h('div', { class: 'lbox checks', translate: 'no' }, ...names.map((n) => ui.check(n.name, true, (v) => { if (v) picked.add(n.name); else picked.delete(n.name); })));
    ui.dialog({ title: 'Apply Names', width: 300, body: h('div', { class: 'col' }, h('label', { text: 'Apply &names:' }), box), buttons: [{ label: 'OK', primary: true, onClick: () => tryRun(() => {
      const targets = names.filter((n) => picked.has(n.name)).map((n) => { const a = F.parse(String(n.ref)); return { name: n.name, a, sheet: (a.sheet || s.name).toLowerCase() }; });
      let rg = O.clip(s, G().range());
      if (rg.r1 === rg.r2 && rg.c1 === rg.c2) rg = { r1: 0, c1: 0, r2: s.maxR, c2: s.maxC };
      O.tx(w, 'Apply Names', () => {
        s.each(rg.r1, rg.c1, rg.r2, rg.c2, (cl, r, c) => {
          if (cl.f == null) return;
          let changed = false;
          const ast = F.map(F.parse(cl.f), (n) => {
            if (n.t !== 'ref' && n.t !== 'area') return null;
            const sheet = (n.sheet || s.name).toLowerCase();
            const hit = targets.find((t) => t.sheet === sheet && t.a.t === n.t && (n.t === 'ref' ? t.a.r === n.r && t.a.c === n.c : t.a.r1 === n.r1 && t.a.c1 === n.c1 && t.a.r2 === n.r2 && t.a.c2 === n.c2));
            if (!hit) return null;
            changed = true;
            return { t: 'name', name: hit.name };
          });
          if (changed) O.setFormula(s, r, c, F.toText(ast, { store: true }));
        });
      });
    }) }, { label: 'Cancel' }] });
  };

  /* ================================================================ Evaluate Formula */
  D.evaluateFormula = function () {
    const s = sh(), w = wb(), sel = G().sel();
    const cell = s.get(sel.r, sel.c);
    if (!cell || cell.f == null) { ui.msg('The active cell does not contain a formula.', { icon: 'info' }); return; }
    const ctx = C.ctx(w, s, sel.r, sel.c, !!cell.af);
    let ast = F.parse(cell.f);
    const lit = (v) => {
      if (v instanceof C.Ref || v instanceof C.RefList) return null;
      if (v && v.rows && Array.isArray(v.rows)) return { t: 'arr', rows: v.rows.slice(0, 20).map((r) => r.slice(0, 20).map((x) => lit(x) || { t: 'str', v: '' })) };
      if (typeof v === 'number') return { t: 'num', v };
      if (typeof v === 'string') return { t: 'str', v };
      if (typeof v === 'boolean') return { t: 'bool', v };
      if (M.isErr(v)) return { t: 'err', v: v.e };
      if (v == null) return { t: 'num', v: 0 };
      return null;
    };
    const isLit = (n) => n && /^(num|str|bool|err|arr|missing)$/.test(n.t);
    /** the next node to evaluate: a single-cell reference, or an operator / call whose operands are values */
    const next = (n) => {
      if (!n) return null;
      if (n.t === 'ref' || n.t === 'name' || n.t === 'struct') return n;
      if (n.t === 'area') return null;
      const kidsOf = n.t === 'fn' ? n.args : n.t === 'op' ? [n.a, n.b] : n.t === 'un' || n.t === 'paren' || n.t === 'pct' ? [n.a] : [];
      for (const k of kidsOf) { const x = next(k); if (x) return x; }
      if (n.t === 'fn' && n.args.every((a) => isLit(a) || a.t === 'area' || a.t === 'ref')) return n;
      if ((n.t === 'op' || n.t === 'un' || n.t === 'pct' || n.t === 'paren') && kidsOf.every((k) => isLit(k) || k.t === 'area')) return n;
      return null;
    };
    const text = h('div', { class: 'ev-text', translate: 'no' });
    const direct = root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected;
    let reactEvaluation = null, reactEvaluationAction = null;
    D.reactEvaluateCatalog = () => reactEvaluation;
    D.reactEvaluateAction = (action) => reactEvaluationAction ? reactEvaluationAction(action) : false;
    const render = (mark) => {
      if (direct) {
        const full = F.toText(ast), part = mark ? F.toText(mark) : '', index = mark ? full.indexOf(part) : -1;
        reactEvaluation = { reference: F.quoteSheet(s.name) + '!' + F.cellName(sel.r, sel.c, true, true), expression: '=' + full, start: index >= 0 ? index + 1 : -1, end: index >= 0 ? index + 1 + part.length : -1, complete: !mark };
        L.bus.emit('excelx-evaluate-formula-dialog'); return;
      }
      text.textContent = '';
      const full = F.toText(ast);
      if (!mark) { text.textContent = '=' + full; return; }
      const part = F.toText(mark);
      const i = full.indexOf(part);
      text.append(document.createTextNode('=' + full.slice(0, Math.max(0, i))), h('u', { translate: 'no', text: part }), document.createTextNode(i >= 0 ? full.slice(i + part.length) : ''));
    };
    let cur = next(ast);
    if (direct) {
      reactEvaluationAction = (action) => {
        if (action === 'close') { reactEvaluation = reactEvaluationAction = null; L.bus.emit('excelx-evaluate-formula-dialog'); return true; }
        if (action !== 'evaluate' && action !== 'restart') return false;
        if (!cur) { ast = F.parse(cell.f); cur = next(ast); render(cur); return true; }
        let value;
        try { value = C.ev(cur, ctx); } catch (_) { value = M.ERR.VALUE; }
        if (value instanceof C.Ref && value.single) value = C.cellValue(value.sheet, value.r1, value.c1);
        const replacement = lit(value) || { t: 'str', v: show(value) }, target = cur;
        ast = target === ast ? replacement : F.map(ast, (node) => node === target ? replacement : null);
        cur = isLit(ast) ? null : next(ast);
        render(cur); return true;
      };
      render(cur); return;
    }
    render(cur);
    const evalBtn = { label: '&Evaluate', primary: true, onClick: () => {
      if (!cur) { ast = F.parse(cell.f); cur = next(ast); render(cur); return false; }
      let v;
      try { v = C.ev(cur, ctx); } catch (e) { v = M.ERR.VALUE; }
      if (v instanceof C.Ref && v.single) v = C.cellValue(v.sheet, v.r1, v.c1);
      const rep = lit(v) || { t: 'str', v: show(v) };
      const target = cur;
      ast = target === ast ? rep : F.map(ast, (n) => (n === target ? rep : null));
      cur = isLit(ast) ? null : next(ast);
      render(cur);
      if (!cur) d.buttons[0].textContent = L.t('Restart');
      return false;
    } };
    const d = ui.dialog({ title: 'Evaluate Formula', width: 520, body: h('div', { class: 'col' }, h('div', { class: 'row', style: 'gap:16px' }, h('div', null, h('div', { class: 'hint', text: 'Reference:' }), h('b', { translate: 'no', text: F.quoteSheet(s.name) + '!' + F.cellName(sel.r, sel.c, true, true) })), h('div', { style: 'flex:1' }, h('div', { class: 'hint', text: 'Evaluation:' }), text)), h('div', { class: 'hint', text: 'To show the result of the underlined expression, click Evaluate. The most recent result appears italicized.' })),
      buttons: [evalBtn, { label: 'Close' }] });
  };

  /* ================================================================ Watch Window */
  let watchDlg = null;
  let reactWatch = null;
  const watchSnapshot = () => {
    if (!reactWatch) return null;
    const { workbook } = reactWatch, book = A().books.find((item) => item.wb === workbook);
    const items = (workbook._watches || []).map((watch, index) => {
      if (!workbook.sheets.includes(watch.sh)) return null;
      const cell = watch.sh.get(watch.r, watch.c);
      return { index, book: book ? book.name : workbook.name, sheet: watch.sh.name, name: '', cell: F.cellName(watch.r, watch.c, true, true), value: L.csv.cellText(watch.sh, watch.r, watch.c), formula: cell && cell.f != null ? '=' + F.display(cell.f) : '' };
    }).filter(Boolean);
    return { items, selected: reactWatch.selected };
  };
  D.reactWatchCatalog = () => watchSnapshot();
  D.reactWatchAction = (action, index) => {
    if (!reactWatch) return false;
    if (action === 'close') { reactWatch = null; L.bus.emit('excelx-watch-window'); return true; }
    if (action === 'add') { D.watchWindow(true); return true; }
    const watches = reactWatch.workbook._watches || [];
    if (action === 'delete') { if (!Number.isInteger(index) || index < 0 || index >= watches.length) return false; watches.splice(index, 1); reactWatch.selected = -1; }
    else if (action === 'select') { if (!Number.isInteger(index) || index < 0 || index >= watches.length) return false; reactWatch.selected = index; }
    else if (action === 'activate') { const watch = watches[index]; if (!watch || !reactWatch.workbook.sheets.includes(watch.sh)) return false; const sheetIndex = reactWatch.workbook.sheets.indexOf(watch.sh); if (sheetIndex !== reactWatch.workbook.active) A().activateSheet(sheetIndex); G().select(watch.r, watch.c); return true; }
    else return false;
    L.bus.emit('excelx-watch-window'); return true;
  };
  D.watchWindow = function (addCurrent) {
    const w = wb();
    if (!w._watches) w._watches = [];
    if (addCurrent) { const s = sh(); for (const rg of G().ranges()) { const r = O.clip(s, rg); for (let rr = r.r1; rr <= Math.min(r.r2, r.r1 + 50); rr++) for (let cc = r.c1; cc <= Math.min(r.c2, r.c1 + 20); cc++) if (!w._watches.some((x) => x.sh === s && x.r === rr && x.c === cc)) w._watches.push({ sh: s, r: rr, c: cc }); } }
    if (root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected) {
      if (!reactWatch || reactWatch.workbook !== w) reactWatch = { workbook: w, selected: -1 };
      L.bus.emit('excelx-watch-window'); return;
    }
    if (watchDlg && document.contains(watchDlg.el)) { watchDlg.draw(); return; }
    const tbl = h('table', { class: 'fr-table' });
    let sel = -1;
    const draw = () => {
      tbl.textContent = '';
      tbl.appendChild(h('tr', null, ...[L.N('Book|workbook'), L.N('Sheet'), L.N('Name'), L.N('Cell'), L.N('Value'), L.N('Formula')].map((t) => h('th', { text: t }))));
      w._watches.forEach((x, i) => {
        if (!w.sheets.includes(x.sh)) return;
        const cl = x.sh.get(x.r, x.c);
        const tr = h('tr', { class: i === sel ? 'on' : '' }, h('td', { translate: 'no', text: A().books[A().cur].name }), h('td', { translate: 'no', text: x.sh.name }), h('td', { translate: 'no', text: '' }), h('td', { translate: 'no', text: F.cellName(x.r, x.c, true, true) }), h('td', { translate: 'no', text: L.csv.cellText(x.sh, x.r, x.c) }), h('td', { translate: 'no', text: cl && cl.f != null ? '=' + F.display(cl.f) : '' }));
        tr.addEventListener('click', () => { sel = i; draw(); });
        tr.addEventListener('dblclick', () => { const k = w.sheets.indexOf(x.sh); if (k !== w.active) A().activateSheet(k); G().select(x.r, x.c); });
        tbl.appendChild(tr);
      });
    };
    draw();
    const off = L.bus.on ? L.bus.on('changed', () => { if (watchDlg && document.contains(watchDlg.el)) draw(); }) : null;
    void off;
    const d = ui.dialog({ title: 'Watch Window', width: 520, modeless: true, body: h('div', { class: 'col' }, h('div', { class: 'row', style: 'gap:6px' }, ui.button('&Add Watch...', () => { D.watchWindow(true); }), ui.button('&Delete Watch', () => { if (sel >= 0) { w._watches.splice(sel, 1); sel = -1; draw(); } })), tbl), buttons: [{ label: 'Close', primary: true }] });
    watchDlg = { el: d.el, draw };
  };

  /* ================================================================ Error Checking */
  const ERRDESC = { '#DIV/0!': 'Divide by Zero Error', '#N/A': 'Value Not Available Error', '#NAME?': 'Invalid Name Error', '#NULL!': 'Null Error', '#NUM!': 'Number Error', '#REF!': 'Invalid Cell Reference Error', '#VALUE!': 'Error in Value', '#SPILL!': 'Spill Range Isn\'t Blank', '#CALC!': 'Calculation Error' };
  const ERRHELP = { '#DIV/0!': 'The formula or function used is dividing by zero or empty cells.', '#N/A': 'A value is not available to the formula or function.', '#NAME?': 'The formula contains unrecognized text.', '#NULL!': 'The intersection of two areas is empty.', '#NUM!': 'A formula or function contains invalid numeric values.', '#REF!': 'A cell reference is not valid.', '#VALUE!': 'A value used in the formula is of the wrong data type.' };
  let reactErrorCheck = null;
  const errorSnapshot = () => {
    if (!reactErrorCheck) return null;
    const { sheet, list, index } = reactErrorCheck, issue = list[index], cell = sheet.get(issue.r, issue.c);
    return { reference: F.cellName(issue.r, issue.c), formula: cell && cell.f != null ? '=' + F.display(cell.f) : String(cell ? cell.v : ''), title: issue.kind === 'numtext' ? L.t('Number Stored as Text') : L.t(ERRDESC[issue.kind] || L.N('Error')), description: issue.kind === 'numtext' ? L.t('The number in this cell is formatted as text or preceded by an apostrophe.') : L.t(ERRHELP[issue.kind] || ''), action: issue.kind === 'numtext' ? L.t('Convert to Number') : L.t('Show Calculation Steps...'), index, count: list.length, canConvert: issue.kind === 'numtext' };
  };
  D.reactErrorCheckCatalog = () => errorSnapshot();
  D.reactErrorCheckAction = (action) => {
    if (!reactErrorCheck) return false;
    const state = reactErrorCheck, issue = state.list[state.index], { sheet, workbook, list } = state;
    const finish = () => { reactErrorCheck = null; L.bus.emit('excelx-error-checking-dialog'); };
    if (action === 'close') { finish(); return true; }
    if (action === 'previous' || action === 'next') { state.index = (state.index + (action === 'next' ? 1 : list.length - 1)) % list.length; G().select(list[state.index].r, list[state.index].c); L.bus.emit('excelx-error-checking-dialog'); return true; }
    if (action === 'help') { L.panes.helpQuery = 'formulas'; finish(); L.panes.task.show('help'); return true; }
    if (action === 'edit') { finish(); E().begin('edit'); return true; }
    if (action === 'fix' && issue.kind !== 'numtext') { D.evaluateFormula(); return true; }
    if (action !== 'fix' && action !== 'ignore') return false;
    if (action === 'ignore') (workbook._ignoredErr || (workbook._ignoredErr = new Set())).add(sheet.id + ':' + issue.r + ':' + issue.c);
    else tryRun(() => O.tx(workbook, 'Convert to Number', () => O.enter(sheet, issue.r, issue.c, String(sheet.val(issue.r, issue.c)).trim())));
    list.splice(state.index, 1);
    if (!list.length) { finish(); ui.msg('The error check of the entire sheet is complete.', { icon: 'info' }); return true; }
    state.index = Math.min(state.index, list.length - 1); G().select(list[state.index].r, list[state.index].c); L.bus.emit('excelx-error-checking-dialog'); return true;
  };
  D.errorChecking = function () {
    const s = sh(), w = wb();
    const ignored = (w._ignoredErr || (w._ignoredErr = new Set()));
    const list = [];
    s.rows.forEach((row, r) => { if (row) row.cells.forEach((cl, c) => {
      if (!cl) return;
      const k = s.id + ':' + r + ':' + c;
      if (ignored.has(k)) return;
      const v = cl.dirty ? C.cellValue(s, r, c) : cl.v;
      if (cl.f != null && M.isErr(v)) list.push({ r, c, kind: v.e });
      else if (cl.f == null && typeof v === 'string' && /^\s*-?[\d,]+(\.\d+)?\s*$/.test(v) && !(LY.styleOf(s, r, c, cl).nf === '@')) list.push({ r, c, kind: 'numtext' });
    }); });
    const sel = G().sel();
    let i = Math.max(0, list.findIndex((x) => x.r > sel.r || (x.r === sel.r && x.c >= sel.c)));
    if (!list.length) { ui.msg('The error check of the entire sheet is complete.', { icon: 'info' }); return; }
    if (root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected) {
      reactErrorCheck = { sheet: s, workbook: w, list, index: i };
      G().select(list[i].r, list[i].c); L.bus.emit('excelx-error-checking-dialog'); return;
    }
    const head = h('b'), formula = h('div', { class: 'ev-text', translate: 'no' }), title = h('div', { style: 'font-weight:bold' }), desc = h('div');
    const showI = () => {
      const x = list[i];
      G().select(x.r, x.c);
      const cl = s.get(x.r, x.c);
      head.textContent = L.t('Error in cell {0}:', F.cellName(x.r, x.c));
      formula.textContent = cl && cl.f != null ? '=' + F.display(cl.f) : String(cl ? cl.v : '');
      title.textContent = x.kind === 'numtext' ? L.t('Number Stored as Text') : L.t(ERRDESC[x.kind] || L.N('Error'));
      desc.textContent = x.kind === 'numtext' ? L.t('The number in this cell is formatted as text or preceded by an apostrophe.') : L.t(ERRHELP[x.kind] || '');
      fixBtn.textContent = x.kind === 'numtext' ? L.t('Convert to Number') : L.t('Show Calculation Steps...');
    };
    const fixBtn = ui.button('Show Calculation Steps...', () => {
      const x = list[i];
      if (x.kind === 'numtext') { tryRun(() => O.tx(w, 'Convert to Number', () => O.enter(s, x.r, x.c, String(s.val(x.r, x.c)).trim()))); list.splice(i, 1); if (!list.length) { d.close(0); ui.msg('The error check of the entire sheet is complete.', { icon: 'info' }); return; } i = Math.min(i, list.length - 1); showI(); }
      else D.evaluateFormula();
    }, { style: 'width:100%' });
    const btn = (label, fn) => ui.button(label, fn, { style: 'width:100%' });
    const d = ui.dialog({ title: 'Error Checking', width: 480, modeless: true, body: h('div', { class: 'row', style: 'gap:12px;align-items:flex-start' }, h('div', { class: 'col', style: 'flex:1' }, head, formula, title, desc),
      h('div', { class: 'col', style: 'width:170px' }, btn('Help on this error', () => { L.panes.helpQuery = 'formulas'; L.panes.task.show('help'); }), fixBtn, btn('Ignore Error', () => { const x = list[i]; ignored.add(s.id + ':' + x.r + ':' + x.c); list.splice(i, 1); if (!list.length) { d.close(0); ui.msg('The error check of the entire sheet is complete.', { icon: 'info' }); return; } i = Math.min(i, list.length - 1); showI(); }), btn('Edit in Formula Bar', () => { d.close(0); E().begin('edit'); }))),
      buttons: [{ label: '&Previous', onClick: () => { i = (i - 1 + list.length) % list.length; showI(); return false; } }, { label: '&Next', primary: true, onClick: () => { i = (i + 1) % list.length; showI(); return false; } }, { label: 'Close' }] });
    showI();
  };
  D.spelling = () => L.spell.check();

  /* ================================================================ Page Setup */
  const HF_PRESETS = ['', '&P', 'Page &P', 'Page &P of &N', '&A', '&F', '&A, Page &P', '&F, Page &P', '&D', 'Confidential, &D, Page &P', '&F, &A', 'Prepared by &[User] &D, Page &P'];
  const hfLabel = (code, s) => { if (!code) return L.t('(none)'); const p = L.print.parseHF(code.replace('&[User]', A().opts.userName || ''), { page: 1, pages: 1, sheet: s.name, file: A().books[A().cur].name }); return ['l', 'c', 'r'].map((k) => p[k].map((x) => x.text).join('')).filter(Boolean).join(', '); };
  D.pageSetup = function (tab) {
    const s = sh(), w = wb();
    const p = JSON.parse(JSON.stringify(s.print));
    p.margins = Object.assign({ l: 0.75, r: 0.75, t: 1, b: 1, header: 0.5, footer: 0.5 }, p.margins);
    const idx = w.sheets.indexOf(s);
    const getName = (key) => { const n = w.names.find((x) => x.name.toLowerCase() === key && x.scope === idx); return n ? String(n.ref).replace(/^=/, '').replace(new RegExp('(?:' + F.quoteSheet(s.name).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')!', 'g'), '') : ''; };
    let area = getName('_xlnm.print_area');
    const titles = L.print.titles(s);
    const direct = root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected;
    if (direct) {
      let state = { p, area, rows: titles.rows ? '$' + (titles.rows.r1 + 1) + ':$' + (titles.rows.r2 + 1) : '', cols: titles.cols ? '$' + F.colName(titles.cols.c1) + ':$' + F.colName(titles.cols.c2) : '', first: p.firstPage != null ? String(p.firstPage) : L.t('Auto'), tab: Math.max(0, ['Page', 'Margins', 'Header/Footer', 'Sheet'].indexOf(tab)), error: '', resolve: null };
      const snapshot = () => state && ({ p: Object.assign({}, state.p, { margins: Object.assign({}, state.p.margins) }), area: state.area, rows: state.rows, cols: state.cols, first: state.first, tab: state.tab, error: state.error, headers: { header: hfLabel(state.p.header, s), footer: hfLabel(state.p.footer, s) }, headerOptions: HF_PRESETS.map((code) => ({ code, label: hfLabel(code, s) })), papers: Object.keys(L.print.PAPERS).map((key) => ({ key: +key, label: L.t('{0} ({1} x {2} in)', L.t(L.print.PAPERS[key][0]), L.print.PAPERS[key][1], L.print.PAPERS[key][2]) })) });
      D.reactPageSetupCatalog = () => snapshot();
      D.reactPageSetupAction = (action, key, value) => {
        if (!state) return false;
        if (action === 'change') {
          if (key === 'area' || key === 'rows' || key === 'cols' || key === 'first') state[key] = value;
          else if (key === 'tab') state.tab = Number(value);
          else if (key.startsWith('margin.')) state.p.margins[key.slice(7)] = Number(value) || 0;
          else if (key === 'print.fit') state.p.fit = value === 'true';
          else if (key.startsWith('print.')) state.p[key.slice(6)] = value;
          else state.p[key] = value;
          state.error = '';
        } else if (action === 'custom') {
          const headerKey = key === 'header' ? 'header' : 'footer';
          D.customHF(state.p[headerKey] || '', headerKey === 'header' ? L.N('Header') : L.N('Footer'), (code) => { if (!state) return; state.p[headerKey] = code; L.bus.emit('excelx-page-setup'); });
          return true;
        } else if (action === 'close') { const resolve = state.resolve; state = null; if (resolve) resolve(0); }
        else if (action === 'apply' || action === 'print' || action === 'preview') {
          state.p.firstPage = /^\d+$/.test(String(state.first).trim()) ? +String(state.first).trim() : null;
          const names = w.names.filter((n) => !(n.scope === idx && /^_xlnm\.(print_area|print_titles)$/i.test(n.name)));
          const q = F.quoteSheet(s.name), a = state.area.trim().replace(/^=/, '');
          if (a) { const parts = a.split(',').map((x) => x.trim()).filter(Boolean).map((x) => (/!/.test(x) ? x : q + '!' + x)); names.push({ name: '_xlnm.Print_Area', ref: parts.join(','), scope: idx }); }
          const tparts = [], rr = state.rows.trim().replace(/^=/, ''), cc = state.cols.trim().replace(/^=/, '');
          if (cc) tparts.push(/!/.test(cc) ? cc : q + '!' + cc);
          if (rr) tparts.push(/!/.test(rr) ? rr : q + '!' + rr);
          if (tparts.length) names.push({ name: '_xlnm.Print_Titles', ref: tparts.join(','), scope: idx });
          try { for (const n of names) if (/^_xlnm/.test(n.name)) F.parse(n.ref); } catch (_) { state.error = L.t('Reference is not valid.'); L.bus.emit('excelx-page-setup'); return false; }
          tryRun(() => O.tx(w, 'Page Setup', () => { O.setPrint(s, state.p); O.setNames(w, names); }));
          L.print.computeBreaks(s);
          if (action !== 'apply') setTimeout(() => action === 'print' ? D.print() : L.print.preview(), 0);
          const resolve = state.resolve; state = null; if (resolve) resolve(1);
        } else return false;
        L.bus.emit('excelx-page-setup'); return true;
      };
      return new Promise((resolve) => { state.resolve = resolve; L.bus.emit('excelx-page-setup'); });
    }
    /* page */
    const adj = ui.spin({ value: p.scale || 100, min: 10, max: 400, step: 5, dec: 0, fmt: (v) => L.t('{0}% normal size', Math.round(v)), onChange: (v) => { p.scale = Math.round(v); } });
    const fitW = ui.spin({ value: p.fitW == null ? 1 : p.fitW, min: 0, max: 32767, step: 1, dec: 0, onChange: (v) => { p.fitW = Math.round(v) || null; } });
    const fitH = ui.spin({ value: p.fitH == null ? 1 : p.fitH, min: 0, max: 32767, step: 1, dec: 0, onChange: (v) => { p.fitH = Math.round(v) || null; } });
    const first = h('input', { type: 'text', value: p.firstPage != null ? String(p.firstPage) : L.t('Auto'), style: 'width:70px' });
    const pageTab = h('div', { class: 'col' },
      ui.group('Orientation', h('div', { class: 'row', style: 'gap:24px' }, ui.radio('ps-o', 'Por&trait', p.orientation !== 'landscape', () => { p.orientation = 'portrait'; }), ui.radio('ps-o', '&Landscape', p.orientation === 'landscape', () => { p.orientation = 'landscape'; }))),
      ui.group('Scaling', h('div', { class: 'row', style: 'gap:8px' }, ui.radio('ps-s', '&Adjust to:', !p.fit, () => { p.fit = false; }), adj), h('div', { class: 'row', style: 'gap:8px;flex-wrap:wrap' }, ui.radio('ps-s', '&Fit to:', !!p.fit, () => { p.fit = true; }), fitW, h('span', { text: 'page(s) wide by' }), fitH, h('span', { text: 'tall' }))),
      ui.field('Paper si&ze:', ui.select(Object.keys(L.print.PAPERS).map((k) => [k, L.t('{0} ({1} x {2} in)', L.t(L.print.PAPERS[k][0]), L.print.PAPERS[k][1], L.print.PAPERS[k][2])]), p.paper || 1, (v) => { p.paper = +v; })),
      ui.field('Print &quality:', ui.select(['600 dpi', '300 dpi', '150 dpi'], '600 dpi', null, { translate: 'no' })),
      ui.field('Fi&rst page number:', first));
    /* margins */
    const mspin = (k) => ui.spin({ value: p.margins[k], min: 0, max: 20, step: 0.25, unit: '', onChange: (v) => { p.margins[k] = v; } });
    const marTab = h('div', { class: 'col' }, h('div', { class: 'ps-margins' }, ui.field('&Top:', mspin('t')), ui.field('H&eader:', mspin('header')), ui.field('&Left:', mspin('l')), ui.field('&Right:', mspin('r')), ui.field('&Bottom:', mspin('b')), ui.field('&Footer:', mspin('footer'))),
      ui.group('Center on page', ui.check('Hori&zontally', !!p.hCenter, (v) => { p.hCenter = v; }), ui.check('&Vertically', !!p.vCenter, (v) => { p.vCenter = v; })), h('div', { class: 'hint', text: 'Margins are in inches.' }));
    /* header / footer */
    /* the headers and footers are the sheet's own text: shown as they print */
    const hfSel = (key) => { const opts = HF_PRESETS.slice(); if (p[key] && !opts.includes(p[key])) opts.unshift(p[key]); const sel = ui.select(opts.map((c) => [c, hfLabel(c, s)]), p[key] || '', (v) => { p[key] = v; prevH.textContent = hfLabel(p.header, s); prevF.textContent = hfLabel(p.footer, s); }, { translate: 'no' }); return sel; };
    const prevH = h('div', { class: 'hf-prev', translate: 'no', text: hfLabel(p.header, s) }), prevF = h('div', { class: 'hf-prev', translate: 'no', text: hfLabel(p.footer, s) });
    const hSel = hfSel('header'), fSel = hfSel('footer');
    const custom = (key) => D.customHF(p[key] || '', key === 'header' ? L.N('Header') : L.N('Footer'), (code) => { p[key] = code; const sel = key === 'header' ? hSel : fSel; if (!Array.from(sel.options).some((o) => o.value === code)) sel.insertBefore(h('option', { value: code, translate: 'no', text: hfLabel(code, s) }), sel.firstChild); sel.value = code; prevH.textContent = hfLabel(p.header, s); prevF.textContent = hfLabel(p.footer, s); });
    const hfTab = h('div', { class: 'col' }, prevH, ui.field('He&ader:', hSel), h('div', { class: 'row', style: 'gap:8px' }, ui.button('&Custom Header...', () => custom('header')), ui.button('C&ustom Footer...', () => custom('footer'))), ui.field('&Footer:', fSel), prevF,
      ui.check('Different &odd and even pages', !!p.diffOddEven, (v) => { p.diffOddEven = v; }), ui.check('Different first &page', !!p.diffFirst, (v) => { p.diffFirst = v; }), ui.check('Scale with document', p.hfScale !== false, (v) => { p.hfScale = v; }), ui.check('Align with page margins', p.hfAlign !== false, (v) => { p.hfAlign = v; }));
    /* sheet */
    const areaIn = D.refBox(area, 220), rowsIn = D.refBox(titles.rows ? '$' + (titles.rows.r1 + 1) + ':$' + (titles.rows.r2 + 1) : '', 160), colsIn = D.refBox(titles.cols ? '$' + F.colName(titles.cols.c1) + ':$' + F.colName(titles.cols.c2) : '', 160);
    const sheetTab = h('div', { class: 'col' }, ui.field('Print &area:', areaIn), ui.group('Print titles', ui.field('&Rows to repeat at top:', rowsIn), ui.field('&Columns to repeat at left:', colsIn)),
      ui.group('Print', h('div', { class: 'ps-print' }, ui.check('&Gridlines', !!p.gridLines, (v) => { p.gridLines = v; }), ui.check('&Black and white', !!p.bw, (v) => { p.bw = v; }), ui.check('Draft &quality', !!p.draft, (v) => { p.draft = v; }), ui.check('Row and colu&mn headings', !!p.headings, (v) => { p.headings = v; }),
        ui.field('Co&mments:', ui.select([['none', '(None)'], ['atEnd', 'At end of sheet'], ['asDisplayed', 'As displayed on sheet']], p.comments || 'none', (v) => { p.comments = v; })), ui.field('Cell &errors as:', ui.select([['displayed', 'displayed'], ['blank', '<blank>'], ['dash', '--'], ['NA', '#N/A']], p.errors || 'displayed', (v) => { p.errors = v; })))),
      ui.group('Page order', ui.radio('ps-po', '&Down, then over', p.pageOrder !== 'overThenDown', () => { p.pageOrder = 'downThenOver'; }), ui.radio('ps-po', 'O&ver, then down', p.pageOrder === 'overThenDown', () => { p.pageOrder = 'overThenDown'; })));
    const apply = () => {
      p.firstPage = /^\d+$/.test(first.value.trim()) ? +first.value.trim() : null;
      const names = w.names.filter((n) => !(n.scope === idx && /^_xlnm\.(print_area|print_titles)$/i.test(n.name)));
      const q = F.quoteSheet(s.name);
      const a = areaIn.input.value.trim().replace(/^=/, '');
      if (a) { const parts = a.split(',').map((x) => x.trim()).filter(Boolean).map((x) => (/!/.test(x) ? x : q + '!' + x)); names.push({ name: '_xlnm.Print_Area', ref: parts.join(','), scope: idx }); }
      const tparts = [];
      const rr = rowsIn.input.value.trim().replace(/^=/, ''), cc = colsIn.input.value.trim().replace(/^=/, '');
      if (cc) tparts.push(/!/.test(cc) ? cc : q + '!' + cc);
      if (rr) tparts.push(/!/.test(rr) ? rr : q + '!' + rr);
      if (tparts.length) names.push({ name: '_xlnm.Print_Titles', ref: tparts.join(','), scope: idx });
      try { for (const n of names) if (/^_xlnm/.test(n.name)) F.parse(n.ref); } catch (e) { ui.msg('Reference is not valid.', { icon: 'warn' }); return false; }
      tryRun(() => O.tx(w, 'Page Setup', () => { O.setPrint(s, p); O.setNames(w, names); }));
      L.print.computeBreaks(s);
      return true;
    };
    return new Promise((res) => {
      ui.dialog({ title: 'Page Setup', width: 470, body: ui.tabs([{ label: 'Page', body: pageTab }, { label: 'Margins', body: marTab }, { label: 'Header/Footer', body: hfTab }, { label: 'Sheet', body: sheetTab }], tab || 0),
        buttons: [{ label: '&Print...', onClick: () => { if (apply() === false) return false; setTimeout(() => D.print(), 0); } }, { label: 'Print Previe&w', onClick: () => { if (apply() === false) return false; setTimeout(() => L.print.preview(), 0); } }, { label: 'OK', primary: true, onClick: () => apply() }, { label: 'Cancel' }] }).done.then(res);
    });
  };
  /** Custom Header / Footer: three sections with the field buttons */
  D.customHF = function (code, kind, done) {
    const secs = { l: '', c: '', r: '' };
    /* split the code into its &L / &C / &R parts, keeping the codes */
    let cur = 'c';
    String(code).split(/(&[LCR])/).forEach((part) => { if (/^&[LCR]$/.test(part)) cur = part[1].toLowerCase(); else secs[cur] += part; });
    if (root.VO && root.VO.reactDirectHost && root.VO.reactDirectHost.root && root.VO.reactDirectHost.root.isConnected) {
      reactCustomHF = { kind, sections: secs };
      reactCustomHFDone = done;
      L.bus.emit('excelx-custom-hf-dialog');
      return;
    }
    const areas = { l: h('textarea', { rows: '5' }), c: h('textarea', { rows: '5' }), r: h('textarea', { rows: '5' }) };
    for (const k in areas) areas[k].value = secs[k];
    let last = areas.c;
    for (const k in areas) areas[k].addEventListener('focus', () => { last = areas[k]; });
    const ins = (t) => { const a = last; const s0 = a.selectionStart, s1 = a.selectionEnd; a.value = a.value.slice(0, s0) + t + a.value.slice(s1); a.focus(); a.setSelectionRange(s0 + t.length, s0 + t.length); };
    const fb = (label, t) => h('button', { type: 'button', class: 'btn small', text: label, onclick: () => ins(t) });
    const font = () => { const f = h('input', { type: 'text', value: 'Arial', style: 'width:110px' }); const st = ui.select([['Regular', L.N('Regular')], ['Bold', L.N('Bold')], ['Italic', L.N('Italic')], ['Bold Italic', 'Bold Italic']], 'Regular'); const sz = h('input', { type: 'text', value: '10', style: 'width:40px' }); ui.dialog({ title: 'Font', width: 320, body: h('div', { class: 'row', style: 'gap:8px' }, ui.field('Font:', f), ui.field('Style:', st), ui.field('Size:', sz)), buttons: [{ label: 'OK', primary: true, onClick: () => ins(`&"${f.value},${st.value}"&${parseInt(sz.value, 10) || 10}`) }, { label: 'Cancel' }] }); };
    ui.dialog({ title: kind, width: 560, body: h('div', { class: 'col' }, h('div', { class: 'hint', text: 'To format text: select the text, then choose the font button. To insert a page number, date, time, file path, filename, or tab name: position the insertion point in the edit box, then choose the appropriate button.' }),
      h('div', { class: 'row', style: 'gap:4px;flex-wrap:wrap' }, h('button', { type: 'button', class: 'btn small', text: 'A (Font)', onclick: font }), fb('Page #', '&P'), fb(L.N('Pages'), '&N'), fb(L.N('Date'), '&D'), fb(L.N('Time'), '&T'), fb('Path & File', '&Z&F'), fb(L.N('File'), '&F'), fb(L.N('Tab'), '&A')),
      h('div', { class: 'hf-secs' }, ui.field('&Left section:', areas.l), ui.field('&Center section:', areas.c), ui.field('&Right section:', areas.r))),
    buttons: [{ label: 'OK', primary: true, onClick: () => { const out = (areas.l.value ? '&L' + areas.l.value : '') + (areas.c.value ? '&C' + areas.c.value : '') + (areas.r.value ? '&R' + areas.r.value : ''); done(out); } }, { label: 'Cancel' }] });
  };

  /* ================================================================ Print */
  D.print = function () {
    const st = { what: 'sheets', range: 'all', from: 1, to: 1, copies: 1 };
    const from = h('input', { type: 'text', value: '1', style: 'width:40px' }), to = h('input', { type: 'text', value: '1', style: 'width:40px' });
    const items = L.print.allPages(A().selectedSheets(), null);
    to.value = String(Math.max(1, items.length));
    return new Promise((res) => {
      ui.dialog({ title: 'Print', width: 440, body: h('div', { class: 'col' },
        ui.group('Printer', ui.field('&Name:', ui.select([['pdf', 'Excelx PDF Writer']], 'pdf')), h('div', { class: 'hint', text: 'Excelx prints by creating a PDF that you can open and print on any printer.' })),
        h('div', { class: 'row', style: 'gap:10px;align-items:stretch' },
          ui.group('Print range', ui.radio('pr-r', '&All', true, () => { st.range = 'all'; }), h('div', { class: 'row', style: 'gap:6px' }, ui.radio('pr-r', 'Pa&ge(s)', false, () => { st.range = 'pages'; }), h('span', { text: 'From:' }), from, h('span', { text: 'To:' }), to)),
          ui.group('Print what', ui.radio('pr-w', 'Selectio&n', false, () => { st.what = 'selection'; }), ui.radio('pr-w', 'Active s&heet(s)', true, () => { st.what = 'sheets'; }), ui.radio('pr-w', '&Entire workbook', false, () => { st.what = 'workbook'; }))),
        ui.field('Number of &copies:', ui.spin({ value: 1, min: 1, max: 99, step: 1, dec: 0 }))),
      buttons: [{ label: 'Pre&view', onClick: () => { setTimeout(() => L.print.preview(), 0); } }, { label: 'OK', primary: true, onClick: () => { const o = { what: st.what }; if (st.range === 'pages') { o.from = parseInt(from.value, 10) || 1; o.to = parseInt(to.value, 10) || o.from; } L.print.exportPDF(o); } }, { label: 'Cancel' }] }).done.then(res);
    });
  };
})(typeof window !== 'undefined' ? window : globalThis);
