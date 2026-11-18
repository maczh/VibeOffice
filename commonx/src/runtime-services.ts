/* Shared Vite-host settings and language services extracted from public/common; classic/Tauri pages keep their originals. */
// @ts-nocheck -- exact compatibility service bodies retain the established VO runtime contract.

const installSharedLookInstalled = new WeakSet<Window>();
export function installSharedLook(runtime: Window): void {
  if (installSharedLookInstalled.has(runtime)) return;
  installSharedLookInstalled.add(runtime);
  const window = runtime;
  (function () {
  'use strict';
  const VO = window.VO || (window.VO = {});
  const has = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
  /* the desktop app's copy */
  const desk = window.__VO_HOST__ && window.__TAURI_INTERNALS__ ? Object.assign({}, window.__VO_SETTINGS__) : null;
  delete window.__VO_SETTINGS__;
  const listeners = new Set();
  let sending = Promise.resolve();
  const take = (key, value) => {
    if ((has(desk, key) ? desk[key] : null) === value) return;
    if (value == null) delete desk[key]; else desk[key] = value;
    for (const fn of listeners) fn(key, value);
  };
  VO.settings = {
    /** the string kept for key, else dflt */
    get(key, dflt = null) {
      if (desk) return has(desk, key) ? desk[key] : dflt;
      try { const v = localStorage.getItem(key); return v == null ? dflt : v; } catch (e) { return dflt; }
    },
    /** keeps value (a string; null removes the key) for every window of the suite; false if it can't be kept */
    set(key, value) {
      value = value == null ? null : String(value);
      if (!desk) {
        try { if (value == null) localStorage.removeItem(key); else localStorage.setItem(key, value); return true; } catch (e) { return false; }
      }
      const before = has(desk, key) ? desk[key] : null;
      if (before === value) return true;
      if (value == null) delete desk[key]; else desk[key] = value;
      // one at a time, so the app takes them in the order they were made
      sending = sending.then(() => window.__TAURI_INTERNALS__.invoke('host_set_setting', { key, value })).catch((e) => {
        console.error('setting not kept:', key, e);
        if ((has(desk, key) ? desk[key] : null) === value) take(key, before);
      });
      return true;
    },
    remove: (key) => VO.settings.set(key, null),
    /** resolves once every change made here has reached the desktop app (before this page reloads) */
    kept: () => sending,
    /** fn(key, value) when another window changes a setting (value null: removed; key null: all may have); returns
     *  the function that stops it */
    on(fn) { listeners.add(fn); return () => listeners.delete(fn); },
  };
  if (desk) window.addEventListener('vo-settings', (e) => take(e.detail.key, e.detail.value));
  else window.addEventListener('storage', (e) => { if (e.storageArea === localStorage) for (const fn of listeners) fn(e.key, e.newValue); });

  const directRoot = () => window.__VIBE_OFFICE_DIRECT_HOST__?.root?.isConnected ? window.__VIBE_OFFICE_DIRECT_HOST__.root : null;
  const d = directRoot() || document.documentElement, l = VO.settings.get('vibeoffice.look') || 'classic';
  d.dataset.look = l;
  d.dataset.layout = l === 'paper2016' ? 'ribbon' : 'bars';

  })();
}

const installSharedLanguageInstalled = new WeakSet<Window>();
export function installSharedLanguage(runtime: Window): void {
  if (installSharedLanguageInstalled.has(runtime)) return;
  installSharedLanguageInstalled.add(runtime);
  const window = runtime;
  (function () {
  'use strict';
  const VO = window.VO || (window.VO = {});
  const LANGS = [['auto', 'Auto'], ['en', 'English'], ['fr', 'Français'], ['de', 'Deutsch'], ['zh-CN', '简体中文'],
    ['zh-Hant', '繁體中文'], ['ja', '日本語'], ['ko', '한국어']];
  const normalize = (tag) => {
    const s = String(tag || '').trim().toLowerCase();
    if (/^zh-(?:tw|hk|mo|hant)(?:-|$)/.test(s)) return 'zh-Hant';
    if (s === 'zh' || s.startsWith('zh-')) return 'zh-CN';
    const base = s.split('-')[0];
    return base === 'qps' || LANGS.some(([id]) => id === base) ? base : null;
  };
  const KEY = 'vibeoffice.lang';   // VO.settings (common/look.js)
  const stored = () => (VO.settings && VO.settings.get(KEY)) || 'auto';
  VO.lang = {
    normalize,
    list: () => LANGS.map(([id, name]) => ({ id, name })),
    get: () => normalize(stored()) || 'auto',
    set(id) {
      id = normalize(id);
      if (!LANGS.some(([l]) => l === id) || id === stored()) return;
      VO.settings.set(KEY, id);
    },
    resolve(choice, languages) {
      const id = normalize(choice);
      if (id && id !== 'auto') return id;
      return (languages || []).map(normalize).find((l) => l && l !== 'auto' && l !== 'qps') || 'en';
    },
  };
  if (typeof document !== 'undefined' && typeof location !== 'undefined') {
    const choice = new URLSearchParams(location.search).get('lang') || VO.lang.get();
    const lang = VO.lang.resolve(choice, navigator.languages || [navigator.language]);
    const directRoot = window.__VIBE_OFFICE_DIRECT_HOST__?.root;
    (directRoot && directRoot.isConnected ? directRoot : document.documentElement).lang = lang === 'qps' ? 'qps-ploc' : lang;
  }

  })();
}

const installSharedI18nInstalled = new WeakSet<Window>();
export function installSharedI18n(runtime: Window, scriptUrl: URL): void {
  if (installSharedI18nInstalled.has(runtime)) return;
  installSharedI18nInstalled.add(runtime);
  const window = runtime;
  (function () {
  'use strict';
  const VO = window.VO || (window.VO = {});
  const doc = typeof document !== 'undefined' ? document : null;
  const directRoot = typeof window !== 'undefined' ? window.__VIBE_OFFICE_DIRECT_HOST__?.root : null;
  const tag = (directRoot && directRoot.isConnected ? directRoot : doc && doc.documentElement).getAttribute('lang') || 'en';
  const lang = VO.lang.normalize(tag) || 'en';

  /* qps: each string of the English catalog (i18n/en.json, tools/i18n.mjs) mechanically changed, as a complete
     translation would change it: it stands out (accented, in brackets) and is about 40% longer, as French and German
     are; {0} placeholders and the letter after & stay as they are. English left on screen isn't in the catalog. */
  const ACC = { a: 'á', c: 'ç', e: 'é', i: 'í', n: 'ñ', o: 'ó', u: 'ú', y: 'ý', A: 'Á', C: 'Ç', E: 'É', I: 'Í', N: 'Ñ', O: 'Ó', U: 'Ú', Y: 'Ý' };
  const pseudo = (s) => {
    const body = s.replace(/(\{\d+\}|&.)|[aceinouyACEINOUY]/g, (m, keep) => keep || ACC[m]);
    const letters = s.replace(/\{\d+\}|[^A-Za-z]/g, '').length;
    return '[' + body + (letters > 2 ? ' ' + '·'.repeat(Math.ceil(letters * 0.4)) : '') + ']';
  };
  const marks = (s) => (s.match(/\{\d+\}/g) || []).sort().join();
  /** the text without its context: "Line|chart" is "Line". A context is words (letters, spaces, hyphens) right after
      the "|" that ends the text; other text with a "|" is shown whole, as text that is the user's can reach VO.t again
      (a file name A|B.docx filled into a message, "Revenue | Expenses") */
  const plain = (s) => s.replace(/(\S)\|[A-Za-z][A-Za-z -]*$/, '$1');
  /* Locale catalogs are also consumed by the out-of-scope desktop app; keep them unchanged and brand only web UI. */
  const webBrand = (s) => window.__VO_HOST__ && window.__TAURI_INTERNALS__ ? s : s.replace(/\bQuire\b/g, 'Wordx').replace(/\bLedger\b/g, 'Excelx').replace(/\bLectern\b/g, 'PPTX');

  let dict = null;
  /* in the pseudo-locale, the text asked for that the catalog lacks (tools/i18n.mjs harvest reads it) */
  const missed = new Set();
  const t = (s, ...args) => {
    if (typeof s !== 'string' || !s) return s;
    const has = (k) => dict && Object.prototype.hasOwnProperty.call(dict, k) && typeof dict[k] === 'string' && dict[k];
    const base = plain(s);
    let out = base;
    if (lang === 'qps') { if (!dict || Object.prototype.hasOwnProperty.call(dict, s)) out = pseudo(webBrand(base)); else if (/[A-Za-z]{2}/.test(s) && s[0] !== '[') missed.add(s); }
    else out = webBrand([s, base].map((k) => has(k) && marks(dict[k]) === marks(base) && dict[k]).find(Boolean) || base);
    return args.length ? out.replace(/\{(\d+)\}/g, (m, i) => (args[i] != null ? String(args[i]) : m)) : out;
  };

  /* the dictionary beside this file (for qps, the catalog) */
  const here = scriptUrl.href;
  let loaded = lang === 'en' || !here || typeof fetch !== 'function';
  const ready = loaded ? Promise.resolve() : fetch(new URL('i18n/' + (lang === 'qps' ? 'en' : lang) + '.json', here))
    .then((r) => (r.ok ? r.json() : null))
    .then((d) => { if (d && typeof d === 'object') dict = d; })
    .catch(() => {})
    .then(() => { loaded = true; });

  VO.i18n = {
    /** the language of this page's UI (common/lang.js) */
    lang,
    /** the locales argument of toLocaleDateString and the like for dates the UI shows: the UI's language, or with
        English the browser's own (an English UI in Britain writes 14 Oct) */
    dates: lang === 'en' || lang === 'qps' ? [] : lang,
    ready,
    /** run fn once the dictionary is here: now if it is (English), else when it arrives */
    whenReady: (fn) => (loaded ? fn() : ready.then(fn)),
    missed,
    t,
    plain,
  };
  VO.t = t;
  /** UI text kept in a table and translated where it is shown (VO.t then): marks it for the catalog, returns it as is */
  VO.N = (s) => s;

  /* the page's own HTML (the Start Center's text, a spreadsheet formula bar): its text and the attributes the UI shows,
     translated once, before the app builds the rest of its UI with VO.t (tools/i18n.mjs gathers the same text) */
  const ATTRS = ['title', 'placeholder', 'aria-label', 'data-tip', 'alt'];
  const translateTree = (top) => {
    if (lang === 'en' || !top) return;
    /* not inside translate="no" (data: a version, a language's own name) */
    const walker = doc.createTreeWalker(top, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT, { acceptNode: (n) => (n.nodeType === 1 && n.getAttribute('translate') === 'no' ? 2 : 1) });   // FILTER_REJECT skips the subtree
    for (let n = walker.currentNode; n; n = walker.nextNode()) {
      if (n.nodeType === 1) {
        for (const a of ATTRS) { const v = n.getAttribute(a); if (v) { const x = t(v); if (x !== v) n.setAttribute(a, x); } }
        continue;
      }
      const p = n.parentNode && n.parentNode.nodeName;
      if (p === 'SCRIPT' || p === 'STYLE') continue;
      const s = n.data.replace(/\s+/g, ' ').trim();
      if (!s || !/[A-Za-z]/.test(s)) continue;
      const x = t(s);
      if (x !== s) n.data = n.data.match(/^\s*/)[0] + x + n.data.match(/\s*$/)[0];
    }
  };
  VO.i18n.translateTree = translateTree;
  if (doc) {
    const run = () => VO.i18n.whenReady(() => translateTree(directRoot && directRoot.isConnected ? directRoot : doc.documentElement));
    if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', run); else run();
  }
  if (window.L) { window.L.t = t; window.L.N = VO.N; }

  })();
}

export function installSharedRuntimeService(runtime: Window, scriptUrl: URL): void {
  if (scriptUrl.pathname.endsWith('/common/look.js')) installSharedLook(runtime);
  else if (scriptUrl.pathname.endsWith('/common/lang.js')) installSharedLanguage(runtime);
  else if (scriptUrl.pathname.endsWith('/common/i18n.js')) installSharedI18n(runtime, scriptUrl);
}

export function initializeSharedRuntime(runtime: Window = window): void {
  installSharedLook(runtime);
  installSharedLanguage(runtime);
  installSharedI18n(runtime, new URL('/common/i18n.js', runtime.location.href));
}
