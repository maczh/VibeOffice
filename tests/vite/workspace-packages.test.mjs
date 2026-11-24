import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { WordxEditor } from '@vibeoffice/wordx';
import { ExcelxEditor } from '@vibeoffice/excelx';
import { PptxEditor } from '@vibeoffice/pptx';
import { AppearanceControls, OfficeDialog, VibeOfficeProvider, installClipArtService, installColorPickerService, installCommandMatcher, installCommandSearchService, installReactCommandService, installReactToolbarService, installReactTooltipService, installReactUiMountService, installReactUiRefreshService, installReactUiTextService, standardColorName } from '@vibeoffice/core';
import { createDirectEngineBridge, createIndexedDbDraftStore, openLocalDocument, resolveUnsavedChanges, saveLocalDocument } from '@vibeoffice/core';
import { IDBFactory } from 'fake-indexeddb';
import { createRemoteFileCallbacks, createUrlReadProvider, createWebDavProvider, remoteConflict } from '@vibeoffice/core';
import { reactCsp } from '../../tools/vite/react-csp.mjs';
import { editorRuntime, legacyRuntime, reactTemplates } from '../../tools/vite/legacy-runtime.mjs';

const editors = [
  { component: WordxEditor, label: 'Wordx editor', id: 'wordx' },
  { component: ExcelxEditor, label: 'Excelx editor', id: 'excelx' },
  { component: PptxEditor, label: 'PPTX editor', id: 'pptx' },
];

test('React Clip Art catalog restores its previous owner on editor disposal', () => {
  const previous = { search: () => ['host'] };
  const runtime = { L: { clipart: previous } };
  const dispose = installClipArtService(runtime);
  const installed = runtime.L.clipart;
  assert.notEqual(installed, previous);
  assert.ok(installed.search('rocket').length > 0);
  dispose();
  assert.equal(runtime.L.clipart, previous);

  const disposeOwned = installClipArtService(runtime);
  const laterOwner = { search: () => ['later'] };
  runtime.L.clipart = laterOwner;
  disposeOwned();
  assert.equal(runtime.L.clipart, laterOwner, 'cleanup must not overwrite a newer service owner');
});

test('React command matcher initializes without loading the legacy matcher script', () => {
  const runtime = { VO: {} };
  installCommandMatcher(runtime);
  assert.deepEqual(runtime.VO.commandMatcher.words('AutoSum 搜索'), ['autosum', '搜索', 'auto', 'sum']);
  const entry = { label: 'Page Setup', alias: 'margins orientation', tip: '', path: [] };
  assert.equal(runtime.VO.commandMatcher.scoreWord('margin', entry), 3.5, 'Office-friendly synonyms keep their ranking');
  assert.equal(runtime.VO.commandMatcher.scoreWord('pag', entry), 4, 'prefix matches keep their ranking');
  const hostMatcher = { words: () => ['host'], scoreWord: () => 1, maxResults: 2, maxRecent: 1 };
  installCommandMatcher({ VO: { commandMatcher: hostMatcher } });
  assert.equal(runtime.VO.commandMatcher.maxResults, 10, 'the React package owns its fallback matcher');
  const embedded = { VO: { commandMatcher: hostMatcher } };
  installCommandMatcher(embedded);
  assert.equal(embedded.VO.commandMatcher, hostMatcher, 'an embedding host may provide its own matcher');
  const classicMatcher = { ...hostMatcher };
  Object.defineProperty(classicMatcher, Symbol.for('vibeoffice.classicCommandMatcher'), { value: true });
  const reactRuntime = { VO: { commandMatcher: classicMatcher } };
  installCommandMatcher(reactRuntime);
  assert.notEqual(reactRuntime.VO.commandMatcher, classicMatcher, 'React replaces only the explicitly marked classic fallback');
  assert.deepEqual(reactRuntime.VO.commandMatcher.words('AutoSum 搜索'), ['autosum', '搜索', 'auto', 'sum']);
  assert.match(readFileSync(new URL('../../wordx/src/editor.tsx', import.meta.url), 'utf8'), /installCommandSearchService\(runtime\)/);
  for (const app of ['wordx', 'excelx', 'pptx']) {
    const editor = readFileSync(new URL(`../../${app}/src/editor.tsx`, import.meta.url), 'utf8');
    assert.match(editor, /installCommandSearchService\(runtime\)/);
    assert.match(editor, /skipLegacyScripts:[^\n]*\/common\/search\.js/);
  }
});

test('app installer modules bind host APIs from their supplied runtime window', () => {
  for (const [app, file] of [
    ['wordx', 'application.ts'], ['wordx', 'app-dialogs.ts'], ['wordx', 'commands.ts'], ['wordx', 'drawing.ts'], ['wordx', 'proofing.ts'], ['wordx', 'dialogs.ts'],
    ['excelx', 'application.ts'], ['excelx', 'panes.ts'], ['excelx', 'grid-runtime.ts'], ['pptx', 'app.ts'], ['pptx', 'panes.ts'], ['pptx', 'textedit.ts'],
  ]) {
    const source = readFileSync(new URL(`../../${app}/src/${file}`, import.meta.url), 'utf8');
    assert.match(source, /const VO = (?:window\.VO|root\.VO|\((?:window|root) as Window & \{ VO\?: any \}\)\.VO);/, `${app}/${file} must not resolve host APIs from the ambient window`);
  }
  const application = readFileSync(new URL('../../wordx/src/application.ts', import.meta.url), 'utf8');
  assert.match(application, /const host = document\.querySelector\('\[data-vibeoffice-engine-root\]'\) \|\| document\.body;/, 'the installer-provided document remains the mount owner');
  for (const file of ['task-content.tsx', 'status-bar.tsx', 'selection-overlay.tsx', 'print-preview.tsx', 'notes.tsx', 'show-controls.tsx', 'slides-pane.tsx', 'rulers.tsx', 'grid-layer.tsx', 'notes-page.tsx', 'shape-grid-menu.tsx']) {
    const source = readFileSync(new URL(`../../pptx/src/${file}`, import.meta.url), 'utf8');
    assert.match(source, /function runtime\(host: HTMLElement\)[\s\S]*?host\.ownerDocument\.defaultView/, `PPTX ${file} resolves engine APIs from its host document`);
    assert.doesNotMatch(source, /\bruntime\(\)/, `PPTX ${file} must not use ambient-window engine lookup`);
    assert.doesNotMatch(source, /\bwindow\.VO\b/, `PPTX ${file} must resolve host APIs through its owner document`);
  }
  for (const file of ['diagram-menu.tsx', 'line-style-menu.tsx', 'shadow-style-menu.tsx']) {
    const source = readFileSync(new URL(`../../pptx/src/${file}`, import.meta.url), 'utf8');
    assert.match(source, /host\.ownerDocument[\s\S]*?ownerDocument\.defaultView/, `PPTX ${file} resolves runtime and browser APIs from its host`);
    assert.doesNotMatch(source, /\bwindow\.|\bdocument\./, `PPTX ${file} must not attach menus to the ambient browser context`);
  }
  for (const file of ['zoom-dialog.tsx', 'grid-dialog.tsx', 'date-time-dialog.tsx', 'change-case-dialog.tsx', 'bullets-dialog.tsx', 'autocorrect-dialog.tsx', 'symbol-dialog.tsx', 'line-spacing-dialog.tsx', 'header-footer-dialog.tsx', 'setup-show-dialog.tsx', 'replace-fonts-dialog.tsx', 'insert-table-dialog.tsx', 'font-dialog.tsx', 'action-settings-dialog.tsx', 'print-dialog.tsx', 'properties-dialog.tsx', 'autocontent-dialog.tsx', 'find-dialog.tsx', 'page-setup-dialog.tsx', 'hyperlink-dialog.tsx', 'save-as-dialog.tsx', 'color-scheme-dialog.tsx', 'diagram-gallery-dialog.tsx', 'table-format-dialog.tsx', 'paste-special-dialog.tsx', 'speaker-notes-dialog.tsx', 'effect-options-dialog.tsx', 'slide-finder-dialog.tsx', 'wordart-gallery-dialog.tsx', 'wordart-text-dialog.tsx', 'options-dialog.tsx', 'photo-album-dialog.tsx', 'fill-effects-dialog.tsx', 'format-shape-dialog.tsx']) {
    const source = readFileSync(new URL(`../../pptx/src/${file}`, import.meta.url), 'utf8');
    assert.match(source, /usePptxRuntime<[^>]+>\(\)/, `PPTX ${file} resolves its engine through the scoped host provider`);
    assert.doesNotMatch(source, /\bwindow\.|function runtime\(\)/, `PPTX ${file} must not use ambient-window runtime lookup`);
  }
  const pptxEditor = readFileSync(new URL('../../pptx/src/editor.tsx', import.meta.url), 'utf8');
  assert.match(pptxEditor, /appLayers=\{\(host\) => <PptxHostProvider host=\{host\}>/, 'PPTX app layers receive the editor host runtime context');
  const pptxInteractions = readFileSync(new URL('../../pptx/src/editor-interactions.tsx', import.meta.url), 'utf8');
  assert.match(pptxInteractions, /function interactionRuntime\(host: HTMLElement\)[\s\S]*?host\.ownerDocument\.defaultView/, 'PPTX editor interactions resolve their runtime from the owning host');
  assert.doesNotMatch(pptxInteractions, /\binteractionRuntime\(\)|\bwindow\.|\bdocument\./, 'PPTX editor interactions must bind events and services to the owner document/window');
  for (const [app, file] of [['excelx', 'commands.ts'], ['pptx', 'commands.ts']]) {
    const source = readFileSync(new URL(`../../${app}/src/${file}`, import.meta.url), 'utf8');
    assert.match(source, /const window = runtime as Window & \{ L\?: any; VO\?: any \};/, `${app}/${file} must resolve host APIs from its installer-provided runtime`);
    assert.doesNotMatch(source, /\bglobalThis\.VO\b|\bwindow\s*=\s*globalThis\.window/, `${app}/${file} must not escape its supplied runtime`);
  }
});

test('React command registry preserves command descriptors, execution hooks, and host replacements', () => {
  const runtime = new EventTarget();
  runtime.CustomEvent = CustomEvent;
  runtime.console = { warn() {}, error() {} };
  const audit = [];
  const ui = {
    cmds: {}, hooks: { beforeExec() { audit.push('before'); } },
    refresh() { audit.push('refresh'); }, msg(text, options) { audit.push(['message', text, options.icon]); },
  };
  runtime.L = { ui, t: (text, detail) => text.replace('{0}', detail) };
  const commands = ui.cmds;
  let eventDetail;
  runtime.addEventListener('vo-command', (event) => { eventDetail = event.detail; });
  const originalCmd = () => 'classic';
  ui.cmd = originalCmd;
  const dispose = installReactCommandService(runtime);
  assert.equal(ui.cmds, commands, 'React reuses the established command catalog object');
  const definition = { label: '&Open', enabled: () => true, checked: () => true, run: (arg, event) => ({ arg, event }) };
  const command = ui.cmd('open', definition);
  assert.equal(command.label, '&Open');
  assert.equal(ui.enabled('open'), true);
  assert.equal(ui.checked('open'), true);
  const existing = Object.create({ inherited: 'kept' });
  Object.defineProperties(existing, { label: { value: '&Save', enumerable: true }, run: { value: () => 'saved', enumerable: true } });
  const registered = ui.cmd('save', existing);
  assert.notEqual(registered, existing, 'registration wraps a definition using the legacy contract');
  assert.equal(Object.getPrototypeOf(registered), Object.prototype);
  assert.equal(Object.getOwnPropertyDescriptor(registered, 'label').enumerable, true);
  assert.equal(registered.inherited, undefined);
  const nativeEvent = new Event('click');
  const result = ui.exec('open', 'argument', nativeEvent);
  assert.deepEqual(result, { arg: 'argument', event: nativeEvent });
  assert.deepEqual(audit, ['before', 'refresh']);
  assert.deepEqual(eventDetail, { id: 'open', result });
  command.run = () => { throw new Error('test failure'); };
  assert.equal(ui.exec('open'), undefined);
  assert.deepEqual(audit.slice(-3), ['before', ['message', 'That action could not be completed: test failure', 'warn'], 'refresh']);
  const laterOwner = () => 'new owner';
  ui.exec = laterOwner;
  dispose();
  assert.equal(ui.exec, laterOwner, 'disposal leaves an API installed by a later owner untouched');
  assert.equal(ui.cmd, originalCmd, 'disposal restores classic APIs still owned by this service');
});

test('React command-search service validates its host and restores only APIs it still owns', () => {
  assert.throws(() => installCommandSearchService({ VO: {} }), /requires the editor UI service/);
  const originals = {
    stripAmp: (text) => text, plainAmp: (text) => text, enabled: () => true,
    searchCommands: () => 'classic-search', searchResults: () => 'classic-results', searchRemember: () => 'classic-remember',
    searchBox: () => 'classic-box', searchInput: () => 'classic-input',
  };
  const runtime = {
    CustomEvent,
    dispatchEvent() {},
    VO: {},
    console: { warn() {}, error() {} },
    document: { documentElement: { dataset: { layout: 'bars' } }, createElement: () => ({ setAttribute() {}, className: '', hidden: false }) },
    L: { ui: { ...originals, stripAmp: (text) => text, plainAmp: (text) => text, ampHTML: (text) => text, enabled: () => true, cmds: {}, menus: [], $: () => [], $$: () => [] }, t: (text) => text },
  };
  const ui = runtime.L.ui;
  const dispose = installCommandSearchService(runtime);
  assert.notEqual(ui.searchCommands, originals.searchCommands);
  const replacement = () => 'new owner';
  ui.searchInput = replacement;
  dispose();
  assert.equal(ui.searchCommands, originals.searchCommands);
  assert.equal(ui.searchResults, originals.searchResults);
  assert.equal(ui.searchRemember, originals.searchRemember);
  assert.equal(ui.searchBox, originals.searchBox);
  assert.equal(ui.searchInput, replacement, 'disposal leaves APIs installed by a later owner untouched');
});

test('shared package declares every React runtime it imports as a peer dependency', () => {
  const metadata = JSON.parse(readFileSync(new URL('../../commonx/package.json', import.meta.url), 'utf8'));
  assert.equal(metadata.peerDependencies.react, '^19.0.0');
  assert.equal(metadata.peerDependencies['react-dom'], '^19.0.0');
  assert.match(readFileSync(new URL('../../commonx/src/legacy-dialog.tsx', import.meta.url), 'utf8'), /from 'react-dom\/client'/);
});

test('React toolbar service restores classic engine APIs on disposal', () => {
  const dispatched = [];
  const root = { isConnected: true, dispatchEvent(event) { dispatched.push(event); return true; } };
  const classicToolbar = () => 'classic';
  const classicButton = () => 'classic-button';
  const ui = {
    toolbar: classicToolbar, tbButton: classicButton, tbSplit: () => 'classic-split', tbDrop: () => 'classic-drop', reactToolbar: () => false,
    stripAmp: (text) => text, ampHTML: (text) => text, cmds: {}, exec() {},
  };
  const runtime = { L: { ui, h: () => ({}) }, VO: { reactDirectHost: { root } }, CustomEvent };
  const dispose = installReactToolbarService(runtime);
  assert.notEqual(ui.toolbar, classicToolbar);
  assert.equal(ui.tbButton('save').__vibeToolbarControl, 'button');
  const target = {};
  assert.equal(ui.reactToolbar('standard', 'Standard', target, []), true);
  assert.equal(dispatched[0].type, 'vibeoffice-react-toolbar');
  assert.equal(dispatched[0].detail.target, target);
  dispose();
  assert.equal(ui.toolbar, classicToolbar);
  assert.equal(ui.tbButton, classicButton);
  assert.equal(ui.reactToolbar(), false);
});

test('React color palette names are translated without the legacy shared UI initializer', () => {
  assert.equal(standardColorName('#000000'), 'Black');
  assert.equal(standardColorName('#800000', (label) => `translated:${label}`), 'translated:Dark Red');
  assert.equal(standardColorName('#abc123'), '#ABC123');

  const events = [];
  const root = { isConnected: true, dispatchEvent(event) { events.push(event); return true; } };
  const colors = ['#000000', '#800000', '#ffffff', '#abc123'];
  const ui = {};
  const runtime = {
    L: { ui, color: { STANDARD: colors }, t: (label) => `translated:${label}` },
    VO: { reactDirectHost: { root } },
    CustomEvent,
  };
  installColorPickerService(runtime);
  assert.equal(ui.colorName, undefined, 'React color labels do not depend on legacy ui.colorName');
  ui.colorMenu({ left: 12, bottom: 34 }, { mode: 'plain' }, () => {});
  const request = events[0].detail;
  const palette = request.sections.find((section) => section.layout === 'grid').items;
  assert.deepEqual(palette.map(({ label }) => label), ['translated:Black', 'translated:Dark Red', 'translated:White', '#ABC123']);
});

test('React tooltip service disposes only the hide API it owns', () => {
  const events = [];
  const root = { isConnected: true, dispatchEvent(event) { events.push(event); return true; } };
  const previous = () => 'classic';
  const ui = { hideTip: previous };
  const runtime = { L: { ui }, VO: { reactDirectHost: { root } }, CustomEvent };
  const dispose = installReactTooltipService(runtime);
  assert.notEqual(ui.hideTip, previous);
  ui.hideTip();
  assert.equal(events[0].type, 'vibeoffice-react-hide-tooltip');
  dispose();
  assert.equal(ui.hideTip, previous, 'disposing the React tooltip restores the classic API');

  const disposeOwned = installReactTooltipService(runtime);
  const laterOwner = () => 'later owner';
  ui.hideTip = laterOwner;
  disposeOwned();
  assert.equal(ui.hideTip, laterOwner, 'disposing never overwrites a newer API owner');
});

test('React editor mount service leaves lifecycle ownership with the direct engine', () => {
  const root = { isConnected: true };
  const mounted = [];
  const unmounted = [];
  const previousMount = (host) => { mounted.push(host); return () => {}; };
  const previousUnmount = (host) => unmounted.push(host);
  const ui = { mount: previousMount, unmount: previousUnmount };
  const runtime = { L: { ui }, VO: { reactDirectHost: { root } } };
  const dispose = installReactUiMountService(runtime);
  const host = {};
  const release = ui.mount(host);
  assert.equal(typeof release, 'function');
  ui.unmount(host);
  release();
  assert.deepEqual(mounted, [], 'React does not start classic UI global listeners or mount bookkeeping');
  assert.deepEqual(unmounted, [], 'React unmount does not run the classic UI teardown');
  dispose();
  assert.equal(ui.mount, previousMount);
  assert.equal(ui.unmount, previousUnmount);
});

test('React command refresh notifies Ribbon without retaining the legacy toolbar refresh loop', () => {
  const events = [];
  const classicRefresh = () => events.push('classic');
  const ui = { refresh: classicRefresh };
  const runtime = {
    L: { ui, rafThrottle: (callback) => callback, bus: { emit: (event) => events.push(event) } },
    VO: { reactDirectHost: { root: {} } },
  };
  const dispose = installReactUiRefreshService(runtime);
  assert.notEqual(ui.refresh, classicRefresh);
  ui.refresh();
  assert.deepEqual(events, ['ui-refresh']);
  dispose();
  assert.equal(ui.refresh, classicRefresh);

  const disposeOwned = installReactUiRefreshService(runtime);
  const laterOwner = () => {};
  ui.refresh = laterOwner;
  disposeOwned();
  assert.equal(ui.refresh, laterOwner, 'disposal must not overwrite a newer refresh owner');
});

test('React shared text helpers preserve translations, access keys, and platform shortcut labels', () => {
  const translated = {
    '&Open|menu': '&Ouvrir|menu',
    'Ctrl|key': 'Strg', 'Shift|key': 'Umschalt', 'Alt|key': 'Alt',
  };
  const ui = {
    plainAmp: () => 'classic-plain', stripAmp: () => 'classic-strip', ampHTML: () => 'classic-html',
    accessKey: () => 'classic-key', keyText: () => 'classic-shortcut',
  };
  const runtime = {
    L: { ui, t: (text) => (translated[text] || text).replace(/\|[A-Za-z][A-Za-z -]*$/, ''), esc: (text) => String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'), isMac: false },
    VO: { reactDirectHost: { root: {} }, i18n: { plain: (text) => text.replace(/\|[A-Za-z][A-Za-z -]*$/, '') } },
  };
  const previous = { ...ui };
  const dispose = installReactUiTextService(runtime);
  assert.equal(ui.plainAmp('&&Keep &X'), '&Keep X');
  assert.equal(ui.stripAmp('&Open|menu'), 'Ouvrir');
  assert.equal(ui.ampHTML('Save && E&xit'), 'Save &amp; E<u>x</u>it');
  assert.equal(ui.ampHTML('Number|Format Cells tab'), 'Number');
  assert.equal(ui.accessKey('&Open|menu'), 'o');
  assert.equal(ui.keyText('Ctrl+Shift+F12'), 'Strg+Umschalt+F12');
  runtime.L.isMac = true;
  assert.equal(ui.keyText('Alt+Shift+Ctrl+Enter'), '⌥⇧⌘Enter');
  dispose();
  assert.deepEqual(ui, previous, 'classic engine helper APIs are restored on editor disposal');

  const disposeOwned = installReactUiTextService(runtime);
  const laterOwner = () => 'later owner';
  ui.stripAmp = laterOwner;
  disposeOwned();
  assert.equal(ui.stripAmp, laterOwner, 'disposal never overwrites a newer text-helper owner');
});

test('editor packages can be packed with the shared peer and same-origin engine assets', () => {
  const core = JSON.parse(readFileSync(new URL('../../commonx/package.json', import.meta.url), 'utf8'));
  assert.notEqual(core.private, true, '@vibeoffice/core must be publishable because each editor declares it as a peer');
  assert.deepEqual(core.files, ['dist']);
  assert.equal(core.scripts.prepack, 'npm run build');
  assert.equal(core.exports['./style.css'], './dist/index.css');
  assert.ok(readFileSync(new URL('../../commonx/dist/index.css', import.meta.url), 'utf8').length > 0, '@vibeoffice/core stylesheet export points to the built shared styles');
  for (const app of ['wordx', 'excelx', 'pptx']) {
    const metadata = JSON.parse(readFileSync(new URL(`../../${app}/package.json`, import.meta.url), 'utf8'));
    assert.ok(metadata.files.includes('dist'), `${app} ships its ESM component, declarations, and stylesheet`);
    assert.ok(metadata.files.includes('standalone'), `${app} ships the same-origin static engine asset tree`);
    assert.equal(metadata.scripts.prepack, 'npm run build && npm run build:app', `${app} builds package and engine assets before packing`);
    assert.equal(metadata.peerDependencies['@vibeoffice/core'], core.version);
    assert.equal(metadata.license, core.license);
    assert.ok(readFileSync(new URL(`../../${app}/standalone/${app}/runtime.json`, import.meta.url), 'utf8').includes('../common/looks.css'), `${app} runtime manifest resolves shared assets from the standalone tree`);
    const retired = { wordx: 'quire', excelx: 'ledger', pptx: 'lectern' }[app];
    assert.ok(existsSync(new URL(`../../suite/dist/${app}`, import.meta.url)), `suite output contains the ${app} engine route`);
    assert.ok(!existsSync(new URL(`../../suite/dist/${retired}`, import.meta.url)), `suite output omits the retired ${retired} engine route`);
    assert.ok(!existsSync(new URL(`../../${app}/standalone/${retired}`, import.meta.url)), `${app} standalone output omits the retired ${retired} engine route`);
  }
});

test('React production engine assets are emitted only under the current product routes', () => {
  const adapter = readFileSync(new URL('../../tools/vite/legacy-runtime.mjs', import.meta.url), 'utf8');
  assert.match(adapter, /for \(const name of apps\) \{\s*const retired = path\.join\(output, legacyFolders\[name\]\);\s*await rm\(retired, \{ recursive: true, force: true \}\);/);
  assert.match(adapter, /\['common', 'common'\]/, 'shared engine assets remain available to the renamed route');
  for (const [app, retired] of [['wordx', 'quire'], ['excelx', 'ledger'], ['pptx', 'lectern']]) {
    assert.match(adapter, new RegExp(`${app}: '${retired}'`), `${app} engine assets are copied to the current product route`);
  }
});

test('quick checks and testing docs use the branded app UI suite paths', () => {
  const quick = readFileSync(new URL('../../tools/quick.mjs', import.meta.url), 'utf8');
  const docs = readFileSync(new URL('../../docs/testing.md', import.meta.url), 'utf8');
  for (const app of ['wordx', 'excelx', 'pptx']) {
    const path = `tests/ui/${app}.test.mjs`;
    assert.match(quick, new RegExp(path.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
    assert.match(docs, new RegExp(`${app}\\.test\\.mjs`));
    assert.ok(readFileSync(new URL(`../../${path}`, import.meta.url), 'utf8').length > 0, `${path} exists`);
  }
  for (const oldPath of ['quire', 'ledger', 'lectern']) {
    assert.doesNotMatch(quick, new RegExp(`tests/ui/${oldPath}\\.test\\.mjs`));
    assert.doesNotMatch(docs, new RegExp(`${oldPath}\\.test\\.mjs`));
  }
});

test('a host can mount each editor package through its public React export', () => {
  for (const { component, label, id } of editors) {
    const html = renderToStaticMarkup(createElement(component, { className: 'host-embed', services: { translate: (message) => message } }));
    assert.match(html, new RegExp(`aria-label="${label}"`));
    assert.match(html, new RegExp(`data-vibe-editor="${id}"`));
    assert.match(html, /class="vibe-editor host-embed"/);
  }
});

test('Wordx exposes a React-owned engine mount point inside its editor frame', () => {
  const html = renderToStaticMarkup(createElement(WordxEditor, {
    mountEngine: () => {},
    children: createElement('p', null, 'Engine fixture'),
  }));
  assert.match(html, /data-vibe-editor="wordx"/);
  assert.match(html, /class="vibe-editor__engine"><p>Engine fixture<\/p><\/div>/);
});

test('async editor mounts dispose late resources and report only live mount failures', async () => {
  const { settleEditorMount } = await import('@vibeoffice/core');
  let disposed = false, disposedLateMount = 0, reported = 0;
  let resolveMount;
  const pending = new Promise((resolve) => { resolveMount = resolve; });
  settleEditorMount(pending, () => disposed, () => assert.fail('disposed mount must not be retained'), () => assert.fail('successful mount must not report an error'));
  disposed = true;
  resolveMount(() => disposedLateMount++);
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(disposedLateMount, 1, 'a disposer returned after unmount runs immediately');

  let rejectMount;
  const lateFailure = new Promise((_resolve, reject) => { rejectMount = reject; });
  settleEditorMount(lateFailure, () => true, () => assert.fail('rejected mount must not be retained'), () => reported++);
  rejectMount(new Error('late failure'));
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(reported, 0, 'rejection after unmount is consumed without calling stale UI');

  settleEditorMount(Promise.reject(new Error('active failure')), () => false, () => {}, () => reported++);
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(reported, 1, 'a live mount failure reaches its error handler');
});

test('React packages expose the same-document direct engine mount adapter', () => {
  assert.equal(typeof createDirectEngineBridge, 'function');
  const frame = readFileSync(new URL('../../commonx/src/editor-frame.tsx', import.meta.url), 'utf8');
  const runtime = readFileSync(new URL('../../commonx/src/direct-engine.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(frame, /iframe|contentWindow|postMessage/);
  assert.match(runtime, /mountLegacyEngineDocument/);
  assert.doesNotMatch(runtime, /createElement\(['"]iframe/);
  assert.match(runtime, /shellOwnedByReact && scriptUrl\.pathname\.endsWith\('\/common\/ribbon\.js'\)/, 'React-owned editor shells skip the classic shared Ribbon module');
  assert.match(runtime, /installReactUiMountService/, 'React-owned engine mounts bypass the classic shared UI mount lifecycle');
  const services = readFileSync(new URL('../../commonx/src/runtime-services.ts', import.meta.url), 'utf8');
  const ribbon = readFileSync(new URL('../../commonx/src/ribbon-service.ts', import.meta.url), 'utf8');
  assert.match(services, /scriptUrl\.pathname\.endsWith\('\/common\/ribbon\.js'\).*installReactRibbonService/);
  assert.match(services, /VO\.look\s*=\s*\{/);
  assert.match(services, /VO\.settings\.on\(\(key\) => \{ if \(key === LOOK \|\| key === null\) applyLook\(/);
  assert.match(runtime, /directRuntime\.reactHost = true/);
  for (const method of ['init', 'select', 'collapse', 'built', 'follow']) assert.match(ribbon, new RegExp(`ribbonUI\\.${method}\\s*=`), `React Ribbon adapter provides ${method}`);
});

test('the Vite host runtime leaves React package source routes to Vite when no legacy asset matches', async () => {
  let middleware;
  legacyRuntime(['wordx', 'excelx', 'pptx']).configureServer({ middlewares: { use(handler) { middleware = handler; } } });
  assert.equal(typeof middleware, 'function');
  const forwarded = await new Promise((resolve, reject) => middleware(
    { url: '/wordx/src/index.ts' },
    { setHeader() {}, end() { reject(new Error('a non-existent legacy asset was served')); } },
    (error) => error ? reject(error) : resolve(true),
  ));
  assert.equal(forwarded, true, 'the React source alias can resolve instead of receiving a legacy-runtime 500');
});

test('React editor runtime manifests preserve ordered classic scripts and recover styles from generated legacy HTML', () => {
  const manifest = JSON.parse(editorRuntime('<html><head><link rel="stylesheet" href="../common/ui.css"><style media="screen">.shell{display:flex}</style></head><div id="app"></div><script src="../common/command-match.js?v=1"></script><script src="../common/core.js"></script><script src="js/app.js"></script></html>'));
  assert.deepEqual(manifest.styles, [
    { href: '../common/ui.css', media: '' },
    { css: '.shell{display:flex}', media: 'screen' },
  ]);
  assert.deepEqual(manifest.scripts, ['../common/core.js', 'js/app.js']);
  assert.doesNotMatch(manifest.scripts.join('\n'), /command-match\.js/, 'the React runtime uses its package-owned matcher, not the classic script');
});

test('React entries initialize shared appearance services from the package instead of classic script tags', () => {
  for (const page of ['wordx/index.html', 'excelx/index.html', 'pptx/index.html', 'suite/index.html']) {
    const html = readFileSync(new URL(`../../${page}`, import.meta.url), 'utf8');
    assert.doesNotMatch(html, /<script\b[^>]*src=["'][^"']*\/common\/(?:look|lang|i18n)\.js["']/i, `${page} must bootstrap shared services from the React package`);
  }
  const runtime = readFileSync(new URL('../../commonx/src/runtime-services.ts', import.meta.url), 'utf8');
  assert.match(runtime, /export function initializeSharedRuntime/);
  assert.match(runtime, /export function installSharedRuntimeService/);
});

test('Excelx package stylesheet includes shared and feature-specific React dialog styles', () => {
  const style = readFileSync(new URL('../../excelx/dist/style.css', import.meta.url), 'utf8');
  assert.match(style, /\.vibeoffice-dialog/);
  assert.match(style, /\.excelx-find-dialog\{position:fixed/);
  assert.match(style, /\.excelx-find-dialog__results/);
  assert.match(style, /\.excelx-insert-function-dialog__search/);
  assert.match(style, /\.excelx-function-arguments-dialog__rows/);
});

test('shared provider exports the existing look and language catalogs and React hooks', async () => {
  const html = renderToStaticMarkup(createElement(VibeOfficeProvider, {
    services: { translate: (message) => message },
    children: createElement('span', null, 'ready'),
  }));
  assert.match(html, /ready/);
  const source = await import('@vibeoffice/core');
  assert.equal(typeof source.useVibeOffice, 'function');
  const bridge = readFileSync(new URL('../../commonx/src/react-bridge.tsx', import.meta.url), 'utf8');
  assert.match(bridge, /id: 'paper2016'/);
  assert.match(bridge, /id: 'zh-Hant'/);
});

test('shared appearance controls render the same look and language choices', () => {
  const html = renderToStaticMarkup(createElement(VibeOfficeProvider, {
    services: { translate: (message) => message },
    children: createElement(AppearanceControls),
  }));
  assert.match(html, /aria-label="Look"/);
  assert.match(html, /value="paper2016">Paper 2016/);
  assert.match(html, /aria-label="Language"/);
  assert.match(html, /value="zh-Hant">繁體中文/);
});

test('shared native dialog exposes an accessible title and request-close affordance', () => {
  const html = renderToStaticMarkup(createElement(OfficeDialog, { title: 'Credentials', onRequestClose() {}, children: createElement('button', null, 'Continue') }));
  assert.match(html, /<dialog[^>]*aria-modal="true"[^>]*aria-labelledby="[^"]+"/);
  assert.match(html, /<h2 id="[^"]+">Credentials<\/h2>/);
  assert.match(html, /Continue/);
});

test('shared native dialog keeps modeless semantics when requested', () => {
  const html = renderToStaticMarkup(createElement(OfficeDialog, { title: 'Modeless', modal: false, children: createElement('p', null, 'Content') }));
  assert.match(html, /<dialog[^>]*aria-labelledby="[^"]+"/);
  assert.doesNotMatch(html, /aria-modal=/);
  assert.match(readFileSync(new URL('../../commonx/src/dialog.tsx', import.meta.url), 'utf8'), /if \(modal\) dialog\.showModal\(\); else dialog\.show\(\);/);
});

test('React template metadata uses only Wordx, Excelx and PPTX routes and preserves generated blank previews', () => {
  const source = JSON.stringify({
    quire: [{ id: '', name: 'Blank Document', img: 'quire/templates/blank.jpg' }, { id: 'letter', name: 'Letter', img: 'quire/templates/letter.jpg' }],
    ledger: [{ id: '', name: 'Blank Workbook', img: 'ledger/templates/blank.jpg' }, { id: 'budget', name: 'Budget', img: 'ledger/templates/budget.jpg' }],
    lectern: [{ id: '', name: 'Blank Presentation', img: 'lectern/templates/blank.jpg' }, { id: 'meeting', name: 'Meeting', img: 'lectern/templates/meeting.jpg' }],
  });
  assert.deepEqual(JSON.parse(reactTemplates(source)), {
    wordx: [{ id: '', name: 'Blank Document', img: 'wordx/templates/blank.jpg' }, { id: 'letter', name: 'Letter', img: 'wordx/templates/letter.jpg' }],
    excelx: [{ id: '', name: 'Blank Workbook', img: 'excelx/templates/blank.jpg' }, { id: 'budget', name: 'Budget', img: 'excelx/templates/budget.jpg' }],
    pptx: [{ id: '', name: 'Blank Presentation', img: 'pptx/templates/blank.jpg' }, { id: 'meeting', name: 'Meeting', img: 'pptx/templates/meeting.jpg' }],
  });
});

test('local drafts use their own database and support recovery without touching Recent Files', async () => {
  const store = createIndexedDbDraftStore(new IDBFactory());
  const draft = { id: 'wordx:local-1', editor: 'wordx', name: 'Plan.docx', updatedAt: 100, snapshot: { text: 'unsaved' } };
  await store.put(draft);
  assert.deepEqual(await store.list('wordx'), [draft]);
  assert.deepEqual(await store.list('excelx'), []);
  await store.remove(draft.id);
  assert.deepEqual(await store.list(), []);
});

test('local open/save honors a writable file handle and returns the updated file', async () => {
  let written;
  const handle = {
    async getFile() { return new File([written ?? 'original'], 'Plan.docx'); },
    async createWritable() {
      return { async write(blob) { written = await blob.text(); }, async close() {} };
    },
  };
  const opened = await openLocalDocument(undefined, { showOpenFilePicker: async () => [handle] });
  assert.equal(await opened.file.text(), 'original');
  const saved = await saveLocalDocument(opened, new Blob(['updated']), 'Plan.docx', { document: { createElement() { throw new Error('download fallback should not run'); } } });
  assert.equal(written, 'updated');
  assert.equal(await saved.file.text(), 'updated');
  assert.equal(saved.handle, handle);
});

test('local file open rejects promptly when its mount AbortSignal is cancelled', async () => {
  const controller = new AbortController();
  let pickerSettled = false;
  let finishPicker;
  const picker = { showOpenFilePicker: () => new Promise((resolve) => { finishPicker = () => { pickerSettled = true; resolve([]); }; }) };
  const opening = openLocalDocument(undefined, picker, controller.signal);
  controller.abort();
  await assert.rejects(opening, { name: 'AbortError' });
  assert.equal(pickerSettled, false, 'the host stops waiting even if the browser-owned picker cannot be programmatically dismissed');
  finishPicker();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(pickerSettled, true, 'the eventual native picker completion is safely ignored');
});

test('local file-input fallback removes its picker when the mount is aborted', async () => {
  const previousDocument = globalThis.document;
  const listeners = new Map();
  const input = { type: '', accept: '', files: [], removed: false, addEventListener(type, listener) { listeners.set(type, listener); }, removeEventListener(type) { listeners.delete(type); }, click() {}, remove() { this.removed = true; } };
  globalThis.document = { createElement: (tag) => { assert.equal(tag, 'input'); return input; } };
  try {
    const controller = new AbortController();
    const opening = openLocalDocument(undefined, { document: globalThis.document }, controller.signal);
    assert.equal(input.type, 'file');
    controller.abort();
    await assert.rejects(opening, { name: 'AbortError' });
    assert.equal(input.removed, true);
    assert.equal(listeners.size, 0);
  } finally { globalThis.document = previousDocument; }
});

test('local file-input fallback is created in the supplied owner window', async () => {
  const file = new File(['owner window'], 'Owner.docx');
  const listeners = new Map();
  const input = { type: '', accept: '', files: [file], removed: false, addEventListener(type, listener) { listeners.set(type, listener); }, removeEventListener(type) { listeners.delete(type); }, click() { listeners.get('change')?.(); }, remove() { this.removed = true; } };
  let globalDocumentCalls = 0;
  const previousDocument = globalThis.document;
  globalThis.document = { createElement() { globalDocumentCalls++; throw new Error('must use the supplied owner window'); } };
  const ownerWindow = { document: { createElement(tag) { assert.equal(tag, 'input'); return input; } } };
  try {
    const opened = await openLocalDocument(undefined, ownerWindow);
    assert.equal(opened.file, file);
    assert.equal(input.removed, true);
    assert.equal(globalDocumentCalls, 0);
    assert.equal(listeners.size, 0);
  } finally { globalThis.document = previousDocument; }
});

test('download fallback uses the supplied owner window for DOM and object URLs', async () => {
  const previousWindow = globalThis.window;
  const downloads = [];
  const revoked = [];
  let fallbackWindowCalls = 0;
  let deferredRevoke;
  const anchor = { href: '', download: '', hidden: false, removed: false, click() { downloads.push({ href: this.href, name: this.download, inOwnerDocument: attached }); }, remove() { this.removed = true; attached = false; } };
  let attached = false;
  globalThis.window = { document: { createElement() { fallbackWindowCalls++; throw new Error('must use the supplied owner window'); } } };
  const ownerWindow = {
    document: { body: { append(node) { assert.equal(node, anchor); attached = true; } }, createElement(tag) { assert.equal(tag, 'a'); return anchor; } },
    URL: { createObjectURL(blob) { assert.equal(blob.type, 'application/octet-stream'); return 'blob:owner-window'; }, revokeObjectURL(url) { revoked.push(url); } },
    setTimeout(callback, delay) { assert.equal(delay, 0); deferredRevoke = callback; return 1; },
  };
  try {
    const saved = await saveLocalDocument({ file: new File([], 'Owner.bin') }, new Blob(['download'], { type: 'application/octet-stream' }), 'Owner.bin', ownerWindow);
    assert.deepEqual(downloads, [{ href: 'blob:owner-window', name: 'Owner.bin', inOwnerDocument: true }]);
    assert.equal(anchor.removed, true);
    assert.equal(fallbackWindowCalls, 0);
    assert.equal(await saved.file.text(), 'download');
    deferredRevoke();
    assert.deepEqual(revoked, ['blob:owner-window']);
  } finally { globalThis.window = previousWindow; }
});

test('local save picker and draft transactions abort with their editor mount', async () => {
  const pickerController = new AbortController();
  const picker = { showSaveFilePicker: () => new Promise(() => {}) };
  const saving = saveLocalDocument({ file: new File([], 'Plan.docx') }, new Blob(['save']), 'Plan.docx', picker, pickerController.signal);
  pickerController.abort();
  await assert.rejects(saving, { name: 'AbortError' });

  const drafts = createIndexedDbDraftStore(new IDBFactory());
  const draftController = new AbortController();
  const writing = drafts.put({ id: 'wordx:abort', editor: 'wordx', name: 'Plan.docx', updatedAt: 1, snapshot: new Blob(['draft']) }, draftController.signal);
  draftController.abort();
  await assert.rejects(writing, { name: 'AbortError' });
});

test('local save aborts an in-progress writable stream', async () => {
  const controller = new AbortController();
  let startWrite;
  let rejectWrite;
  let abortCalls = 0;
  const writeStarted = new Promise((resolve) => { startWrite = resolve; });
  const handle = {
    async createWritable() {
      return {
        write() {
          startWrite();
          return new Promise((_, reject) => { rejectWrite = reject; });
        },
        async abort() {
          abortCalls++;
          rejectWrite?.(new DOMException('Aborted', 'AbortError'));
        },
        async close() {},
      };
    },
  };
  const saving = saveLocalDocument({ file: new File([], 'Plan.docx'), handle }, new Blob(['save']), 'Plan.docx', {}, controller.signal);
  await writeStarted;
  controller.abort();
  await assert.rejects(saving, { name: 'AbortError' });
  assert.ok(abortCalls > 0, 'cancelling the mount aborts the active writable stream');
});

test('an aborted mount never starts the local download fallback', async () => {
  const previousWindow = globalThis.window;
  let anchorsCreated = 0;
  globalThis.window = { document: { createElement() { anchorsCreated++; throw new Error('download must not start after unmount'); } } };
  try {
    const controller = new AbortController();
    controller.abort();
    await assert.rejects(saveLocalDocument({ file: new File([], 'Plan.docx') }, new Blob(['save']), 'Plan.docx', {}, controller.signal), { name: 'AbortError' });
    assert.equal(anchorsCreated, 0);
  } finally { globalThis.window = previousWindow; }
});

test('dirty local documents preserve Save, Discard, and Cancel behavior', async () => {
  for (const [choice, expected, expectedCalls] of [
    ['save', true, ['save']],
    ['discard', true, ['discard']],
    ['cancel', false, []],
  ]) {
    const calls = [];
    let dirty = true;
    const allowed = await resolveUnsavedChanges({
      isDirty: () => dirty,
      choose: async () => choice,
      async save() { calls.push('save'); dirty = false; },
      async discard() { calls.push('discard'); dirty = false; },
    });
    assert.equal(allowed, expected);
    assert.deepEqual(calls, expectedCalls);
  }
  await assert.rejects(resolveUnsavedChanges({
    isDirty: () => true,
    choose: async () => 'save',
    async save() { throw new Error('save failed'); },
    async discard() {},
  }), /save failed/);
});

test('remote URL provider is read-only and rejects unconfigured origins', async () => {
  let requested;
  const provider = createUrlReadProvider(['https://files.example'], async (url) => {
    requested = String(url);
    return new Response('document', { headers: { ETag: '"v1"' } });
  });
  const result = await provider.read('https://files.example/a.docx');
  assert.equal(requested, 'https://files.example/a.docx');
  assert.equal(await result.blob.text(), 'document');
  assert.deepEqual(result.file.revision, { value: '"v1"', kind: 'etag' });
  await assert.rejects(provider.read('https://other.example/a.docx'), /not allowed/);
  await assert.rejects(provider.read('https://user:secret@files.example/a.docx'), /not allowed/);
  assert.equal(remoteConflict().code, 'revision-conflict');
});

test('WebDAV provider limits origin and path, keeps credentials callback-scoped, and uses conditional writes', async () => {
  const calls = [];
  let credentialCalls = 0;
  const provider = createWebDavProvider({
    baseUrl: 'https://dav.example/docs/', allowedOrigins: ['https://dav.example'],
    credentials: async () => { credentialCalls++; return { username: 'u', password: 'p' }; },
    fetcher: async (url, init) => {
      calls.push({ url: String(url), init });
      if (String(url).includes('/conflict')) return new Response('', { status: 412 });
      return new Response('', { status: init.method === 'PROPFIND' ? 207 : 201, headers: { ETag: '"v2"' } });
    },
  });
  await provider.read('/docs/a.docx');
  await provider.save('/docs/a.docx', new Blob(['saved']), { ifRevision: { value: '"v1"', kind: 'etag' } });
  assert.equal(calls[0].init.method, 'GET');
  assert.equal(calls[1].init.headers['If-Match'], '"v1"');
  assert.match(calls[0].init.headers.Authorization, /^Basic /);
  assert.equal(credentialCalls, 2);
  await assert.rejects(provider.read('https://evil.example/a.docx'), /escaped/);
  await assert.rejects(provider.read('/docs-escape/a.docx'), /escaped/);
  await assert.rejects(provider.read('https://user:secret@dav.example/docs/a.docx'), /escaped/);
  await assert.rejects(provider.save('/docs/conflict.docx', new Blob(['x']), { ifRevision: { value: '"stale"', kind: 'etag' } }), (error) => error.code === 'revision-conflict');
  assert.throws(() => createWebDavProvider({ baseUrl: 'https://evil.example/', allowedOrigins: ['https://dav.example'] }), /not allowed/);
  assert.throws(() => createWebDavProvider({ baseUrl: 'https://user:secret@dav.example/docs/', allowedOrigins: ['https://dav.example'] }), /not allowed/);
});

test('remote editor callbacks retain file revision and use conditional save or create', async () => {
  const calls = [];
  const provider = {
    async read(id, signal) { calls.push(['read', id, signal]); return { file: { id, name: 'a.docx', revision: { value: 'v1', kind: 'etag' } }, blob: new Blob(['opened']) }; },
    async save(id, blob, options) { calls.push(['save', id, options]); return { id, name: 'a.docx', revision: { value: 'v2', kind: 'etag' } }; },
    async create(parentId, name, _blob, signal) { calls.push(['create', parentId, name, signal]); return { id: 'created', name }; },
  };
  const abort = new AbortController();
  const callbacks = createRemoteFileCallbacks(provider,
    async (signal) => { calls.push(['chooseOpen', signal]); return { id: 'file-1', name: 'a.docx' }; },
    async (name, signal) => { calls.push(['chooseSave', name, signal]); return { parentId: 'folder', name: 'new.docx' }; });
  const opened = await callbacks.onOpenFile(abort.signal);
  assert.equal(await opened.file.text(), 'opened');
  const saved = await callbacks.onSaveFile('a.docx', new Blob(['updated']), opened, abort.signal);
  assert.equal(saved.remote.file.revision.value, 'v2');
  await callbacks.onSaveFile('new.docx', new Blob(['new']), opened, abort.signal, { saveAs: true });
  assert.deepEqual(calls.map(([op]) => op), ['chooseOpen', 'read', 'save', 'chooseSave', 'create']);
  assert.equal(calls[0][1], abort.signal);
  assert.equal(calls[1][2], abort.signal);
  assert.deepEqual(calls[2][2], { ifRevision: { value: 'v1', kind: 'etag' }, signal: abort.signal });
  assert.equal(calls[3][1], 'new.docx');
  assert.equal(calls[3][2], abort.signal);
  assert.equal(calls[4][3], abort.signal);
});

test('remote Save As can explicitly update a selected destination revision', async () => {
  const calls = [];
  const destinationRevision = { value: 'v9', kind: 'custom' };
  const provider = {
    async read(id) { return { file: { id, name: 'source.docx', revision: { value: 'v1', kind: 'etag' } }, blob: new Blob(['source']) }; },
    async save(id, _blob, options) { calls.push(['save', id, options]); return { id, name: 'existing.docx' }; },
  };
  const callbacks = createRemoteFileCallbacks(provider, async () => ({ id: 'source', name: 'source.docx' }), async () => ({ id: 'existing', name: 'existing.docx', revision: destinationRevision }));
  const opened = await callbacks.onOpenFile();
  const saved = await callbacks.onSaveFile('copy.docx', new Blob(['copy']), opened, undefined, { saveAs: true });
  assert.equal(saved.remote.file.id, 'existing');
  assert.deepEqual(calls, [['save', 'existing', { ifRevision: destinationRevision }]]);
});

test('host POST and signed-object providers plug into editor open, save and Save As callbacks', async () => {
  const postRequests = [];
  let serviceRevision = 4;
  const postTransport = async (path, blob, headers = {}) => {
    postRequests.push({ path, method: 'POST', blob, headers });
    serviceRevision++;
    return { id: path.split('/').pop() || 'created', name: path.split('/').pop() || 'created.docx', revision: { value: `r${serviceRevision}`, kind: 'custom' } };
  };
  const postProvider = {
    async list() { return [{ id: 'files/source.docx', name: 'source.docx', revision: { value: 'r4', kind: 'custom' } }]; },
    async read(id) { return { file: { id, name: 'source.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', revision: { value: 'r4', kind: 'custom' } }, blob: new Blob(['source']) }; },
    async save(id, blob, options) { return postTransport(`/api/documents/${id}`, blob, { 'If-Match': options?.ifRevision?.value || '' }); },
    async create(parentId, name, blob) { return postTransport(`/api/documents/${parentId}/${name}`, blob, { 'If-None-Match': '*' }); },
  };
  const postCallbacks = createRemoteFileCallbacks(postProvider,
    async () => ({ id: 'files/source.docx', name: 'source.docx', revision: { value: 'r4', kind: 'custom' } }),
    async (name) => ({ parentId: 'archive', name }));
  const postOpened = await postCallbacks.onOpenFile();
  const postSaved = await postCallbacks.onSaveFile('source.docx', new Blob(['edited']), postOpened);
  await postCallbacks.onSaveFile('copy.docx', new Blob(['copy']), postOpened, undefined, { saveAs: true });
  assert.equal(postSaved.remote.file.revision.value, 'r5');
  assert.deepEqual(postRequests.map(({ path, method, headers }) => ({ path, method, headers })), [
    { path: '/api/documents/files/source.docx', method: 'POST', headers: { 'If-Match': 'r4' } },
    { path: '/api/documents/archive/copy.docx', method: 'POST', headers: { 'If-None-Match': '*' } },
  ]);

  const signerCalls = [];
  const signedRequests = [];
  const sign = async (method, key, conditions = {}) => {
    signerCalls.push({ method, key, ...conditions });
    return { url: `https://objects.example/${key}?short-lived=${signerCalls.length}`, headers: method === 'PUT' ? { 'Content-Type': 'application/octet-stream', ...(conditions.ifRevision ? { 'If-Match': conditions.ifRevision.value } : {}), ...(conditions.ifNoneMatch ? { 'If-None-Match': conditions.ifNoneMatch } : {}) } : {} };
  };
  const signedFetch = async (url, init = {}) => {
    signedRequests.push({ url: String(url), method: init.method || 'GET', headers: init.headers || {}, body: init.body });
    return { ok: true, blob: async () => new Blob(['signed download']), revision: { value: 'v8', kind: 'version' } };
  };
  const signedProvider = {
    async list() { return [{ id: 'bucket/a.docx', name: 'a.docx', revision: { value: 'v7', kind: 'version' } }]; },
    async read(id) {
      const authorization = await sign('GET', id);
      const response = await signedFetch(authorization.url, { headers: authorization.headers });
      return { file: { id, name: 'a.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', revision: response.revision }, blob: await response.blob() };
    },
    async save(id, blob, options) {
      const authorization = await sign('PUT', id, { ifRevision: options?.ifRevision });
      const response = await signedFetch(authorization.url, { method: 'PUT', headers: authorization.headers, body: blob });
      return { id, name: 'a.docx', revision: response.revision || options?.ifRevision };
    },
    async create(parentId, name, blob) {
      const key = `${parentId}/${name}`;
      const authorization = await sign('PUT', key, { ifNoneMatch: '*' });
      await signedFetch(authorization.url, { method: 'PUT', headers: authorization.headers, body: blob });
      return { id: key, name, revision: { value: 'v9', kind: 'version' } };
    },
  };
  const signedCallbacks = createRemoteFileCallbacks(signedProvider,
    async () => ({ id: 'bucket/a.docx', name: 'a.docx', revision: { value: 'v7', kind: 'version' } }),
    async (name) => ({ parentId: 'bucket/archive', name }));
  const signedOpened = await signedCallbacks.onOpenFile();
  assert.equal(await signedOpened.file.text(), 'signed download');
  const signedSaved = await signedCallbacks.onSaveFile('a.docx', new Blob(['signed edit']), signedOpened);
  await signedCallbacks.onSaveFile('copy.docx', new Blob(['signed copy']), signedOpened, undefined, { saveAs: true });
  assert.deepEqual(signerCalls, [
    { method: 'GET', key: 'bucket/a.docx' },
    { method: 'PUT', key: 'bucket/a.docx', ifRevision: { value: 'v8', kind: 'version' } },
    { method: 'PUT', key: 'bucket/archive/copy.docx', ifNoneMatch: '*' },
  ]);
  assert.equal(signedSaved.remote.file.revision.value, 'v8');
  assert.ok(signedRequests.every(({ url, headers }) => url.startsWith('https://objects.example/') && url.includes('short-lived=') && !JSON.stringify(headers).includes('AWS')));
  assert.deepEqual(signedRequests.map(({ method }) => method), ['GET', 'PUT', 'PUT']);
  assert.deepEqual(signedRequests.slice(1).map(({ headers }) => ({ match: headers['If-Match'], createOnly: headers['If-None-Match'] })), [
    { match: 'v8', createOnly: undefined },
    { match: undefined, createOnly: '*' },
  ]);
});

test('remote editor callbacks stop after abort and pass cancellation into pickers and providers', async () => {
  const abort = new AbortController();
  const provider = { async read(_id, signal) { abort.abort(); assert.equal(signal, abort.signal); return { file: { id: 'file', name: 'a.docx' }, blob: new Blob(['late']) }; } };
  const callbacks = createRemoteFileCallbacks(provider, async (signal) => { assert.equal(signal, abort.signal); return { id: 'file', name: 'a.docx' }; });
  assert.equal(await callbacks.onOpenFile(abort.signal), null);
});

test('suite CSP only adds explicitly configured remote origins', () => {
  const plugin = reactCsp('https://dav.example,https://api.example');
  const html = plugin.transformIndexHtml('<head></head>');
  assert.match(html, /connect-src 'self' https:\/\/dav\.example https:\/\/api\.example/);
  assert.match(html, /script-src 'self'/);
  assert.match(html, /frame-src 'none'/, 'the React suite must prohibit nested browsing contexts');
  assert.throws(() => reactCsp('https://dav.example/path'), /Invalid remote origin/);
});
