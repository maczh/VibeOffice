import { installReactBrushService } from './brush.js';
import { installReactDialogService } from './react-dialog-service.js';
import { installReactTooltipService } from './tooltip-service.js';
import { installReactMenuService } from './menu-service.js';
import { installReactLegacyDialogService } from './legacy-dialog-service.js';
import { installReactButtonService } from './button-service.js';
import { installReactChoiceService } from './choice-service.js';
import { installReactSelectService } from './select-service.js';
import { installReactSpinService } from './spin-service.js';
import { installReactStructureService } from './structure-service.js';
import { installReactComboService } from './combo-service.js';
import { installReactToolbarService } from './toolbar-service.js';
import { installReactDragService } from './drag-service.js';
import { installReactUiMountService } from './ui-mount-service.js';
import { installReactUiRefreshService } from './ui-refresh-service.js';
import { installReactUiTextService } from './ui-text-service.js';
import { installReactCommandService } from './command-service.js';

export interface DirectEngineDocument {
  host: HTMLElement;
  signal: AbortSignal;
  fetch?: typeof fetch;
  directHost?: Record<string, unknown> & { editor?: 'wordx' | 'excelx' | 'pptx' };
  skipLegacyScripts?: readonly string[];
  beforeLegacyScript?: (runtime: Window, script: URL) => void | (() => void);
  shellOwnedByReact?: boolean;
}

interface EditorRuntimeManifest {
  styles: Array<{ href?: string; css?: string; media?: string }>;
  scripts: string[];
}

const loadedScripts = new WeakMap<Window, Map<string, Promise<void>>>();
const mountingWindows = new WeakSet<Window>();
const engineNamespaces = new WeakSet<object>();
const engineNamespacesByHost = new WeakMap<Window, object>();

function abortError(): DOMException {
  return new DOMException('The editor mount was cancelled', 'AbortError');
}

function assertSameOrigin(url: URL, base: URL): void {
  if (!['http:', 'https:'].includes(url.protocol) || url.origin !== base.origin) {
    throw new TypeError('The editor runtime must be served from the host origin');
  }
}

function loadScript(document: Document, url: URL, signal: AbortSignal): Promise<void> {
  const view = document.defaultView;
  if (!view) return Promise.reject(new Error('The editor host document has no window'));
  let cache = loadedScripts.get(view);
  if (!cache) { cache = new Map(); loadedScripts.set(view, cache); }
  const existing = cache.get(url.href);
  if (existing) return existing;

  const pending = new Promise<void>((resolve, reject) => {
    if (signal.aborted) { reject(abortError()); return; }
    const script = document.createElement('script');
    script.async = false;
    script.src = url.href;
    const finish = (error?: Error) => {
      script.removeEventListener('load', onLoad);
      script.removeEventListener('error', onError);
      signal.removeEventListener('abort', onAbort);
      script.remove();
      if (error) reject(error); else resolve();
    };
    const onLoad = () => finish();
    const onError = () => finish(new Error(`Could not load editor module ${url.pathname}`));
    const onAbort = () => finish(abortError());
    script.addEventListener('load', onLoad, { once: true });
    script.addEventListener('error', onError, { once: true });
    signal.addEventListener('abort', onAbort, { once: true });
    document.head.appendChild(script);
  });
  cache.set(url.href, pending);
  return pending.catch((error: unknown) => {
    cache?.delete(url.href);
    throw error;
  });
}

function splitSelectorList(value: string): string[] {
  const selectors: string[] = [];
  let start = 0, depth = 0, quote = '';
  for (let index = 0; index < value.length; index++) {
    const char = value[index];
    if (quote) { if (char === quote && value[index - 1] !== '\\') quote = ''; continue; }
    if (char === '"' || char === "'") quote = char;
    else if (char === '(' || char === '[') depth++;
    else if (char === ')' || char === ']') depth--;
    else if (char === ',' && depth === 0) { selectors.push(value.slice(start, index).trim()); start = index + 1; }
  }
  selectors.push(value.slice(start).trim());
  return selectors.filter(Boolean);
}

function scopeSelector(selector: string): string {
  const root = '.vibeoffice-legacy-root[data-vibeoffice-engine-root]';
  if (selector === '*') return `${root}, ${root} *`;
  if (/^\[data-(?:look|layout)(?:\s*=|\])/i.test(selector)) return `${root}${selector}`;
  if (/^:is\([^)]*\[data-look/i.test(selector)) return `${root}${selector}`;
  if (/^(?::root|html|body)(?=$|[\s>+~.#:[])/i.test(selector)) {
    return selector.replace(/(^|[\s>+~,(])(?:html|body|:root)(?=$|[\s>+~.#:[])/gi, (_match, before: string) => `${before}${root}`);
  }
  return `${root} ${selector}`;
}

function serializeScopedRules(rules: CSSRuleList): string {
  return Array.from(rules).map((rule) => {
    const styleRule = rule as CSSStyleRule;
    if (typeof styleRule.selectorText === 'string') {
      return `${splitSelectorList(styleRule.selectorText).map(scopeSelector).join(', ')} { ${styleRule.style.cssText} }`;
    }
    const cssText = rule.cssText;
    if (/^@(?:-\w+-)?keyframes\b/i.test(cssText)) return cssText;
    const grouping = rule as CSSGroupingRule;
    if (grouping.cssRules) {
      const brace = cssText.indexOf('{');
      if (brace >= 0) return `${cssText.slice(0, brace).trim()} { ${serializeScopedRules(grouping.cssRules)} }`;
    }
    return cssText;
  }).join('\n');
}

function scopedStylesheet(document: Document, source: string, sourceUrl: URL, media: string, root: HTMLElement): HTMLStyleElement {
  const view = document.defaultView;
  if (!view) throw new Error('The editor host document has no window');
  const absoluteUrls = source.replace(/url\(\s*(?:"([^"]*)"|'([^']*)'|([^)]*))\s*\)/gi, (_match, double: string, single: string, bare: string) => {
    const value = (double ?? single ?? bare ?? '').trim();
    if (!value || /^(?:data:|blob:|https?:|\/\/|#)/i.test(value)) return `url("${value}")`;
    return `url("${new URL(value, sourceUrl).href}")`;
  });
  const parser = document.createElement('style');
  parser.textContent = absoluteUrls;
  root.appendChild(parser);
  try {
    const rules = parser.sheet?.cssRules;
    if (!rules) throw new Error(`Could not parse editor stylesheet ${sourceUrl.pathname}`);
    const output = document.createElement('style');
    output.media = media || 'all';
    output.textContent = serializeScopedRules(rules);
    return output;
  } finally { parser.remove(); }
}

/** Load editor styles and ordered classic modules, using either the legacy shell or a package-owned React shell. */
export async function mountLegacyEngineDocument(url: string | URL, { host, signal, fetch: fetcher, directHost = {}, skipLegacyScripts = [], beforeLegacyScript, shellOwnedByReact = false }: DirectEngineDocument): Promise<() => void> {
  const document = host.ownerDocument;
  const view = document.defaultView;
  if (!view) throw new Error('The editor host document has no window');
  const base = new URL(view.location.href);
  const source = new URL(url, base);
  assertSameOrigin(source, base);
  if (signal.aborted) throw abortError();
  const request = fetcher ?? view.fetch.bind(view);
  if (mountingWindows.has(view)) throw new Error('Only one VibeOffice editor can be mounted in a host document at a time');
  mountingWindows.add(view);
  let reservationActive = true;
  const releaseReservation = () => {
    if (!reservationActive) return;
    reservationActive = false;
    mountingWindows.delete(view);
  };
  try {
  let scriptUrls: string[];
  let stylesheetEntries: Array<{ href?: string; css?: string; media?: string }>;
  let parsed: Document | undefined;
  if (shellOwnedByReact) {
    const manifestUrl = new URL(`${source.pathname.replace(/\/+$/, '')}/runtime.json`, source);
    const response = await request(manifestUrl, { signal });
    if (!response.ok) throw new Error(`Could not load editor runtime manifest (${response.status})`);
    const manifest = await response.json() as EditorRuntimeManifest;
    if (!Array.isArray(manifest.scripts) || !Array.isArray(manifest.styles)) throw new TypeError('The editor runtime manifest is invalid');
    scriptUrls = manifest.scripts;
    stylesheetEntries = manifest.styles;
  } else {
    const response = await request(source, { signal });
    if (!response.ok) throw new Error(`Could not load editor shell (${response.status})`);
    const markup = await response.text();
    if (signal.aborted) throw abortError();
    const Parser = view.DOMParser;
    parsed = new Parser().parseFromString(markup, 'text/html');
    scriptUrls = Array.from(parsed.querySelectorAll('script[src]'), (script) => script.getAttribute('src') || '');
    stylesheetEntries = Array.from(parsed.head.querySelectorAll('link[rel="stylesheet"][href], style'), (node) => ({
      ...(node instanceof view.HTMLLinkElement ? { href: node.getAttribute('href') || '' } : { css: node.textContent || '' }),
      media: node.getAttribute('media') || '',
    }));
  }
  if (signal.aborted) throw abortError();
  const styles: HTMLElement[] = [];
  const root = shellOwnedByReact
    ? host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root][data-vibeoffice-react-shell]')
    : document.createElement('div');
  if (!root) throw new Error('The React editor shell is missing its engine root');
  root.dataset.lookRoot = '';
  if (!shellOwnedByReact) {
    root.className = 'vibeoffice-legacy-root';
    root.dataset.vibeofficeEngineRoot = '';
  }
  const runtimeWindow = view as Window & {
    __VIBE_OFFICE_DIRECT_HOST__?: Record<string, unknown> & { root?: HTMLElement | null };
    L?: { mountDomRoot?: (element: HTMLElement) => () => void; APP_ID?: string; APP_NAME?: string; APP?: string; ui?: Record<string, unknown> };
    VO?: { t?: (value: string, ...args: unknown[]) => string; N?: (value: string) => string };
    console: Console;
  };
  const directRuntime = runtimeWindow.__VIBE_OFFICE_DIRECT_HOST__ || {};
  directRuntime.reactHost = true;
  directRuntime.preserveHostLook = true;
  if (directRuntime.root?.isConnected) throw new Error('Only one VibeOffice editor can be mounted in a host document at a time');
  const hostEngineDescriptor = Object.getOwnPropertyDescriptor(runtimeWindow, 'L');
  const hostEngineNamespace = runtimeWindow.L;
  const cachedEngineNamespace = engineNamespacesByHost.get(view);
  const preserveHostEngineNamespace = !!hostEngineDescriptor && !engineNamespaces.has(hostEngineNamespace as object);
  if (preserveHostEngineNamespace) {
    if (!hostEngineDescriptor.configurable && !('value' in hostEngineDescriptor && hostEngineDescriptor.writable)) {
      throw new Error('The embedding window.L namespace cannot be isolated for the VibeOffice editor');
    }
    const isolatedDescriptor = 'value' in hostEngineDescriptor
      ? { ...hostEngineDescriptor, value: cachedEngineNamespace }
      : { configurable: hostEngineDescriptor.configurable ?? false, enumerable: hostEngineDescriptor.enumerable ?? false, writable: true, value: cachedEngineNamespace };
    Object.defineProperty(runtimeWindow, 'L', isolatedDescriptor);
  }
  let restoreDomRoot: (() => void) | undefined;
  let restoreBrushService: (() => void) | undefined;
  const restoreReactServices: Array<() => void> = [];
  let hostEngineNamespaceActive = preserveHostEngineNamespace;
  let disposed = false;
  const cleanup = () => {
    if (disposed) return;
    disposed = true;
    const errors: unknown[] = [];
    const attempt = (action: () => void) => { try { action(); } catch (error) { errors.push(error); } };
    attempt(releaseReservation);
    attempt(() => {
      if (runtimeWindow.__VIBE_OFFICE_DIRECT_HOST__ === directRuntime && directRuntime.root === root) {
        directRuntime.root = null;
        for (const key of ['launch', 'open', 'save', 'saveAs', 'ready']) delete directRuntime[key];
      }
    });
    const restoreDomRootNow = restoreDomRoot;
    restoreDomRoot = undefined;
    if (restoreDomRootNow) attempt(restoreDomRootNow);
    const restoreBrushNow = restoreBrushService;
    restoreBrushService = undefined;
    if (restoreBrushNow) attempt(restoreBrushNow);
    for (const restore of restoreReactServices.splice(0).reverse()) attempt(restore);
    for (const style of styles.splice(0)) attempt(() => style.remove());
    if (!shellOwnedByReact) attempt(() => root.remove());
    if (hostEngineNamespaceActive) {
      hostEngineNamespaceActive = false;
      attempt(() => Object.defineProperty(runtimeWindow, 'L', hostEngineDescriptor!));
    }
    if (errors.length) {
      const cleanupError = new AggregateError(errors, 'One or more VibeOffice editor cleanup actions failed');
      try {
        if (typeof runtimeWindow.reportError === 'function') runtimeWindow.reportError(cleanupError);
        else runtimeWindow.console.error(cleanupError);
      } catch { /* reporting must not interrupt the remaining mount failure/disposal path */ }
    }
  };
  if (!shellOwnedByReact) host.append(root);

  try {
    for (const entry of stylesheetEntries) {
      const stylesheetUrl = new URL(entry.href || source.href, source);
      assertSameOrigin(stylesheetUrl, base);
      const css = entry.href
        ? await request(stylesheetUrl, { signal }).then(async (stylesheet) => {
          if (!stylesheet.ok) throw new Error(`Could not load editor stylesheet (${stylesheet.status})`);
          return stylesheet.text();
        })
        : entry.css || '';
      const scoped = scopedStylesheet(document, css, stylesheetUrl, entry.media || '', root);
      root.appendChild(scoped);
      styles.push(scoped);
    }
  } catch (error) { cleanup(); throw error; }
  parsed?.body.querySelectorAll('script').forEach((script) => script.remove());
  for (const element of shellOwnedByReact || !parsed ? [] : Array.from(parsed.body.children)) {
    const copy = element.cloneNode(true) as HTMLElement;
    for (const asset of [copy, ...Array.from(copy.querySelectorAll<HTMLElement>('[src], [href], [poster]'))]) {
      for (const attribute of ['src', 'href', 'poster']) {
        const value = asset.getAttribute(attribute);
        if (value && !value.startsWith('#') && !/^(data:|blob:|javascript:)/i.test(value)) {
          asset.setAttribute(attribute, new URL(value, source).href);
        }
      }
    }
    root.appendChild(copy);
  }

  Object.assign(directRuntime, directHost, { root, reactHost: true });
  const loadedScripts = Array.isArray(directRuntime.loadedLegacyScripts) ? directRuntime.loadedLegacyScripts as string[] : [];
  loadedScripts.length = 0;
  directRuntime.loadedLegacyScripts = loadedScripts;
  runtimeWindow.__VIBE_OFFICE_DIRECT_HOST__ = directRuntime;
  try {
    for (const script of scriptUrls) {
      if (signal.aborted) throw abortError();
      const scriptUrl = new URL(script, source);
      assertSameOrigin(scriptUrl, base);
      const disposeScriptService = beforeLegacyScript?.(runtimeWindow, scriptUrl);
      if (typeof disposeScriptService === 'function') restoreReactServices.push(disposeScriptService);
      if (preserveHostEngineNamespace && runtimeWindow.L && runtimeWindow.L !== hostEngineNamespace) {
        engineNamespaces.add(runtimeWindow.L);
        engineNamespacesByHost.set(view, runtimeWindow.L);
      }
      if (signal.aborted) throw abortError();
      const reactOwnsSharedUi = shellOwnedByReact && scriptUrl.pathname.endsWith('/common/ui.js');
      if (reactOwnsSharedUi) {
        if (!runtimeWindow.L) throw new Error('The shared engine namespace must load before the React UI services');
        const reactUi = runtimeWindow.L.ui || (runtimeWindow.L.ui = {});
        if (!reactUi.cmds || typeof reactUi.cmds !== 'object') reactUi.cmds = {};
        if (!reactUi.hooks) reactUi.hooks = { beforeExec() {}, refocus() {}, commitEdit() {}, ownsText() { return false; } };
        for (const name of ['dialogOpen', 'dialogKey', 'menuOpen', 'menuKey']) if (typeof reactUi[name] !== 'function') reactUi[name] = () => false;
        const originals = new Map<string, PropertyDescriptor | undefined>();
        for (const name of Object.keys(reactUi)) originals.set(name, Object.getOwnPropertyDescriptor(reactUi, name));
        [
          installReactCommandService,
          installReactDialogService,
          installReactTooltipService,
          installReactMenuService,
          installReactLegacyDialogService,
          installReactButtonService,
          installReactChoiceService,
          installReactSelectService,
          installReactSpinService,
          installReactStructureService,
          installReactComboService,
          installReactToolbarService,
          installReactDragService,
          installReactUiMountService,
          installReactUiRefreshService,
          installReactUiTextService,
        ].forEach((install) => { const restore = install(runtimeWindow); if (typeof restore === 'function') restoreReactServices.push(restore); });
        restoreBrushService = installReactBrushService(runtimeWindow);
        const installed = new Map<string, PropertyDescriptor>();
        for (const name of Object.keys(reactUi)) {
          const descriptor = Object.getOwnPropertyDescriptor(reactUi, name);
          if (descriptor && descriptor.value !== originals.get(name)?.value) installed.set(name, descriptor);
        }
        restoreReactServices.push(() => {
          for (const [name, descriptor] of installed) {
            if (Object.getOwnPropertyDescriptor(reactUi, name)?.value !== descriptor.value) continue;
            const original = originals.get(name);
            if (original) Object.defineProperty(reactUi, name, original);
            else delete reactUi[name];
          }
        });
      }
      if (reactOwnsSharedUi || skipLegacyScripts.some((path) => scriptUrl.pathname.endsWith(path)) || (shellOwnedByReact && scriptUrl.pathname.endsWith('/common/ribbon.js'))) continue;
      loadedScripts.push(scriptUrl.pathname);
      await loadScript(document, scriptUrl, signal);
      if (scriptUrl.pathname.endsWith('/common/core.js') && runtimeWindow.L && runtimeWindow.VO?.t) {
        Object.assign(runtimeWindow.L, { t: runtimeWindow.VO.t, N: runtimeWindow.VO.N });
      }
      if (scriptUrl.pathname.endsWith('/common/core.js') && runtimeWindow.L) {
        engineNamespaces.add(runtimeWindow.L);
        engineNamespacesByHost.set(view, runtimeWindow.L);
      }
      if (scriptUrl.pathname.endsWith('/common/core.js') && runtimeWindow.L?.mountDomRoot) {
        restoreDomRoot = runtimeWindow.L.mountDomRoot(root);
        const editor = directHost.editor;
        const names = { wordx: 'Wordx', excelx: 'Excelx', pptx: 'PPTX' };
        if (editor && Object.hasOwn(names, editor)) {
          runtimeWindow.L.APP_ID = editor;
          runtimeWindow.L.APP_NAME = names[editor];
          runtimeWindow.L.APP = names[editor];
        }
      }
    }
  } catch (error) {
    cleanup();
    throw error;
  }
  if (signal.aborted) { cleanup(); throw abortError(); }
  return cleanup;
  } catch (error) {
    releaseReservation();
    throw error;
  }
}
