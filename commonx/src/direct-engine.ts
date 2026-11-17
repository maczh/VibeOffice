import { installReactBrushService } from './brush.js';

export interface DirectEngineDocument {
  host: HTMLElement;
  signal: AbortSignal;
  fetch?: typeof fetch;
  directHost?: Record<string, unknown> & { editor?: 'wordx' | 'excelx' | 'pptx' };
  skipLegacyScripts?: readonly string[];
  beforeLegacyScript?: (runtime: Window, script: URL) => void;
  shellOwnedByReact?: boolean;
}

interface EditorRuntimeManifest {
  styles: Array<{ href?: string; css?: string; media?: string }>;
  scripts: string[];
}

const loadedScripts = new WeakMap<Window, Map<string, Promise<void>>>();
const mountingWindows = new WeakSet<Window>();

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
  if (!shellOwnedByReact) {
    root.className = 'vibeoffice-legacy-root';
    root.dataset.vibeofficeEngineRoot = '';
  }
  const runtimeWindow = view as Window & {
    __VIBE_OFFICE_DIRECT_HOST__?: Record<string, unknown> & { root?: HTMLElement | null };
    L?: { mountDomRoot?: (element: HTMLElement) => () => void; APP_ID?: string; APP_NAME?: string; APP?: string };
  };
  const directRuntime = runtimeWindow.__VIBE_OFFICE_DIRECT_HOST__ || {};
  if (directRuntime.root?.isConnected) throw new Error('Only one VibeOffice editor can be mounted in a host document at a time');
  let restoreDomRoot: (() => void) | undefined;
  let restoreBrushService: (() => void) | undefined;
  const cleanup = () => {
    releaseReservation();
    if (runtimeWindow.__VIBE_OFFICE_DIRECT_HOST__ === directRuntime && directRuntime.root === root) {
      directRuntime.root = null;
      for (const key of ['launch', 'open', 'save', 'saveAs', 'ready']) delete directRuntime[key];
    }
    restoreDomRoot?.();
    restoreDomRoot = undefined;
    restoreBrushService?.();
    restoreBrushService = undefined;
    styles.forEach((style) => style.remove());
    if (!shellOwnedByReact) root.remove();
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

  Object.assign(directRuntime, directHost, { root });
  const loadedScripts = Array.isArray(directRuntime.loadedLegacyScripts) ? directRuntime.loadedLegacyScripts as string[] : [];
  loadedScripts.length = 0;
  directRuntime.loadedLegacyScripts = loadedScripts;
  runtimeWindow.__VIBE_OFFICE_DIRECT_HOST__ = directRuntime;
  try {
    for (const script of scriptUrls) {
      if (signal.aborted) throw abortError();
      const scriptUrl = new URL(script, source);
      assertSameOrigin(scriptUrl, base);
      beforeLegacyScript?.(runtimeWindow, scriptUrl);
      if (signal.aborted) throw abortError();
      if (skipLegacyScripts.some((path) => scriptUrl.pathname.endsWith(path))) continue;
      loadedScripts.push(scriptUrl.pathname);
      await loadScript(document, scriptUrl, signal);
      if (shellOwnedByReact && scriptUrl.pathname.endsWith('/common/ui.js')) restoreBrushService = installReactBrushService(runtimeWindow);
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
