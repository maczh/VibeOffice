import { mountLegacyEngineDocument } from '../../commonx/src/direct-engine.js';

declare global {
  interface Window {
    mountEngineFixture(host: HTMLElement, options?: { signal?: AbortSignal; beforeLegacyScript?: (runtime: Window, script: URL) => void }): Promise<() => void>;
  }
}

window.mountEngineFixture = async (host, options = {}) => mountLegacyEngineDocument('/fixture/index.html', {
  host,
  signal: options.signal ?? new AbortController().signal,
  beforeLegacyScript: options.beforeLegacyScript,
  skipLegacyScripts: ['/fixture/pre-xml.js', '/fixture/pre-extension.js'],
  fetch: async (input) => new URL(String(input), location.href).pathname.endsWith('/fixture/styles.css')
    ? new Response('.fixture-content{color:rgb(1, 2, 3)}', { status: 200 })
    : new Response('<!doctype html><html><head><link rel="stylesheet" href="/fixture/styles.css"></head><body><main class="fixture-content">Mounted engine shell</main><script src="/fixture/pre-xml.js"></script><script src="/fixture/pre-extension.js"></script></body></html>', { status: 200 }),
});
