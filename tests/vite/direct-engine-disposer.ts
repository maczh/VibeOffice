import { mountLegacyEngineDocument } from '../../commonx/src/direct-engine.js';

declare global {
  interface Window {
    mountDisposerFixture(host: HTMLElement): Promise<() => void>;
    disposerFixtureCalls: number;
  }
}

window.disposerFixtureCalls = 0;
window.mountDisposerFixture = (host) => mountLegacyEngineDocument('/fixture/index.html', {
  host,
  signal: new AbortController().signal,
  skipLegacyScripts: ['/fixture/skip.js'],
  beforeLegacyScript() {
    window.disposerFixtureCalls++;
    return () => { window.disposerFixtureCalls++; };
  },
  fetch: async (input) => {
    const path = new URL(String(input), location.href).pathname;
    if (path === '/fixture/styles.css') return new Response('.fixture-content{color:rgb(1, 2, 3)}', { status: 200 });
    return new Response('<!doctype html><html><head><link rel="stylesheet" href="/fixture/styles.css"></head><body><main class="fixture-content">Disposer fixture</main><script src="/fixture/skip.js"></script></body></html>', { status: 200 });
  },
});
