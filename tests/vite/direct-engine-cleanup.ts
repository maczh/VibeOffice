import { mountLegacyEngineDocument } from '../../commonx/src/direct-engine.js';

declare global {
  interface Window {
    mountCleanupFixture(host: HTMLElement): Promise<() => void>;
  }
}

window.mountCleanupFixture = (host) => mountLegacyEngineDocument('/fixture/index.html', {
  host,
  signal: new AbortController().signal,
  skipLegacyScripts: ['/fixture/skip.js'],
  beforeLegacyScript(runtime) { runtime.L = { engine: true }; },
  fetch: async (input) => {
    const path = new URL(String(input), location.href).pathname;
    if (path.endsWith('/fixture/styles.css')) return new Response('.fixture-content{color:rgb(1, 2, 3)}', { status: 200 });
    return new Response('<!doctype html><html><head><link rel="stylesheet" href="/fixture/styles.css"></head><body><main class="fixture-content">Cleanup fixture</main><script src="/fixture/skip.js"></script></body></html>', { status: 200 });
  },
});
