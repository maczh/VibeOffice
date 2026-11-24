import { createDirectEngineBridge } from '../../commonx/src/direct-host.js';
import { mountLegacyEngineDocument } from '../../commonx/src/direct-engine.js';

declare global {
  interface Window {
    mountEngineFixture(host: HTMLElement, options?: { signal?: AbortSignal; beforeLegacyScript?: (runtime: Window, script: URL) => void }): Promise<() => void>;
    mountBridgeFixture(host: HTMLElement, signal: AbortSignal, onStarted: () => void, onStopped: () => void, onConfigured?: () => void): Promise<() => void>;
    finishBridgeFixtureMount?: () => void;
  }
}

window.mountEngineFixture = async (host, options = {}) => mountLegacyEngineDocument('/fixture/index.html', {
  host,
  signal: options.signal ?? new AbortController().signal,
    beforeLegacyScript: options.beforeLegacyScript,
  skipLegacyScripts: ['/fixture/pre-xml.js', '/fixture/pre-extension.js'],
    fetch: async (input) => {
      const path = new URL(String(input), location.href).pathname;
      if (path.endsWith('/fixture/styles.css')) return new Response('.fixture-content{color:rgb(1, 2, 3)}', { status: 200 });
      return new Response('<!doctype html><html><head><link rel="stylesheet" href="/fixture/styles.css"></head><body><main class="fixture-content">Mounted engine shell</main><script src="/fixture/pre-xml.js"></script><script src="/fixture/pre-extension.js"></script></body></html>', { status: 200 });
    },
});

window.mountBridgeFixture = async (host, signal, onStarted, onStopped, onConfigured) => {
  const bridge = createDirectEngineBridge({
    editor: 'wordx', engine: { url: '/tests/vite/engine-bridge-fixture.html' }, accept: { 'text/plain': ['.txt'] },
    skipLegacyScripts: ['/@vite/client'],
    beforeLegacyScript(runtime) {
      const view = runtime as Window & { L?: { app?: { mount?: (root: HTMLElement) => Promise<() => void> } } };
      view.L = { app: { mount: (_root) => new Promise((resolve) => {
        window.finishBridgeFixtureMount = () => resolve(() => onStopped());
        onStarted();
      }) } };
    },
    configureRuntime: () => onConfigured?.(),
  });
  return bridge({ host, editor: 'wordx', signal });
};
