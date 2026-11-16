import { mountLegacyEngineDocument } from '../../commonx/src/direct-engine.js';

declare global {
  interface Window {
    mountEngineFixture(host: HTMLElement): Promise<() => void>;
  }
}

window.mountEngineFixture = async (host) => mountLegacyEngineDocument('/fixture/index.html', {
  host,
  signal: new AbortController().signal,
  fetch: async () => new Response('<!doctype html><html><head><style>.fixture-content{color:rgb(1, 2, 3)}</style></head><body><main class="fixture-content">Mounted engine shell</main></body></html>', { status: 200 }),
});
