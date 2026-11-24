import test from 'node:test';
import assert from 'node:assert/strict';
import { createWebDavProvider } from '@vibeoffice/core';

test('WebDAV passes cancellation into credential callbacks and skips transport after cancellation', async () => {
  const controller = new AbortController();
  let credentialSignal;
  let fetchCalls = 0;
  const provider = createWebDavProvider({
    baseUrl: 'https://dav.example/docs/',
    allowedOrigins: ['https://dav.example'],
    credentials: async (signal) => {
      credentialSignal = signal;
      controller.abort();
      return { username: 'user', password: 'secret' };
    },
    fetcher: async () => { fetchCalls++; return new Response('unexpected'); },
  });

  await assert.rejects(provider.read('/docs/document.docx', controller.signal), { name: 'AbortError' });
  assert.equal(credentialSignal, controller.signal);
  assert.equal(fetchCalls, 0, 'cancelled credential resolution must not start a network request');
});
