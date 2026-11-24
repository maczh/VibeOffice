import test from 'node:test';
import assert from 'node:assert/strict';
import { createUrlReadProvider, createWebDavProvider } from '@vibeoffice/core';

test('built-in remote providers refuse automatic redirects outside the configured origin boundary', async () => {
  const requests = [];
  const fetcher = async (url, init) => {
    requests.push({ url: String(url), init });
    return new Response(null, { status: 302, headers: { Location: 'https://unconfigured.example/collect' } });
  };
  const urlProvider = createUrlReadProvider(['https://files.example'], fetcher);
  await assert.rejects(urlProvider.read('https://files.example/document.docx'), /URL open failed/);

  const webDavProvider = createWebDavProvider({
    baseUrl: 'https://dav.example/docs/',
    allowedOrigins: ['https://dav.example'],
    fetcher,
  });
  await assert.rejects(webDavProvider.read('/docs/document.docx'), /WebDAV GET failed/);

  assert.equal(requests.length, 2);
  assert.deepEqual(requests.map(({ init }) => init.redirect), ['error', 'error'], 'neither built-in provider permits fetch to follow a response to an unconfigured origin');
});
