import test from 'node:test';
import assert from 'node:assert/strict';
import { createWebDavProvider } from '@vibeoffice/core';

test('WebDAV rejects encoded separators and traversal before making a request', async () => {
  const requested = [];
  const provider = createWebDavProvider({
    baseUrl: 'https://dav.example/docs/',
    allowedOrigins: ['https://dav.example'],
    fetcher: async (url) => { requested.push(String(url)); return new Response('file'); },
  });

  for (const path of ['/docs/%2fsecret', '/docs/%252fsecret', '/docs/%252e%252e/secret', '/docs/%255csecret']) {
    await assert.rejects(provider.read(path), /encoded separator or traversal/);
  }
  assert.deepEqual(requested, [], 'unsafe encoded paths are rejected before reaching fetch');

  const safe = await provider.read('/docs/annual%20summary.docx');
  assert.equal(await safe.blob.text(), 'file');
  assert.deepEqual(requested, ['https://dav.example/docs/annual%20summary.docx']);

  for (const path of ['/docs', '/docs/']) {
    await assert.rejects(provider.save(path, new Blob(['x'])), /save target must be a file/);
  }
  for (const name of ['', '.', '..', 'nested/file.docx', 'nested\\file.docx']) {
    await assert.rejects(provider.create('/docs/', name, new Blob(['x'])), /file name is invalid/);
  }
  const created = await provider.create('/docs/', 'annual summary.docx', new Blob(['new']));
  assert.equal(created.id, '/docs/annual%20summary.docx');
  assert.deepEqual(requested, [
    'https://dav.example/docs/annual%20summary.docx',
    'https://dav.example/docs/annual%20summary.docx',
  ]);
});
