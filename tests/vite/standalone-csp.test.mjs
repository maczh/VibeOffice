import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { reactCsp } from '../../tools/vite/react-csp.mjs';

test('independent production editors install the same restrictive CSP as the suite', () => {
  const plugin = reactCsp('https://files.example');
  const transform = plugin.transformIndexHtml;
  assert.equal(typeof transform, 'function');

  for (const app of ['wordx', 'excelx', 'pptx']) {
    const config = readFileSync(new URL(`../../${app}/vite.app.config.ts`, import.meta.url), 'utf8');
    assert.match(config, /reactCsp\(\)/, `${app} standalone build must install its CSP plugin`);
    const html = transform('<!doctype html><html><head></head><body></body></html>');
    const policies = [...html.matchAll(/<meta http-equiv="Content-Security-Policy" content="([^"]+)">/g)];
    assert.equal(policies.length, 1, `${app} standalone page gets exactly one CSP`);
    assert.match(policies[0][1], /script-src 'self'/);
    assert.match(policies[0][1], /frame-src 'none'/);
    assert.match(policies[0][1], /connect-src 'self' https:\/\/files\.example/);
  }
});
