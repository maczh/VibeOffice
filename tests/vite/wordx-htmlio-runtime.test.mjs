import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { transform } from 'esbuild';

const source = readFileSync(new URL('../../wordx/src/html-io.ts', import.meta.url), 'utf8');
const compiled = await transform(source, { loader: 'ts', format: 'esm', target: 'es2022' });
const moduleUrl = `data:text/javascript;base64,${Buffer.from(compiled.code).toString('base64')}`;
const { installWordxHtmlIo } = await import(moduleUrl);

test('Wordx HTML import uses the DOMParser from its supplied runtime window', async () => {
  const original = globalThis.DOMParser;
  let runtimeCalls = 0;
  let returnDocument = false;
  const body = {
    tagName: 'BODY', className: '', children: [], childNodes: [],
    getAttribute: () => null, querySelector: () => null, querySelectorAll: () => [],
  };
  class RuntimeParser {
    parseFromString() {
      runtimeCalls++;
      if (!returnDocument) throw new Error('runtime parser reached');
      return { body, documentElement: body, querySelector: () => null };
    }
  }
  class AmbientParser {
    parseFromString() { throw new Error('ambient parser reached'); }
  }
  globalThis.DOMParser = AmbientParser;
  try {
    const D = {
      doc: undefined,
      newDoc: () => ({ props: {}, styles: {}, main: { blocks: [] } }),
      para: () => ({ t: 'p', runs: [] }),
      plen: () => 0,
      normalize() {},
    };
    const runtime = { DOMParser: RuntimeParser, document: {}, L: { D, R: {}, clone: (value) => structuredClone(value) } };
    installWordxHtmlIo(runtime);
    await assert.rejects(runtime.L.htmlio.htmlToBlocks('<p>sample</p>'), /runtime parser reached/);
    assert.equal(runtimeCalls, 1);
    returnDocument = true;
    const parsed = await runtime.L.htmlio.docFromHTML('<p>sample</p>');
    assert.equal(parsed.main.blocks[0].t, 'p');
    assert.equal(runtimeCalls, 3, 'docFromHTML and its nested block parser both use the supplied runtime');
  } finally {
    if (original === undefined) delete globalThis.DOMParser;
    else globalThis.DOMParser = original;
  }
});
