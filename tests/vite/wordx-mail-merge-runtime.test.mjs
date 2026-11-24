import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { transform } from 'esbuild';

const source = readFileSync(new URL('../../wordx/src/mail-merge.ts', import.meta.url), 'utf8');
const compiled = await transform(source, { loader: 'ts', format: 'esm', target: 'es2022' });
const moduleUrl = `data:text/javascript;base64,${Buffer.from(compiled.code).toString('base64')}`;
const { installWordxMailMerge } = await import(moduleUrl);

test('Wordx mail-merge HTML source uses DOMParser from its supplied runtime window', async () => {
  const original = globalThis.DOMParser;
  let runtimeCalls = 0;
  class RuntimeParser {
    parseFromString() {
      runtimeCalls++;
      return {
        querySelector: () => ({ rows: [{ cells: [{ textContent: 'Name' }] }, { cells: [{ textContent: 'Ada' }] }] }),
      };
    }
  }
  class AmbientParser {
    parseFromString() { throw new Error('ambient parser reached'); }
  }
  const runtime = {
    DOMParser: RuntimeParser,
    document: {},
    L: {
      D: {}, ed: {}, layout: {},
      panes: { task: {}, reactMailMergeState: null },
      h: () => ({}), ui: { msg() {} },
      N: (value) => value,
      pickFiles: async () => [{ name: 'recipients.html' }],
      readAsText: async () => '<table></table>',
      t: (value) => value,
    },
  };
  globalThis.DOMParser = AmbientParser;
  try {
    installWordxMailMerge(runtime);
    await runtime.L.mailmerge.openSource();
    assert.equal(runtimeCalls, 1);
    assert.deepEqual(runtime.L.mailmerge.reactRecipientsCatalog(), {
      name: 'recipients.html', fields: ['Name'], rows: [{ index: 0, values: ['Ada'], included: true }],
    });
  } finally {
    if (original === undefined) delete globalThis.DOMParser;
    else globalThis.DOMParser = original;
  }
});
