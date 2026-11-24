import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const source = readFileSync(new URL('../../pptx/src/background-dialog.tsx', import.meta.url), 'utf8');

test('PPTX background dialog resolves runtime and animation frames from its mounted host', () => {
  assert.match(source, /usePptxRuntime<BackgroundRuntime>\(\)/);
  assert.match(source, /usePptxHost\(\)/);
  assert.match(source, /ownerDocument\.defaultView/);
  assert.doesNotMatch(source, /\bwindow\.(?:L|requestAnimationFrame)|(?<![\w.])cancelAnimationFrame\(/);
});

test('PPTX engine shell resolves presentation state from its rendered document', () => {
  const shell = readFileSync(new URL('../../pptx/src/engine-shell.tsx', import.meta.url), 'utf8');
  assert.match(shell, /ownerDocument\.defaultView/);
  assert.doesNotMatch(shell, /\bwindow\.L|presentationRuntime\(\)/);
});
