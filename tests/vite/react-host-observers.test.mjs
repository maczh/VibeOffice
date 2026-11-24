import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const files = [
  'wordx/src/list-box.tsx',
  'wordx/src/menu-content.tsx',
  'wordx/src/task-content.tsx',
  'excelx/src/formula-bar.tsx',
  'pptx/src/notes.tsx',
  'pptx/src/status-bar.tsx',
  'pptx/src/show-controls.tsx',
  'pptx/src/print-preview.tsx',
  'pptx/src/slides-pane.tsx',
  'pptx/src/task-content.tsx',
];

test('React editor layers construct mutation observers in the mounted host document realm', () => {
  for (const file of files) {
    const content = readFileSync(new URL(`../../${file}`, import.meta.url), 'utf8');
    assert.match(content, /ownerDocument\.defaultView\?\.MutationObserver/, `${file} uses the host document realm`);
    assert.doesNotMatch(content, /new MutationObserver\(/, `${file} does not instantiate an ambient observer`);
  }
});
