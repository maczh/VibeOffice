import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const source = (path) => readFileSync(new URL(path, import.meta.url), 'utf8');
const hostBoundWordxLayers = [
  ['autocomplete tip', source('../../wordx/src/autocomplete-tip.tsx')],
  ['list box', source('../../wordx/src/list-box.tsx')],
  ['menu content', source('../../wordx/src/menu-content.tsx')],
  ['task content', source('../../wordx/src/task-content.tsx')],
];

test('Wordx React layers resolve engine services and DOM ownership from their editor host', () => {
  for (const [name, content] of hostBoundWordxLayers) {
    assert.doesNotMatch(content, /\bwindow\.(?:L|VO)\b|function runtime\(\)/, `${name} must not read ambient engine globals`);
    assert.match(content, /ownerDocument\.defaultView/, `${name} resolves its runtime from the mounted document`);
  }
  assert.match(hostBoundWordxLayers[3][1], /createContext<HTMLElement \| null>/, 'portaled task UI receives an explicit editor-host context');
  assert.match(hostBoundWordxLayers[3][1], /WordxTaskHostContext\.Provider value=\{host\}/, 'the exported layer supplies the host to its React descendants');
  assert.doesNotMatch(hostBoundWordxLayers[3][1], /\bwindow\.(?:addEventListener|removeEventListener|setTimeout|clearTimeout|requestAnimationFrame|cancelAnimationFrame)\b|\bdocument\.(?:activeElement|documentElement|getElementById|createElement|addEventListener|removeEventListener)\b/, 'task UI lifecycle and focus operations stay inside the mounted host document');
  assert.match(hostBoundWordxLayers[3][1], /const doc = host\.ownerDocument/);
  assert.match(hostBoundWordxLayers[3][1], /doc\.addEventListener\('selectionchange'/, 'selection changes are observed by the owning document');
});

test('Wordx Find creates highlight ranges in the editor host document realm', () => {
  const find = source('../../wordx/src/find.ts');
  assert.match(find, /new runtime\.Range\(\)/, 'highlight ranges must be constructed by the mounted editor window');
  assert.doesNotMatch(find, /new Range\(\)/, 'ambient Range can belong to a different document in embedded hosts');
});
