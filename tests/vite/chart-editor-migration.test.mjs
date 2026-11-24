import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const read = (path) => readFileSync(new URL(path, import.meta.url), 'utf8');

test('shared React Chart editor ownership is in commonx and legacy chart service keeps only the classic fallback', () => {
  const legacyCharts = read('../../public/common/charts.js');
  const reactService = read('../../commonx/src/chart-editor-service.ts');
  const reactDialog = read('../../commonx/src/chart-editor-dialog.tsx');
  const frame = read('../../commonx/src/editor-frame.tsx');

  assert.doesNotMatch(legacyCharts, /CH\.reactEditor(?:Catalog|Action|ColorMenu)|reactDirectHost/, 'the shared classic chart file no longer contains React-host UI state or routing');
  assert.match(legacyCharts, /const sheet = h\('div', \{ class: 'datasheet' \}\)/, 'classic pages retain their existing DOM chart editor');
  for (const adapter of ['reactEditorCatalog', 'reactEditorAction', 'reactEditorColorMenu']) assert.match(reactService, new RegExp(`${adapter}\\s*\\(`), `commonx owns ${adapter}`);
  assert.match(reactDialog, /installReactChartEditor\(view\)/, 'the React dialog installs the adapter after runtime readiness');
  assert.match(frame, /<ChartEditorDialog host=\{engineHost\}/, 'the shared component binds lifecycle to the active editor host');
});
