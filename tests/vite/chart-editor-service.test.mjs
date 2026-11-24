import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { transform } from 'esbuild';

const source = readFileSync(new URL('../../commonx/src/chart-editor-service.ts', import.meta.url), 'utf8');
const compiled = await transform(source, { loader: 'ts', format: 'esm', target: 'es2022' });
const serviceUrl = `data:text/javascript;base64,${Buffer.from(compiled.code).toString('base64')}`;
const { installReactChartEditor } = await import(serviceUrl);

function runtime() {
  const events = [];
  const original = {
    edit(shape, done) { done(shape); },
    sample() { return { v: 1, kind: 'bar', title: '', legend: 'r', gridY: true, labels: false, cats: ['A'], series: [{ name: 'S', vals: [1] }] }; },
    render() { return { setAttribute() {}, outerHTML: '<svg />' }; },
    kinds() { return [['bar', 'Bar']]; },
    syncV2() {}, setSeriesColor() {}, PALETTE: ['#000000'],
  };
  const chart = { ...original };
  const L = { chart, clone: (value) => structuredClone(value), bus: { emit(name) { events.push(name); } }, t: (value) => value };
  return { events, original, chart, window: { L, HTMLElement: class HTMLElement {} } };
}

test('React chart editor service restores its APIs and drops the active edit on dispose', () => {
  const view = runtime();
  let committed;
  const dispose = installReactChartEditor(view.window);
  assert.notEqual(view.chart.edit, view.original.edit);

  view.chart.edit({ chart: { v: 1, kind: 'bar', cats: ['A'], series: [{ name: 'S', vals: [1] }] } }, (chart) => { committed = chart; });
  assert.ok(view.chart.reactEditorCatalog());
  assert.equal(view.events.length, 1);
  dispose();

  assert.equal(committed, undefined, 'disposing an open editor must not commit its draft');
  assert.equal(view.events.length, 2, 'closing the active React editor notifies its subscribed view');
  for (const [name, method] of Object.entries(view.original)) assert.equal(view.chart[name], method, `${name} is restored`);
  for (const name of ['reactEditorCatalog', 'reactEditorAction', 'reactEditorColorMenu']) assert.equal(name in view.chart, false, `${name} is removed`);
  dispose();
  assert.equal(view.events.length, 2, 'disposal is idempotent');
});

test('React chart editor disposal preserves methods installed by a later owner', () => {
  const view = runtime();
  const dispose = installReactChartEditor(view.window);
  const later = () => 'later owner';
  view.chart.reactEditorAction = later;
  dispose();
  assert.equal(view.chart.reactEditorAction, later);
  assert.equal(view.chart.edit, view.original.edit);
});
