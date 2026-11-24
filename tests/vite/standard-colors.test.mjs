import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { STANDARD_COLOR_NAMES, standardColorName } from '@vibeoffice/core';

test('React shared color labels cover the complete Office standard palette', () => {
  const engine = readFileSync(new URL('../../public/common/core.js', import.meta.url), 'utf8');
  const palette = engine.match(/C\.STANDARD\s*=\s*\[([\s\S]*?)\n\s*\];/)?.[1];
  assert.ok(palette, 'the engine exposes its 40-color standard palette');
  const colors = [...palette.matchAll(/'(#(?:[0-9A-F]{6}))'/gi)].map(([, color]) => color.toLowerCase());
  assert.equal(colors.length, 40);
  assert.equal(new Set(colors).size, 40, 'the engine palette has no duplicate swatches');
  assert.deepEqual(Object.keys(STANDARD_COLOR_NAMES).sort(), [...colors].sort(), 'React has exactly one translated name for every standard swatch');
  for (const color of colors) {
    const label = standardColorName(color);
    assert.notEqual(label, color.toUpperCase(), `${color} resolves to an Office color name`);
    assert.equal(standardColorName(color.toUpperCase()), label, `${color} lookup is case-insensitive`);
  }
});
