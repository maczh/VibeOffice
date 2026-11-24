import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const main = readFileSync(new URL('../../suite/src/main.tsx', import.meta.url), 'utf8');
const styles = readFileSync(new URL('../../suite/src/style.css', import.meta.url), 'utf8');
const looks = readFileSync(new URL('../../public/common/looks.css', import.meta.url), 'utf8');

test('Start Center chrome uses colors from the active look settings', () => {
  assert.doesNotMatch(main, /#[\da-f]{3,8}\b/i);
  assert.doesNotMatch(styles, /#[\da-f]{3,8}\b/i);
  for (const token of ['--p-accent', '--p-accent-soft', '--p-accent-mid', '--surface', '--field-border', '--chrome-line']) {
    assert.match(looks, new RegExp(`${token}:`), `${token} is supplied by the look stylesheet`);
  }
});
