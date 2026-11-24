import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const tool = fileURLToPath(new URL('../../tools/wordx-pages.mjs', import.meta.url));

test('Wordx page-count comparison reads the renamed output field', () => {
  const dir = mkdtempSync(path.join(os.tmpdir(), 'wordx-pages-'));
  try {
    const before = path.join(dir, 'before.jsonl');
    const after = path.join(dir, 'after.jsonl');
    writeFileSync(before, [
      JSON.stringify({ file: 'away.docx', word: 4, wordx: 4 }),
      JSON.stringify({ file: 'toward.docx', word: 4, wordx: 5 }),
    ].join('\n'));
    writeFileSync(after, [
      JSON.stringify({ file: 'away.docx', word: 4, wordx: 5 }),
      JSON.stringify({ file: 'toward.docx', word: 4, wordx: 4 }),
    ].join('\n'));
    const output = execFileSync(process.execPath, [tool, '--compare', before, after], { encoding: 'utf8' });
    assert.match(output, /2 files counted both times/);
    assert.match(output, /1 toward Word's count or level, 1 away from it/);
    assert.match(output, /away: away\.docx  Word 4, Wordx 4 -> 5/);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
