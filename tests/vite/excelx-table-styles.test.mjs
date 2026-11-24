import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

async function loadTableStyles() {
  const source = await readFile(path.join(root, 'excelx/src/table-styles.ts'), 'utf8');
  const output = ts.transpileModule(source, { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.ESNext } }).outputText;
  const url = `data:text/javascript;base64,${Buffer.from(output).toString('base64')}`;
  return import(url);
}

function runtime() {
  const model = {
    DEFAULT_THEME: { colors: ['#ffffff', '#000000', '#e7e6e6', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#0563c1', '#954f72', '#000000'] },
    rgb: (rgb) => ({ rgb }),
    applyTint: (hex, tint) => `${hex}@${tint}`,
  };
  const window = { L: { model } };
  return { window, model };
}

test('Excelx table-style service exposes all built-ins and resolves header, stripes, and custom DXFs', async () => {
  const { installExcelxTableStyles } = await loadTableStyles();
  const { window } = runtime();
  installExcelxTableStyles(window);
  const api = window.L.tstyle;
  assert.equal(api.NAMES.length, 60);
  assert.deepEqual([api.NAMES[0], api.NAMES[20], api.NAMES[21], api.NAMES[48], api.NAMES[49], api.NAMES[59]], [
    'TableStyleLight1', 'TableStyleLight21', 'TableStyleMedium1', 'TableStyleMedium28', 'TableStyleDark1', 'TableStyleDark11',
  ]);

  const workbook = {
    id: 'fixture',
    tables: [{}],
    dxfs: [{ fill: { fg: '#123456' }, font: { b: true }, border: { b: { style: 'thin', color: { rgb: '#000000' } } } }],
    tableAt: () => ({ ref: { r1: 0, r2: 3, c1: 0, c2: 1 }, style: { name: 'TableStyleLight1', rowStripes: true, colStripes: false }, header: true }),
  };
  const sheet = { wb: workbook };
  const header = api.at(sheet, 0, 0);
  assert.equal(header.font.b, true);
  assert.equal(header.border.b.style, 'thin');
  const stripe = api.at(sheet, 1, 0);
  assert.deepEqual(stripe.fill, { pattern: 'solid', fg: { rgb: '#D9D9D9' } });

  workbook.tableStyles = { custom: [{ name: 'Custom', elements: [{ type: 'wholeTable', dxf: 0 }] }] };
  workbook.tableAt = () => ({ ref: { r1: 0, r2: 3, c1: 0, c2: 1 }, style: { name: 'Custom' }, header: true });
  assert.deepEqual(api.at(sheet, 3, 0), {
    fill: { pattern: 'solid', fg: '#123456' },
    font: { b: true },
    border: { b: { style: 'thin', color: { rgb: '#000000' } } },
  });
  workbook.dxfs[0].fill.fg = '#654321';
  assert.equal(api.at(sheet, 3, 0).fill.fg, '#123456', 'the resolver reuses its cached spec until explicitly reset');
  api.reset();
  assert.equal(api.at(sheet, 3, 0).fill.fg, '#654321', 'reset allows recomputing cached style specs after workbook updates');
});
