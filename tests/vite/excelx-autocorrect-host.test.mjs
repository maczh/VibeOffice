import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const editor = readFileSync(new URL('../../excelx/src/editor.tsx', import.meta.url), 'utf8');
const dialog = readFileSync(new URL('../../excelx/src/auto-correct-dialog.tsx', import.meta.url), 'utf8');
const autoFormat = readFileSync(new URL('../../excelx/src/auto-format-dialog.tsx', import.meta.url), 'utf8');
const printDialog = readFileSync(new URL('../../excelx/src/print-dialog.tsx', import.meta.url), 'utf8');
const nameDialogs = [
  ['Apply Names', readFileSync(new URL('../../excelx/src/apply-names-dialog.tsx', import.meta.url), 'utf8')],
  ['Create Names', readFileSync(new URL('../../excelx/src/create-names-dialog.tsx', import.meta.url), 'utf8')],
  ['Paste Name', readFileSync(new URL('../../excelx/src/paste-name-dialog.tsx', import.meta.url), 'utf8')],
];
const hostedLayers = [
  ['Comment Editor', readFileSync(new URL('../../excelx/src/comment-editor.tsx', import.meta.url), 'utf8')],
  ['Hyperlink', readFileSync(new URL('../../excelx/src/hyperlink-dialog.tsx', import.meta.url), 'utf8')],
  ['Edit Links', readFileSync(new URL('../../excelx/src/links-dialog.tsx', import.meta.url), 'utf8')],
];
const popupLayers = [
  ['Chart Type menu', readFileSync(new URL('../../excelx/src/chart-type-menu.tsx', import.meta.url), 'utf8')],
  ['Border palette', readFileSync(new URL('../../excelx/src/border-palette.tsx', import.meta.url), 'utf8')],
];
const formattingLayers = [
  ['Conditional Formatting', readFileSync(new URL('../../excelx/src/conditional-format-dialog.tsx', import.meta.url), 'utf8')],
  ['Conditional Format Style', readFileSync(new URL('../../excelx/src/conditional-format-style-dialog.tsx', import.meta.url), 'utf8')],
];
const protectionDialogs = readFileSync(new URL('../../excelx/src/protection-dialog.tsx', import.meta.url), 'utf8');
const tableSortDialogs = [
  ['Create List', readFileSync(new URL('../../excelx/src/create-list-dialog.tsx', import.meta.url), 'utf8')],
  ['Sort', readFileSync(new URL('../../excelx/src/sort-dialog.tsx', import.meta.url), 'utf8')],
];
const calculationDialogs = [
  ['Goal Seek', readFileSync(new URL('../../excelx/src/goal-seek-dialog.tsx', import.meta.url), 'utf8')],
  ['Data Table', readFileSync(new URL('../../excelx/src/data-table-dialog.tsx', import.meta.url), 'utf8')],
];
const fillDialogs = [
  ['Fill Across', readFileSync(new URL('../../excelx/src/fill-across-dialog.tsx', import.meta.url), 'utf8')],
  ['Series', readFileSync(new URL('../../excelx/src/series-dialog.tsx', import.meta.url), 'utf8')],
];
const filterDialogs = [
  ['Advanced Filter', readFileSync(new URL('../../excelx/src/advanced-filter-dialog.tsx', import.meta.url), 'utf8')],
  ['AutoFilter', readFileSync(new URL('../../excelx/src/auto-filter-dialog.tsx', import.meta.url), 'utf8')],
];
const worksheetDialogs = [
  ['Insert Sheet', readFileSync(new URL('../../excelx/src/insert-sheet-dialog.tsx', import.meta.url), 'utf8')],
  ['Unhide Sheet', readFileSync(new URL('../../excelx/src/unhide-sheet-dialog.tsx', import.meta.url), 'utf8')],
  ['Size', readFileSync(new URL('../../excelx/src/size-dialog.tsx', import.meta.url), 'utf8')],
  ['Zoom', readFileSync(new URL('../../excelx/src/zoom-dialog.tsx', import.meta.url), 'utf8')],
];
const editorSurfaces = [
  ['Formula bar', readFileSync(new URL('../../excelx/src/formula-bar.tsx', import.meta.url), 'utf8')],
  ['Status bar', readFileSync(new URL('../../excelx/src/status-bar-interactions.tsx', import.meta.url), 'utf8')],
  ['Grid interactions', readFileSync(new URL('../../excelx/src/grid-interactions.tsx', import.meta.url), 'utf8')],
  ['File drop', readFileSync(new URL('../../excelx/src/file-drop.tsx', import.meta.url), 'utf8')],
];
const formulaReviewDialogs = [
  ['Defined Names', readFileSync(new URL('../../excelx/src/defined-names-dialog.tsx', import.meta.url), 'utf8')],
  ['Watch Window', readFileSync(new URL('../../excelx/src/watch-window.tsx', import.meta.url), 'utf8')],
  ['Evaluate Formula', readFileSync(new URL('../../excelx/src/evaluate-formula-dialog.tsx', import.meta.url), 'utf8')],
];
const navigationDialogs = [
  ['Go To', readFileSync(new URL('../../excelx/src/go-to-dialog.tsx', import.meta.url), 'utf8')],
  ['Go To Special', readFileSync(new URL('../../excelx/src/go-to-special-dialog.tsx', import.meta.url), 'utf8')],
  ['Shift Cells', readFileSync(new URL('../../excelx/src/shift-cells-dialog.tsx', import.meta.url), 'utf8')],
  ['Move or Copy Sheet', readFileSync(new URL('../../excelx/src/move-copy-sheet-dialog.tsx', import.meta.url), 'utf8')],
  ['Tab Color', readFileSync(new URL('../../excelx/src/tab-color-dialog.tsx', import.meta.url), 'utf8')],
];
const outlineDialogs = [
  ['Style', readFileSync(new URL('../../excelx/src/style-dialog.tsx', import.meta.url), 'utf8')],
  ['Subtotals', readFileSync(new URL('../../excelx/src/subtotals-dialog.tsx', import.meta.url), 'utf8')],
  ['Outline Settings', readFileSync(new URL('../../excelx/src/outline-settings-dialog.tsx', import.meta.url), 'utf8')],
  ['Symbol', readFileSync(new URL('../../excelx/src/symbol-dialog.tsx', import.meta.url), 'utf8')],
];
const contentDialogs = [
  ['Scenarios', readFileSync(new URL('../../excelx/src/scenarios-dialog.tsx', import.meta.url), 'utf8')],
  ['Data Validation', readFileSync(new URL('../../excelx/src/validation-dialog.tsx', import.meta.url), 'utf8')],
  ['Workbook Properties', readFileSync(new URL('../../excelx/src/properties-dialog.tsx', import.meta.url), 'utf8')],
  ['Custom Header Footer', readFileSync(new URL('../../excelx/src/custom-header-footer-dialog.tsx', import.meta.url), 'utf8')],
];
const dataDialogs = [
  ['Text to Columns', readFileSync(new URL('../../excelx/src/text-to-columns-dialog.tsx', import.meta.url), 'utf8')],
  ['Paste Special', readFileSync(new URL('../../excelx/src/paste-special-dialog.tsx', import.meta.url), 'utf8')],
  ['Error Checking', readFileSync(new URL('../../excelx/src/error-checking-dialog.tsx', import.meta.url), 'utf8')],
  ['Find and Replace', readFileSync(new URL('../../excelx/src/find-dialog.tsx', import.meta.url), 'utf8')],
];
const workbookSurfaces = [
  ['Sheet Tabs', readFileSync(new URL('../../excelx/src/sheet-tabs.tsx', import.meta.url), 'utf8')],
  ['Shape Menu', readFileSync(new URL('../../excelx/src/shape-menu.tsx', import.meta.url), 'utf8')],
  ['Task Content', readFileSync(new URL('../../excelx/src/task-content.tsx', import.meta.url), 'utf8')],
];
const advancedDataDialogs = [
  ['Consolidate', readFileSync(new URL('../../excelx/src/consolidate-dialog.tsx', import.meta.url), 'utf8')],
  ['Data Form', readFileSync(new URL('../../excelx/src/data-form-dialog.tsx', import.meta.url), 'utf8')],
  ['PivotTable', readFileSync(new URL('../../excelx/src/pivot-dialog.tsx', import.meta.url), 'utf8')],
];
const formulaDialogs = [
  ['Chart Wizard', readFileSync(new URL('../../excelx/src/chart-wizard-dialog.tsx', import.meta.url), 'utf8')],
  ['Function Arguments', readFileSync(new URL('../../excelx/src/function-arguments-dialog.tsx', import.meta.url), 'utf8')],
];
const drawingImportDialogs = [
  ['Drawing Format', readFileSync(new URL('../../excelx/src/drawing-format-dialog.tsx', import.meta.url), 'utf8')],
  ['Text Import', readFileSync(new URL('../../excelx/src/text-import-dialog.tsx', import.meta.url), 'utf8')],
];
const settingsDialogs = [
  ['Options', readFileSync(new URL('../../excelx/src/options-dialog.tsx', import.meta.url), 'utf8')],
  ['Page Setup', readFileSync(new URL('../../excelx/src/page-setup-dialog.tsx', import.meta.url), 'utf8')],
  ['Save As', readFileSync(new URL('../../excelx/src/save-as-dialog.tsx', import.meta.url), 'utf8')],
  ['Spelling', readFileSync(new URL('../../excelx/src/spell-dialog.tsx', import.meta.url), 'utf8')],
  ['Customize', readFileSync(new URL('../../excelx/src/customize-dialog.tsx', import.meta.url), 'utf8')],
  ['Insert Function', readFileSync(new URL('../../excelx/src/insert-function-dialog.tsx', import.meta.url), 'utf8')],
  ['Format Cells', readFileSync(new URL('../../excelx/src/format-cells-dialog.tsx', import.meta.url), 'utf8')],
];

test('Excelx AutoCorrect dialog resolves its engine API from the mounted editor host', () => {
  assert.match(editor, /<ExcelxHostProvider host=\{host\}>/);
  assert.match(editor, /<\/ExcelxHostProvider>/);
  assert.match(dialog, /useExcelxRuntime<Runtime>\(\)/);
  assert.doesNotMatch(dialog, /\bwindow\.L|function runtime\(\)/);
});

test('Excelx AutoFormat dialog resolves its engine API from the mounted editor host', () => {
  assert.match(autoFormat, /useExcelxRuntime<Runtime>\(\)/);
  assert.doesNotMatch(autoFormat, /\bwindow\.L|function runtime\(\)/);
});

test('Excelx Print dialog resolves its engine API from the mounted editor host', () => {
  assert.match(printDialog, /useExcelxRuntime<Runtime>\(\)/);
  assert.doesNotMatch(printDialog, /\bwindow\.L|function runtime\(\)/);
});

test('Excelx defined-name dialogs resolve their engine APIs from the mounted editor host', () => {
  for (const [name, source] of nameDialogs) {
    assert.match(source, /useExcelxRuntime<(?:\w+)?Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx comment and link layers use host-scoped runtime and DOM APIs', () => {
  for (const [name, source] of hostedLayers) {
    assert.match(source, /useExcelxRuntime<Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient engine globals`);
  }
  assert.match(hostedLayers[0][1], /const host = useExcelxHost\(\)/);
  assert.match(hostedLayers[0][1], /host\.ownerDocument\.defaultView/);
  assert.doesNotMatch(hostedLayers[0][1], /\bdocument\.|(?<![\w.])requestAnimationFrame\(|(?<![\w.])cancelAnimationFrame\(/);
});

test('Excelx popup layers scope runtime, dismissal, and focus to their rendered document', () => {
  for (const [name, source] of popupLayers) {
    assert.match(source, /useExcelxRuntime<Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)|\bdocument\./, `${name} avoids ambient globals`);
    assert.match(source, /ownerDocument\.addEventListener\('pointerdown'/, `${name} binds outside dismissal to its owner document`);
  }
  assert.match(popupLayers[1][1], /ownerDocument\.activeElement/);
});

test('Excelx conditional-format dialogs resolve engine APIs from the mounted editor host', () => {
  for (const [name, source] of formattingLayers) {
    assert.match(source, /useExcelxRuntime<Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx sheet and workbook protection dialogs resolve APIs from the mounted editor host', () => {
  assert.equal((protectionDialogs.match(/useExcelxRuntime</g) || []).length, 2);
  assert.doesNotMatch(protectionDialogs, /\bwindow\.L|function (?:workbook)?runtime\(\)/);
});

test('Excelx list and sort dialogs resolve engine APIs from the mounted editor host', () => {
  for (const [name, source] of tableSortDialogs) {
    assert.match(source, /useExcelxRuntime<(?:\w+)?Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx what-if analysis dialogs resolve engine APIs from the mounted editor host', () => {
  for (const [name, source] of calculationDialogs) {
    assert.match(source, /useExcelxRuntime<(?:\w+)?Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx fill and series dialogs resolve engine APIs from the mounted editor host', () => {
  for (const [name, source] of fillDialogs) {
    assert.match(source, /useExcelxRuntime<(?:\w+)?Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx filter dialogs resolve APIs and outside-dismissal from their mounted host', () => {
  for (const [name, source] of filterDialogs) {
    assert.match(source, /useExcelxRuntime<(?:\w+)?Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient engine globals`);
  }
  assert.match(filterDialogs[1][1], /ownerDocument\.addEventListener\('pointerdown'/);
  assert.doesNotMatch(filterDialogs[1][1], /\bdocument\./);
});

test('Excelx worksheet layout dialogs resolve engine APIs from the mounted editor host', () => {
  for (const [name, source] of worksheetDialogs) {
    assert.match(source, /useExcelxRuntime<\w+Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx editor surfaces resolve engine and browser APIs from their mounted host', () => {
  for (const [name, source] of editorSurfaces) {
    assert.match(source, /useExcelxRuntime<(?:\w+Runtime|\{[^\n]+\})>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient engine globals`);
  }
  assert.match(editorSurfaces[0][1], /ownerWindow\?\.requestAnimationFrame/);
  assert.match(editorSurfaces[1][1], /ownerDocument\.defaultView\?\.MutationObserver/);
  assert.match(editorSurfaces[2][1], /ownerDocument\.addEventListener\('pointerleave'/);
  assert.match(editorSurfaces[2][1], /ownerWindow\?\.addEventListener\('beforeunload'/);
  assert.match(editorSurfaces[3][1], /ownerDocument\.defaultView\?\.MutationObserver/);
});

test('Excelx formula-review dialogs resolve their engine APIs from the mounted editor host', () => {
  for (const [name, source] of formulaReviewDialogs) {
    assert.match(source, /useExcelxRuntime<Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx navigation and worksheet dialogs resolve their APIs from the mounted editor host', () => {
  for (const [name, source] of navigationDialogs) {
    assert.match(source, /useExcelxRuntime<\w+Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx outline and symbol dialogs resolve APIs from the mounted editor host', () => {
  for (const [name, source] of outlineDialogs) {
    assert.match(source, /useExcelxRuntime<(?:\w+)?Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx content dialogs resolve APIs and browser timers from the mounted editor host', () => {
  for (const [name, source] of contentDialogs) {
    assert.match(source, /useExcelxRuntime<(?:\w+)?Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
  assert.match(contentDialogs[3][1], /ownerWindow\?\.requestAnimationFrame/);
  assert.doesNotMatch(contentDialogs[3][1], /(?<![\w.])requestAnimationFrame\(|(?<![\w.])cancelAnimationFrame\(/);
});

test('Excelx data operation dialogs resolve APIs from the mounted editor host', () => {
  for (const [name, source] of dataDialogs) {
    assert.match(source, /useExcelxRuntime<(?:\w+)?Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx sheet and shape surfaces scope runtime and browser APIs to their host', () => {
  for (const [name, source] of workbookSurfaces) {
    assert.match(source, /useExcelxRuntime<\w+Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)|(?<![\w.])document\.|(?<![\w.])window\./, `${name} avoids ambient globals`);
  }
  assert.match(workbookSurfaces[0][1], /ownerDocument\.elementFromPoint/);
  assert.match(workbookSurfaces[0][1], /ownerWindow\?\.requestAnimationFrame/);
  assert.match(workbookSurfaces[1][1], /ownerDocument\.defaultView\?\.MutationObserver/);
});

test('Excelx advanced data dialogs resolve their APIs from the mounted editor host', () => {
  for (const [name, source] of advancedDataDialogs) {
    assert.match(source, /useExcelxRuntime<\w+Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx chart and formula dialogs resolve APIs from the mounted editor host', () => {
  for (const [name, source] of formulaDialogs) {
    assert.match(source, /useExcelxRuntime<Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});

test('Excelx drawing and text import dialogs resolve APIs from their mounted host', () => {
  for (const [name, source] of drawingImportDialogs) {
    assert.match(source, /useExcelxRuntime<Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
  assert.match(drawingImportDialogs[0][1], /host\.ownerDocument\.addEventListener/);
  assert.doesNotMatch(drawingImportDialogs[0][1], /(?<![\w.])document\./);
});

test('Excelx settings and Save As dialogs resolve APIs from their mounted host', () => {
  for (const [name, source] of settingsDialogs) {
    assert.match(source, /useExcelxRuntime<Runtime>\(\)/, `${name} uses the editor host`);
    assert.doesNotMatch(source, /\bwindow\.L|function runtime\(\)/, `${name} avoids ambient globals`);
  }
});
