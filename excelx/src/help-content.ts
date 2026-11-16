export interface ExcelxHelpData {
  topics: Array<[string, string]>;
  keys: Array<[string, string]>;
}

/** Localized Excelx help copy owned by the React package; L.t() translates at render time. */
export const EXCELX_HELP: ExcelxHelpData = {
  topics: [
    ['Entering data', 'Click a cell and type, then press Enter (moves down), Tab (moves right) or an arrow key. F2 edits the active cell; Alt+Enter starts a new line inside a cell. Ctrl+Enter fills the typed entry into every selected cell. Numbers, dates (3/14/2026, 14-Mar), times (9:30 AM), percentages (12%) and currency ($1,200) are recognised and formatted automatically. Start an entry with an apostrophe to keep it as text.'],
    ['Formulas', 'Start with = and use cell references (A1, $A$1, Sheet2!B3, named ranges, table references like Sales[Amount]). While typing a formula, click or drag on the sheet to insert references, and press F4 to cycle absolute and relative references. Ctrl+Shift+Enter enters an array formula; dynamic-array functions such as FILTER, SORT, UNIQUE and SEQUENCE spill their results. Insert ▸ Function (Shift+F3) lists more than 450 functions with argument help.'],
    ['AutoFill and series', 'Drag the fill handle (the square at the bottom-right of the selection) to copy cells or extend a series: 1, 2 → 3, 4…; Mon → Tue; Jan → Feb; Q1 → Q2; dates by day, weekday, month or year. Double-click the handle to fill down beside your data. Edit ▸ Fill ▸ Series gives full control; Ctrl+D and Ctrl+R fill down and right.'],
    ['Formatting cells', 'Format ▸ Cells (Ctrl+1) has Number, Alignment, Font, Border, Patterns and Protection tabs. Custom number formats use Excel codes such as #,##0.00;[Red](#,##0.00), 0%, d-mmm-yy, [h]:mm and "text"@. The Formatting toolbar sets fonts, alignment, Merge and Center, currency, percent, comma style, decimals, indents, borders, fill and font colours. The Format Painter copies formats.'],
    ['Rows, columns and sheets', 'Drag a header border to resize; double-click it to AutoFit. Right-click a header to insert, delete, hide or unhide. Window ▸ Freeze Panes keeps headings in view. Right-click a sheet tab to insert, delete, rename, move or copy sheets, or to set a tab colour; Ctrl+Page Up / Page Down switch sheets.'],
    ['Sorting and filtering', 'Data ▸ Sort sorts by up to three keys, ascending or descending, with or without a header row; the Sort buttons sort by the active column. Data ▸ Filter ▸ AutoFilter adds drop-down arrows to the header row: pick a value, (Top 10...), (Custom...), (Blanks) or (NonBlanks). Advanced Filter uses a criteria range and can copy unique records elsewhere. Data ▸ Subtotals inserts SUBTOTAL rows at each change in a column and outlines the list.'],
    ['Lists (tables)', 'Data ▸ List ▸ Create List (Ctrl+L) turns a range into a list with a header row, AutoFilter and an optional Total Row. Formulas can refer to it by name: =SUM(Table1[Amount]).'],
    ['Conditional formatting and validation', 'Format ▸ Conditional Formatting applies up to three conditions per range (cell value between, greater than…, or a formula), plus data bars, colour scales and icon sets read from newer workbooks. Data ▸ Validation restricts entries to whole numbers, decimals, lists, dates, times, text length or a custom formula, with input messages and error alerts.'],
    ['Charts and pictures', 'Insert ▸ Chart (F11) opens the Chart Wizard: column, bar, line, pie, XY scatter, area and doughnut charts, embedded in the sheet or on their own chart sheet. Charts follow their data as it changes. Insert ▸ Picture adds images and AutoShapes; drag to move, drag a handle to resize, double-click to format.'],
    ['Names, comments and hyperlinks', 'Type a name in the Name Box to name the selection, or use Insert ▸ Name ▸ Define (Ctrl+F3). Insert ▸ Comment (Shift+F2) attaches a note; cells with comments show a red triangle. Insert ▸ Hyperlink (Ctrl+K) links to web pages, e-mail addresses or places in the workbook.'],
    ['Formula auditing', 'Tools ▸ Formula Auditing traces precedents and dependents with arrows, evaluates a formula step by step and shows a Watch Window. Ctrl+` shows formulas instead of results. Error Checking walks through cells with errors.'],
    ['Printing and PDF', 'File ▸ Page Setup sets orientation, scaling (fit to N pages), paper size, margins, centering, headers and footers (&[Page], &[Pages], &[Date], &[File], &[Tab]), print area, rows to repeat and gridlines. View ▸ Page Break Preview shows the pages; Insert ▸ Page Break adds manual breaks. File ▸ Print and Print Preview produce a PDF you can print anywhere.'],
    ['Files and compatibility', 'Excelx reads and writes Office Open XML workbooks (.xlsx, macro-enabled .xlsm and templates), keeping formulas, styles, merged cells, comments, lists, conditional formats, validation, charts, pictures, defined names, page setup and macros. It also opens CSV and text files through the Text Import Wizard, opens password-protected workbooks and can save with a password (Save As ▸ Tools ▸ General Options). Binary .xls files must be converted to .xlsx first.'],
    ['Protection', 'Tools ▸ Protection ▸ Protect Sheet locks every cell whose Locked box (Format ▸ Cells ▸ Protection) is checked; choose what users may still do, and set an optional password. Protect Workbook locks the sheet structure.'],
    ['What-if analysis', 'Tools ▸ Goal Seek finds the input that makes a formula reach a target. Data ▸ Table builds one- and two-variable data tables. Tools ▸ Scenarios stores sets of input values to compare.'],
  ],
  keys: [
    ['Ctrl+N / O / S / P', 'New, Open, Save, Print'], ['F12', 'Save As'], ['Ctrl+Z / Ctrl+Y', 'Undo / Redo'], ['F4', 'Repeat; in a formula, cycle $ references'],
    ['Ctrl+X / C / V', 'Cut, Copy, Paste'], ['Ctrl+Alt+V', 'Paste Special'], ['Ctrl+F / H', 'Find, Replace'], ['F5 / Ctrl+G', 'Go To'],
    ['F2', 'Edit the active cell'], ['Esc', 'Cancel an entry'], ['Alt+Enter', 'New line in a cell'], ['Ctrl+Enter', 'Fill the selection with the entry'], ['Ctrl+Shift+Enter', 'Enter an array formula'],
    ['Ctrl+D / Ctrl+R', 'Fill down / right'], ['Ctrl+; / Ctrl+Shift+:', 'Insert the date / time'], ["Ctrl+' / Ctrl+Shift+\"", 'Copy formula / value from the cell above'], ['Alt+=', 'AutoSum'], ['Shift+F3', 'Insert Function'], ['Ctrl+A (in a formula)', 'Function Arguments'], ['F9', 'Calculate now'],
    ['Arrow keys', 'Move one cell'], ['Ctrl+Arrow', 'Move to the edge of the data region'], ['Home / Ctrl+Home', 'Start of row / A1'], ['Ctrl+End', 'Last used cell'], ['End, Arrow', 'End mode jump'], ['Page Up / Down', 'Move one screen'], ['Alt+Page Up / Down', 'Move one screen left / right'], ['Ctrl+Page Up / Down', 'Previous / next sheet'],
    ['Shift+Arrow', 'Extend the selection'], ['Ctrl+Shift+Arrow', 'Extend to the edge of the data'], ['Ctrl+Space / Shift+Space', 'Select column / row'], ['Ctrl+A', 'Select the current region, then the whole sheet'], ['Ctrl+Shift+*', 'Select the current region'], ['Ctrl+/', 'Select the current array'],
    ['Ctrl+1', 'Format Cells'], ['Ctrl+B / I / U / 5', 'Bold, Italic, Underline, Strikethrough'], ['Ctrl+Shift+~ $ % ^ # @ !', 'General, Currency, Percent, Scientific, Date, Time, Number formats'], ['Ctrl+Shift+& / _', 'Outline border / remove borders'],
    ['Ctrl+9 / Ctrl+Shift+(', 'Hide / unhide rows'], ['Ctrl+0 / Ctrl+Shift+)', 'Hide / unhide columns'], ['Ctrl+- / Ctrl+Shift++', 'Delete / insert cells'], ['Shift+F11', 'Insert worksheet'], ['F11', 'Chart'], ['Shift+F2', 'Insert or edit comment'], ['Ctrl+K', 'Hyperlink'], ['Ctrl+F3 / F3', 'Define / paste name'], ['Ctrl+L', 'Create List'],
    ['Alt+Down', 'Pick from drop-down list'], ['Ctrl+`', 'Show formulas'], ['F7', 'Spelling'], ['Shift+F10', 'Shortcut menu'], ['Ctrl+F1', 'Task pane'], ['F1', 'Help'], ['Ctrl+Scroll', 'Zoom'],
  ],
};

interface ExcelxHelpRuntime { L?: { reactModules?: Record<string, unknown> } }

/** Provide the React-owned help catalog before classic engine modules initialize. */
export function installExcelxHelpContent(runtime: Window): void {
  const L = (runtime as Window & ExcelxHelpRuntime).L;
  if (!L) throw new Error('Excelx engine runtime is not ready for its help catalog');
  L.reactModules ??= {};
  L.reactModules.excelxHelpContent = EXCELX_HELP;
}
