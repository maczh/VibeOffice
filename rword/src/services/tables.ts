import { emptyPara, uid, type Doc, type Table } from '../engine/model'
export function createTable(rows: number, columns: number): Table { const r = Math.max(1, Math.floor(rows)), c = Math.max(1, Math.floor(columns)); return { kind: 'tbl', id: uid('tbl'), widths: Array(c).fill(96), rows: Array.from({ length: r }, () => Array.from({ length: c }, () => ({ paras: [emptyPara()] }))) } }
export function tableCell(table: Table, row: number, column: number) { return table.rows[row]?.[column] }
export function addTableRow(table: Table): Table { const columns = table.widths.length || table.rows[0]?.length || 1; table.rows.push(Array.from({ length: columns }, () => ({ paras: [emptyPara()] }))); return table }
export function addTableColumn(table: Table): Table { table.widths.push(96); for (const row of table.rows) row.push({ paras: [emptyPara()] }); return table }
export function removeTableRow(table: Table, row: number): Table { if (table.rows.length > 1) table.rows.splice(row, 1); return table }
export function removeTableColumn(table: Table, column: number): Table { if (table.widths.length > 1) { table.widths.splice(column, 1); for (const row of table.rows) row.splice(column, 1) } return table }
export function insertTable(doc: Doc, rows: number, columns: number): Table { const table = createTable(rows, columns); const index = doc.blocks.length && doc.blocks[doc.blocks.length - 1].kind === 'p' ? doc.blocks.length - 1 : doc.blocks.length; doc.blocks.splice(index, 0, table); return table }
