/* VExcel workbook model. Cells keyed by A1 address; styles deduped into a table. */

export interface CellStyle {
  font?: { name?: string; size?: number; bold?: boolean; italic?: boolean; underline?: boolean; color?: string }
  fill?: string
  border?: { top?: string; right?: string; bottom?: string; left?: string }
  align?: { h?: 'left' | 'center' | 'right'; v?: 'top' | 'middle' | 'bottom'; wrap?: boolean }
  numFmt?: string
}

export interface Cell {
  /** computed value */
  v?: string | number | boolean | null
  /** formula text without '=' */
  f?: string
  /** cached type: n number, s string, b bool, e error */
  t?: 'n' | 's' | 'b' | 'e'
  s?: number
}

export interface Merge { r1: number; c1: number; r2: number; c2: number }

export interface Sheet {
  name: string
  cells: Map<string, Cell>
  merges: Merge[]
  colWidths: Map<number, number>
  rowHeights: Map<number, number>
  defaultColWidth: number
  defaultRowHeight: number
  frozen?: { rows: number; cols: number }
}

export interface Workbook {
  sheets: Sheet[]
  styles: CellStyle[]
  active: number
}

export const colName = (c: number): string => {
  let s = ''
  c++
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26) }
  return s
}
export const colIndex = (s: string): number => {
  let n = 0
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}
export const addr = (r: number, c: number): string => colName(c) + (r + 1)
export function parseAddr(a: string): { r: number; c: number } | null {
  const m = /^([A-Za-z]+)(\d+)$/.exec(a)
  if (!m) return null
  return { r: +m[2] - 1, c: colIndex(m[1]) }
}

export const newSheet = (name: string): Sheet => ({
  name, cells: new Map(), merges: [], colWidths: new Map(), rowHeights: new Map(),
  defaultColWidth: 64, defaultRowHeight: 17,
})

export const newWorkbook = (): Workbook => ({ sheets: [newSheet('Sheet1')], styles: [{}], active: 0 })

export function styleIndex(wb: Workbook, s: CellStyle): number {
  const key = JSON.stringify(s)
  const i = wb.styles.findIndex((x) => JSON.stringify(x) === key)
  if (i >= 0) return i
  wb.styles.push(s)
  return wb.styles.length - 1
}

export function getCell(sheet: Sheet, r: number, c: number): Cell | undefined {
  return sheet.cells.get(addr(r, c))
}
export function setCell(sheet: Sheet, r: number, c: number, cell: Cell | undefined): void {
  const a = addr(r, c)
  if (!cell || (cell.v == null && !cell.f && cell.s == null)) sheet.cells.delete(a)
  else sheet.cells.set(a, cell)
}

/** display text of a cell (formula results already computed) */
export function cellText(cell: Cell | undefined, wb: Workbook, fmt: (v: Cell['v'], f?: string) => string): string {
  if (!cell) return ''
  if (cell.t === 'e') return String(cell.v)
  const style = cell.s != null ? wb.styles[cell.s] : undefined
  return fmt(cell.v, style?.numFmt)
}

export const cloneWorkbook = (wb: Workbook): Workbook => ({
  sheets: wb.sheets.map((s) => ({
    ...s,
    cells: new Map(s.cells),
    merges: s.merges.map((m) => ({ ...m })),
    colWidths: new Map(s.colWidths),
    rowHeights: new Map(s.rowHeights),
  })),
  styles: wb.styles.map((s) => ({ ...s })),
  active: wb.active,
})
