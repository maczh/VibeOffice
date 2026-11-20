/* Real implementations for the Quire commands that act on the document model.
 * Every function is pure with respect to the React tree: the caller commits the
 * mutated Doc and refreshes the editable surface. */
import {
  type Block, type Cell, type Doc, type Para, type RPr, type Run, type Table,
  blockText, emptyPara, isPara, isTable, paraText, uid,
} from '../engine/model'
import { toggleCase } from './textOps'
import { autoFormatText } from './autocorrect'

export interface Location {
  blockIndex: number
  para: Para | null
  table: { blockIndex: number; row: number; col: number; cell: Cell; table: Table } | null
}

export interface CommandEnv {
  doc: Doc
  loc: Location | null
  selection: string
  arg?: string
}

export type CommandFn = (env: CommandEnv) => string | undefined

export function locate(doc: Doc, id: string | null | undefined): Location | null {
  if (!id) return null
  for (let i = 0; i < doc.blocks.length; i++) {
    const block = doc.blocks[i]
    if (isTable(block)) {
      if (block.id === id) {
        const cell = block.rows[0]?.[0] ?? { paras: [emptyPara()] }
        return { blockIndex: i, para: cell.paras[0] ?? null, table: { blockIndex: i, row: 0, col: 0, cell, table: block } }
      }
      for (let r = 0; r < block.rows.length; r++) {
        for (let c = 0; c < block.rows[r].length; c++) {
          const cell = block.rows[r][c]
          const para = cell.paras.find((p) => p.id === id)
          if (para) return { blockIndex: i, para, table: { blockIndex: i, row: r, col: c, cell, table: block } }
        }
      }
    } else if (block.id === id) {
      return { blockIndex: i, para: block, table: null }
    }
  }
  return null
}

export function fallbackLocation(doc: Doc): Location {
  for (let i = doc.blocks.length - 1; i >= 0; i--) {
    const block = doc.blocks[i]
    if (isPara(block)) return { blockIndex: i, para: block, table: null }
    const cell = block.rows[0]?.[0]
    if (cell?.paras[0]) return { blockIndex: i, para: cell.paras[0], table: { blockIndex: i, row: 0, col: 0, cell, table: block } }
  }
  const para = emptyPara()
  doc.blocks.push(para)
  return { blockIndex: doc.blocks.length - 1, para, table: null }
}

const needTable = (env: CommandEnv): { blockIndex: number; table: Table; row: number; col: number; cell: Cell } | null => {
  if (env.loc?.table) return env.loc.table
  const block = env.doc.blocks[env.loc?.blockIndex ?? -1]
  return isTable(block) ? { blockIndex: env.loc?.blockIndex ?? -1, table: block, row: 0, col: 0, cell: block.rows[0]?.[0] ?? { paras: [emptyPara()] } } : null
}

const blankCell = (): Cell => ({ paras: [emptyPara()] })

/* ------------------------------------------------------------------ tables */

function insertRow(env: CommandEnv, before: boolean): string | undefined {
  const at = needTable(env)
  if (!at) return 'Place the cursor in a table first.'
  const { table, row } = at
  const count = table.widths.length || table.rows[0]?.length || 1
  table.rows.splice(before ? row : row + 1, 0, Array.from({ length: count }, blankCell))
  return undefined
}

function insertColumn(env: CommandEnv, before: boolean): string | undefined {
  const at = needTable(env)
  if (!at) return 'Place the cursor in a table first.'
  const { table, col } = at
  table.widths.splice(before ? col : col + 1, 0, 96)
  for (const row of table.rows) row.splice(before ? col : col + 1, 0, blankCell())
  return undefined
}

function deleteCells(env: CommandEnv): string | undefined {
  const at = needTable(env)
  if (!at) return 'Place the cursor in a table first.'
  const { table, row, col } = at
  if (table.widths.length > 1) {
    table.widths.splice(col, 1)
    for (const r of table.rows) r.splice(col, 1)
  } else if (table.rows.length > 1) {
    table.rows.splice(row, 1)
  }
  return undefined
}

function tableCellText(cell: Cell): string {
  return cell.paras.map(paraText).join(' ').trim()
}

function numeric(text: string): number | undefined {
  const cleaned = text.replace(/[^0-9.\-]/g, '')
  if (!cleaned || cleaned === '-' || cleaned === '.') return undefined
  const value = Number(cleaned)
  return Number.isFinite(value) ? value : undefined
}

export const tableCommands: Record<string, CommandFn> = {
  rowAbove: (env) => insertRow(env, true),
  rowBelow: (env) => insertRow(env, false),
  colLeft: (env) => insertColumn(env, true),
  colRight: (env) => insertColumn(env, false),
  insertCells: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    const count = Math.max(1, Math.min(20, Number(env.arg) || 1))
    for (let i = 0; i < count; i++) {
      at.table.widths.splice(at.col + 1, 0, 96)
      for (const row of at.table.rows) row.splice(at.col + 1, 0, blankCell())
    }
    return undefined
  },
  deleteRows: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    if (at.table.rows.length > 1) at.table.rows.splice(at.row, 1)
    else at.table.rows[at.row] = [blankCell()]
    return undefined
  },
  deleteCols: deleteCells,
  deleteCells,
  deleteTable: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    env.doc.blocks.splice(at.blockIndex, 1)
    if (!env.doc.blocks.length) env.doc.blocks.push(emptyPara())
    return undefined
  },
  mergeCells: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    const row = at.table.rows[at.row]
    const first = row[at.col]
    if (!first) return undefined
    const span = Math.max(1, Math.min(20, Number(env.arg) || row.length - at.col))
    let width = first.width || at.table.widths[at.col] || 96
    for (let i = 1; i < span && at.col + i < row.length; i++) {
      const next = row[at.col + i]
      width += next.width || at.table.widths[at.col + i] || 96
      next.borders = undefined
    }
    row.splice(at.col + 1, Math.max(0, span - 1))
    at.table.widths.splice(at.col + 1, Math.max(0, span - 1))
    first.gridSpan = span
    first.width = width
    at.table.widths[at.col] = width
    return undefined
  },
  splitCells: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    const row = at.table.rows[at.row]
    const cell = row[at.col]
    const parts = Math.max(2, Math.min(20, Number(env.arg) || 2))
    const width = (cell.width || at.table.widths[at.col] || 96) / parts
    cell.gridSpan = undefined
    cell.width = width
    at.table.widths[at.col] = width
    const added: Cell[] = []
    for (let i = 1; i < parts; i++) { added.push({ paras: [emptyPara()], width }); at.table.widths.splice(at.col + i, 0, width) }
    row.splice(at.col + 1, 0, ...added)
    return undefined
  },
  splitTable: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    const { table, row } = at
    if (row >= table.rows.length - 1) return 'The cursor must be above the last row.'
    const tail = table.rows.splice(row + 1)
    const widthSum = tail[0]?.map((c) => c.width || 96) || table.widths
    const second: Table = { kind: 'tbl', id: uid('tbl'), rows: tail, widths: [...widthSum] }
    env.doc.blocks.splice(at.blockIndex + 1, 0, second, emptyPara())
    return undefined
  },
  sortAsc: (env) => sortTable(env, true),
  sortDesc: (env) => sortTable(env, false),
  sortDlg: (env) => sortTable(env, env.arg !== 'desc'),
  autoSum: (env) => formula(env, 'sum'),
  formulaDlg: (env) => formula(env, env.arg || 'sum'),
  textToTable: (env) => {
    const start = env.loc?.blockIndex ?? 0
    const paras = env.doc.blocks.slice(start).filter(isPara)
    if (paras.length < 1) return 'Select one or more paragraphs first.'
    const cols = Math.max(...paras.map((p) => p.runs.filter((r) => r.tab).length + 1))
    const rows = paras.map((p) => {
      const cells: Cell[] = []
      let cell: Para = emptyPara()
      for (const run of p.runs) {
        if (run.tab) { cells.push({ paras: [cell] }); cell = emptyPara() }
        else cell.runs.push({ ...run })
      }
      cells.push({ paras: [cell] })
      while (cells.length < cols) cells.push(blankCell())
      return cells
    })
    const table: Table = { kind: 'tbl', id: uid('tbl'), rows, widths: Array(cols).fill(96) }
    env.doc.blocks.splice(start, paras.length, table)
    return undefined
  },
  tableToText: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    const paras: Para[] = at.table.rows.map((row) => {
      const p = emptyPara()
      row.forEach((cell, i) => {
        if (i) p.runs.push({ tab: true })
        for (const run of cell.paras.flatMap((cp) => cp.runs)) p.runs.push({ ...run })
      })
      return p
    })
    env.doc.blocks.splice(at.blockIndex, 1, ...paras)
    return undefined
  },
  tableAutoFormat: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    at.table.headerRows = 1
    for (const cell of at.table.rows[0] || []) cell.fill = 'D9E2F3'
    return undefined
  },
  headingRows: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    at.table.headerRows = at.table.headerRows ? undefined : 1
    return undefined
  },
  autofitContents: (env) => setFit(env, 'content'),
  autofitWindow: (env) => setFit(env, 'window'),
  fixedWidth: (env) => setFit(env, 'fixed'),
  distRows: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    const height = Math.round(720 / Math.max(1, at.table.rows.length))
    for (const row of at.table.rows) for (const cell of row) cell.paras[0] = { ...(cell.paras[0] || emptyPara()), pPr: { ...cell.paras[0]?.pPr, line: height / 2 } }
    return undefined
  },
  distCols: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    const width = 576 / Math.max(1, at.table.widths.length)
    at.table.widths = at.table.widths.map(() => width)
    for (const row of at.table.rows) row.forEach((cell, i) => { cell.width = width; at.table.widths[i] = width })
    return undefined
  },
  alignTop: (env) => cellAlign(env, 'top'),
  cellTop: (env) => cellAlign(env, 'top'),
  cellMiddle: (env) => cellAlign(env, 'center'),
  cellBottom: (env) => cellAlign(env, 'bottom'),
  textDirection: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table cell first.'
    at.cell.textDirection = at.cell.textDirection === 'vertical' ? undefined : 'vertical'
    return undefined
  },
  eraser: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table cell first.'
    at.cell.borders = at.cell.borders && Object.keys(at.cell.borders).length ? undefined : { top: { val: 'single', sz: 0.5, color: 'auto' }, bottom: { val: 'single', sz: 0.5, color: 'auto' }, left: { val: 'single', sz: 0.5, color: 'auto' }, right: { val: 'single', sz: 0.5, color: 'auto' } }
    return undefined
  },
  gridlines: (env) => { env.doc.gridlines = !env.doc.gridlines; return undefined },
  tableProps: (env) => {
    const at = needTable(env)
    if (!at) return 'Place the cursor in a table first.'
    if (env.arg === 'left' || env.arg === 'center' || env.arg === 'right') at.table.align = env.arg
    return undefined
  },
}

function setFit(env: CommandEnv, fit: NonNullable<Table['autoFit']>): string | undefined {
  const at = needTable(env)
  if (!at) return 'Place the cursor in a table first.'
  at.table.autoFit = fit
  return undefined
}

function cellAlign(env: CommandEnv, value: NonNullable<Cell['vAlign']>): string | undefined {
  const at = needTable(env)
  if (!at) return 'Place the cursor in a table cell first.'
  at.cell.vAlign = value
  return undefined
}

function sortTable(env: CommandEnv, ascending: boolean): string | undefined {
  const at = needTable(env)
  if (!at) return 'Place the cursor in a table first.'
  const key = at.col
  const header = at.table.headerRows || 0
  const head = at.table.rows.slice(0, header)
  const body = at.table.rows.slice(header).sort((a, b) => {
    const av = tableCellText(a[key] ?? { paras: [] })
    const bv = tableCellText(b[key] ?? { paras: [] })
    const an = numeric(av); const bn = numeric(bv)
    const cmp = an != null && bn != null ? an - bn : av.localeCompare(bv, undefined, { numeric: true })
    return ascending ? cmp : -cmp
  })
  at.table.rows = [...head, ...body]
  return undefined
}

function formula(env: CommandEnv, kind: string): string | undefined {
  const at = needTable(env)
  if (!at) return 'Place the cursor in a table cell first.'
  const current = at.table.rows[at.row]?.[at.col]
  if (!current) return undefined
  const sources: string[] = []
  for (let r = 0; r < at.row; r++) sources.push(tableCellText(at.table.rows[r]?.[at.col] ?? { paras: [] }))
  if (!sources.some((s) => numeric(s) != null)) {
    for (let c = 0; c < at.col; c++) sources.push(tableCellText(at.table.rows[at.row]?.[c] ?? { paras: [] }))
  }
  const values = sources.map(numeric).filter((n): n is number => n != null)
  const result = kind === 'average' && values.length ? values.reduce((a, b) => a + b, 0) / values.length : values.reduce((a, b) => a + b, 0)
  current.paras = [{ ...emptyPara(), runs: [{ text: kind === 'average' ? result.toFixed(2) : String(result) }] }]
  return undefined
}

/* --------------------------------------------------------- character format */

function mapRuns(para: Para, fn: (r: RPr) => RPr): void {
  if (!para.runs.length) { para.runs = [{ text: '', rPr: {} }] }
  for (const run of para.runs) {
    if (run.raw) continue
    run.rPr = fn({ ...run.rPr })
    if (!Object.keys(run.rPr).length) delete run.rPr
  }
}

function sizeOf(env: CommandEnv): number {
  const r = env.loc?.para?.runs.find((run) => run.rPr?.size)?.rPr?.size
  return r || env.doc.defaultSize || 12
}

function setSize(env: CommandEnv, delta: number): string | undefined {
  const para = env.loc?.para
  if (!para) return 'Place the cursor in a paragraph first.'
  const next = Math.max(4, Math.min(72, sizeOf(env) + delta))
  mapRuns(para, (r) => ({ ...r, size: next }))
  return undefined
}

export const formatCommands: Record<string, CommandFn> = {
  changeCase: (env) => {
    const para = env.loc?.para
    if (!para) return undefined
    const mode = env.arg === 'upper' || env.arg === 'lower' || env.arg === 'title' || env.arg === 'sentence' ? env.arg : 'title'
    for (const run of para.runs) if (run.text && !run.raw) run.text = mode === 'sentence'
      ? run.text.charAt(0).toUpperCase() + run.text.slice(1).toLowerCase()
      : toggleCase(run.text, mode === 'title' ? 'title' : mode)
    return undefined
  },
  changeCaseCycle: (env) => {
    const text = paraText(env.loc?.para || emptyPara())
    const mode = text === text.toUpperCase() ? 'lower' : text === text.toLowerCase() ? 'title' : 'upper'
    return formatCommands.changeCase({ ...env, arg: mode })
  },
  allCaps: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    mapRuns(para, (r) => ({ ...r, caps: r.caps === 'all' ? undefined : 'all' }))
    return undefined
  },
  smallCaps: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    mapRuns(para, (r) => ({ ...r, caps: r.caps === 'small' ? undefined : 'small' }))
    return undefined
  },
  dblUnderline: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    mapRuns(para, (r) => ({ ...r, underline: true, underlineStyle: r.underlineStyle === 'double' ? undefined : 'double' }))
    return undefined
  },
  wordUnderline: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    mapRuns(para, (r) => ({ ...r, underline: true, underlineStyle: r.underlineStyle === 'words' ? undefined : 'words' }))
    return undefined
  },
  hiddenText: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    mapRuns(para, (r) => ({ ...r, hidden: !r.hidden }))
    return undefined
  },
  superscript: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    mapRuns(para, (r) => ({ ...r, vertAlign: r.vertAlign === 'superscript' ? undefined : 'superscript' }))
    return undefined
  },
  subscript: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    mapRuns(para, (r) => ({ ...r, vertAlign: r.vertAlign === 'subscript' ? undefined : 'subscript' }))
    return undefined
  },
  growFont: (env) => setSize(env, 1),
  growFont1: (env) => setSize(env, 1),
  shrinkFont: (env) => setSize(env, -1),
  shrinkFont1: (env) => setSize(env, -1),
  resetChar: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    for (const run of para.runs) if (!run.raw) delete run.rPr
    return undefined
  },
  resetPara: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = undefined
    return undefined
  },
  hangIndent: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    const indent = para.pPr?.indent || 48
    para.pPr = { ...para.pPr, indent, hanging: para.pPr?.hanging || 24 }
    return undefined
  },
  unhangIndent: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, hanging: undefined }
    return undefined
  },
  incIndent: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, indent: (para.pPr?.indent || 0) + 24 }
    return undefined
  },
  decIndent: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, indent: Math.max(0, (para.pPr?.indent || 0) - 24) }
    return undefined
  },
  spaceBefore: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, spaceBefore: para.pPr?.spaceBefore ? 0 : 12 }
    return undefined
  },
  single: (env) => setLine(env, 1),
  onePointFive: (env) => setLine(env, 1.5),
  double: (env) => setLine(env, 2),
  lineSpacing: (env) => {
    const current = env.loc?.para?.pPr?.line || 1
    return setLine(env, current < 1.2 ? 1.5 : current < 1.7 ? 2 : 1)
  },
  keepNext: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, keepNext: !para.pPr?.keepNext }
    return undefined
  },
  keepLines: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, keepLines: !para.pPr?.keepLines }
    return undefined
  },
  bordersDlg: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, borders: para.pPr?.borders ? undefined : { top: { val: 'single', sz: 0.5, color: 'auto' }, bottom: { val: 'single', sz: 0.5, color: 'auto' }, left: { val: 'single', sz: 0.5, color: 'auto' }, right: { val: 'single', sz: 0.5, color: 'auto' } } }
    return undefined
  },
  borderBtn: (env) => formatCommands.bordersDlg(env),
  dropCap: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    const first = para.runs.find((r) => r.text)
    if (first?.text) {
      first.rPr = { ...first.rPr, size: Math.max(28, (first.rPr?.size || 12) * 3), bold: true }
      first.text = first.text.slice(0, 1)
    }
    return undefined
  },
  olFirstLine: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, indent: 0, hanging: undefined }
    return undefined
  },
  olFormatting: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, numId: para.pPr?.numId ? undefined : 1, ilvl: 0 }
    return undefined
  },
  restartNumbering: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, numId: 1, ilvl: 0 }
    return 'Numbering restarted at 1.'
  },
  continueNumbering: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    const previous = env.doc.blocks[env.loc!.blockIndex - 1]
    const numId = isPara(previous) ? previous.pPr?.numId : undefined
    if (numId == null) return 'The previous paragraph is not numbered.'
    para.pPr = { ...para.pPr, numId, ilvl: isPara(previous) ? previous.pPr?.ilvl || 0 : 0 }
    return undefined
  },
  autoFormat: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    for (const run of para.runs) if (run.text) run.text = autoFormatText(run.text)
    return undefined
  },
  pasteFormat: (env) => {
    const para = env.loc?.para
    if (!para || !formatBrush) return 'Copy formatting first.'
    for (const run of para.runs) if (!run.raw) run.rPr = { ...formatBrush }
    return 'Formatting applied.'
  },
  copyFormat: (env) => {
    const para = env.loc?.para
    if (!para) return undefined
    const source = env.selection ? para.runs.find((r) => r.text && r.rPr) : para.runs[0]
    formatBrush = { ...(source?.rPr || {}) }
    return 'Formatting copied.'
  },
  painter: (env) => {
    if (!formatBrush) return formatCommands.copyFormat!(env)
    return formatCommands.pasteFormat!(env)
  },
}

let formatBrush: RPr | undefined

function setLine(env: CommandEnv, value: number): string | undefined {
  const para = env.loc?.para
  if (!para) return undefined
  para.pPr = { ...para.pPr, line: value }
  return undefined
}

/* -------------------------------------------------------------- outlines */

function styleLevel(style: string | undefined): number {
  const match = /^Heading([1-9])$/i.exec(style || '')
  return match ? Number(match[1]) : 0
}

export const outlineCommands: Record<string, CommandFn> = {
  promoteH1: (env) => setLevel(env, 1),
  promote: (env) => setLevel(env, Math.max(1, styleLevel(env.loc?.para?.pPr?.style) - 1 || 1)),
  demote: (env) => setLevel(env, Math.min(9, styleLevel(env.loc?.para?.pPr?.style) + 1 || 2)),
  demoteBody: (env) => setLevel(env, 0),
  moveUp: (env) => {
    const at = env.loc
    if (!at || at.blockIndex <= 0) return 'Already at the start of the document.'
    const blocks = env.doc.blocks
    ;[blocks[at.blockIndex - 1], blocks[at.blockIndex]] = [blocks[at.blockIndex], blocks[at.blockIndex - 1]]
    at.blockIndex--
    return undefined
  },
  moveDown: (env) => {
    const at = env.loc
    if (!at || at.blockIndex >= env.doc.blocks.length - 1) return 'Already at the end of the document.'
    const blocks = env.doc.blocks
    ;[blocks[at.blockIndex + 1], blocks[at.blockIndex]] = [blocks[at.blockIndex], blocks[at.blockIndex + 1]]
    at.blockIndex++
    return undefined
  },
  expand: (env) => setCollapsed(env, false),
  collapse: (env) => setCollapsed(env, true),
  styleNormal: (env) => setLevel(env, 0),
  styleH1: (env) => setLevel(env, 1),
  styleH2: (env) => setLevel(env, 2),
  styleH3: (env) => setLevel(env, 3),
  styleListBullet: (env) => {
    const para = env.loc?.para; if (!para) return undefined
    para.pPr = { ...para.pPr, numId: 1, ilvl: 0, style: undefined }
    return undefined
  },
  updateTOC: (env) => {
    const headings = env.doc.blocks.filter(isPara).filter((p) => styleLevel(p.pPr?.style))
    return `Table of contents updated: ${headings.length} heading(s).`
  },
  gotoTOC: () => 'The first heading is selected in the Document Map.',
}

function setLevel(env: CommandEnv, level: number): string | undefined {
  const para = env.loc?.para
  if (!para) return undefined
  para.pPr = {
    ...para.pPr,
    style: level ? `Heading${level}` : 'Normal',
    outlineLvl: level ? level - 1 : undefined,
  }
  return undefined
}

function setCollapsed(env: CommandEnv, collapsed: boolean): string | undefined {
  const para = env.loc?.para
  if (!para) return undefined
  const level = styleLevel(para.pPr?.style)
  if (!level) return 'Place the cursor in a heading first.'
  for (let i = env.loc!.blockIndex + 1; i < env.doc.blocks.length; i++) {
    const block = env.doc.blocks[i]
    if (!isPara(block)) continue
    if (styleLevel(block.pPr?.style) && styleLevel(block.pPr?.style) <= level) break
    block.pPr = { ...block.pPr, collapsed }
  }
  return undefined
}

/* --------------------------------------------------------------- assembly */

export const documentCommands: Record<string, CommandFn> = {
  ...tableCommands, ...formatCommands, ...outlineCommands,
}

export function commandText(env: CommandEnv): string {
  return env.selection || (env.loc?.para ? paraText(env.loc.para) : '')
}

export function allBlockText(doc: Doc): string {
  return doc.blocks.map((b: Block) => blockText(b)).join('\n')
}

export type { Run }
