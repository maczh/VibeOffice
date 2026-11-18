/* XLSX → workbook model. Reads workbook, shared strings, styles and each sheet;
 * keeps the source Package for write-back of untouched parts. */
import { Package, NS } from '@voffice/react-shared'
import type { El } from '@voffice/react-shared'
import { type Workbook, type Sheet, type Cell, type CellStyle, newSheet, colIndex, addr, parseAddr, styleIndex } from './model'

const S = NS.s
const R = NS.r
const kids = (el: El, l: string): El[] => el.children.filter((c) => c.localName === l)
const kid = (el: El, l: string): El | undefined => el.children.find((c) => c.localName === l)
const at = (el: El | undefined, n: string): string | null => (el ? el.getAttribute(n) : null)

const BUILTIN_FMT: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 5: '#,##0;[Red]-#,##0',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM',
  20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm', 37: '#,##0;(#,##0)', 38: '#,##0;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 44: '_("$"* #,##0.00_)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
}

function siText(si: El): string {
  const t = kid(si, 't')
  if (t) return t.textContent
  return kids(si, 'r').map((r) => kid(r, 't')?.textContent || '').join('')
}

interface StyleTables { fonts: El[]; fills: El[]; borders: El[]; xfs: El[]; numFmts: Map<number, string> }

function readStyles(pkg: Package): StyleTables {
  const t: StyleTables = { fonts: [], fills: [], borders: [], xfs: [], numFmts: new Map() }
  if (!pkg.has('xl/styles.xml')) return t
  const root = pkg.xml('xl/styles.xml')
  const nf = kid(root, 'numFmts')
  if (nf) for (const n of kids(nf, 'numFmt')) t.numFmts.set(+(at(n, 'numFmtId') || 0), at(n, 'formatCode') || 'General')
  const fo = kid(root, 'fonts'); if (fo) t.fonts = kids(fo, 'font')
  const fi = kid(root, 'fills'); if (fi) t.fills = kids(fi, 'fill')
  const bo = kid(root, 'borders'); if (bo) t.borders = kids(bo, 'border')
  const xf = kid(root, 'cellXfs'); if (xf) t.xfs = kids(xf, 'xf')
  return t
}

function colorOf(el: El | undefined): string | undefined {
  if (!el) return undefined
  const rgb = at(el, 'rgb')
  if (rgb && rgb.length === 8) return ('#' + rgb.slice(2)).toLowerCase()
  if (rgb && rgb.length === 6) return ('#' + rgb).toLowerCase()
  return undefined
}

function xfToStyle(xf: El | undefined, t: StyleTables): CellStyle {
  const s: CellStyle = {}
  if (!xf) return s
  const font = t.fonts[+(at(xf, 'fontId') || 0)]
  if (font) {
    const f: NonNullable<CellStyle['font']> = {}
    const sz = kid(font, 'sz'); if (sz) f.size = +(at(sz, 'val') || 11)
    const nm = kid(font, 'name'); if (nm) f.name = at(nm, 'val') || undefined
    if (kid(font, 'b')) f.bold = true
    if (kid(font, 'i')) f.italic = true
    if (kid(font, 'u')) f.underline = true
    const col = colorOf(kid(font, 'color')); if (col) f.color = col
    if (Object.keys(f).length) s.font = f
  }
  const fill = t.fills[+(at(xf, 'fillId') || 0)]
  if (fill) {
    const pf = kid(fill, 'patternFill')
    const fg = colorOf(kid(pf || fill, 'fgColor'))
    if (fg && at(pf, 'patternType') !== 'none') s.fill = fg
  }
  const bd = t.borders[+(at(xf, 'borderId') || 0)]
  if (bd) {
    const b: NonNullable<CellStyle['border']> = {}
    for (const side of ['left', 'right', 'top', 'bottom'] as const) {
      const e = kid(bd, side)
      if (e && at(e, 'style')) b[side] = colorOf(kid(e, 'color')) || '#000000'
    }
    if (Object.keys(b).length) s.border = b
  }
  const al = kid(xf, 'alignment')
  if (al) {
    const a: NonNullable<CellStyle['align']> = {}
    const h = at(al, 'horizontal'); if (h) a.h = h === 'general' ? undefined : (h as 'left' | 'center' | 'right')
    const v = at(al, 'vertical'); if (v) a.v = v === 'center' ? 'middle' : (v as 'top' | 'bottom')
    if (at(al, 'wrapText') === '1') a.wrap = true
    if (Object.keys(a).length) s.align = a
  }
  const numId = +(at(xf, 'numFmtId') || 0)
  const fmt = t.numFmts.get(numId) || BUILTIN_FMT[numId]
  if (fmt && fmt !== 'General') s.numFmt = fmt
  return s
}

export interface XlsxReadResult { wb: Workbook; pkg: Package; sheetParts: string[] }

export async function readXlsx(bytes: Uint8Array | ArrayBuffer): Promise<XlsxReadResult> {
  const pkg = await Package.open(bytes)
  const wbRoot = pkg.xml('xl/workbook.xml')
  const wbRels = pkg.relsOf('xl/workbook.xml')
  const shared: string[] = []
  if (pkg.has('xl/sharedStrings.xml')) {
    const sr = pkg.xml('xl/sharedStrings.xml')
    for (const si of kids(sr, 'si')) shared.push(siText(si))
  }
  const st = readStyles(pkg)
  const wb: Workbook = { sheets: [], styles: [{}], active: 0 }
  const sheetParts: string[] = []
  const sheetsEl = kid(wbRoot, 'sheets')
  const entries = sheetsEl ? kids(sheetsEl, 'sheet') : []
  for (const sh of entries) {
    const name = at(sh, 'name') || 'Sheet'
    const rid = sh.getAttributeNS(R, 'id') || at(sh, 'r:id') || ''
    const rel = wbRels.find((r) => r.id === rid)
    const part = rel ? (rel.target.startsWith('/') ? rel.target.slice(1) : 'xl/' + rel.target.replace(/^\.\//, '')) : ''
    sheetParts.push(part)
    const sheet = newSheet(name)
    if (part && pkg.has(part)) readSheet(pkg, part, sheet, shared, st, wb)
    wb.sheets.push(sheet)
  }
  if (!wb.sheets.length) { wb.sheets.push(newSheet('Sheet1')); sheetParts.push('xl/worksheets/sheet1.xml') }
  const av = at(kid(wbRoot, 'bookViews') && kid(kid(wbRoot, 'bookViews')!, 'workbookView'), 'activeTab')
  if (av) wb.active = Math.min(+av, wb.sheets.length - 1)
  return { wb, pkg, sheetParts }
}

function readSheet(pkg: Package, part: string, sheet: Sheet, shared: string[], st: StyleTables, wb: Workbook): void {
  const root = pkg.xml(part)
  const cols = kid(root, 'cols')
  if (cols) for (const c of kids(cols, 'col')) {
    const min = +(at(c, 'min') || 1), max = +(at(c, 'max') || min)
    const w = +(at(c, 'width') || 8.43)
    for (let i = min; i <= max && i <= 16384; i++) sheet.colWidths.set(i - 1, Math.round(w * 7 + 5))
  }
  const sd = kid(root, 'sheetData')
  if (sd) for (const row of kids(sd, 'row')) {
    const r = +(at(row, 'r') || 1) - 1
    const ht = at(row, 'ht')
    if (ht) sheet.rowHeights.set(r, Math.round(+ht))
    for (const c of kids(row, 'c')) {
      const ref = at(c, 'r') || ''
      const rc = parseAddr(ref)
      if (!rc) continue
      const t = at(c, 't') || 'n'
      const sIdx = at(c, 's')
      const vEl = kid(c, 'v')
      const fEl = kid(c, 'f')
      const cell: Cell = {}
      if (fEl) cell.f = fEl.textContent
      if (t === 's') { cell.v = shared[+(vEl?.textContent || 0)] ?? ''; cell.t = 's' }
      else if (t === 'inlineStr') { const is = kid(c, 'is'); cell.v = is ? siText(is) : ''; cell.t = 's' }
      else if (t === 'str') { cell.v = vEl?.textContent || ''; cell.t = 's' }
      else if (t === 'b') { cell.v = vEl?.textContent === '1'; cell.t = 'b' }
      else if (t === 'e') { cell.v = vEl?.textContent || '#VALUE!'; cell.t = 'e' }
      else { const raw = vEl?.textContent; cell.v = raw != null && raw !== '' ? +raw : null; cell.t = 'n' }
      if (sIdx != null) {
        const style = xfToStyle(st.xfs[+sIdx], st)
        if (Object.keys(style).length) cell.s = styleIndex(wb, style)
      }
      if (cell.v != null || cell.f || cell.s != null) sheet.cells.set(addr(rc.r, rc.c), cell)
    }
  }
  const mc = kid(root, 'mergeCells')
  if (mc) for (const m of kids(mc, 'mergeCell')) {
    const [a, b] = (at(m, 'ref') || '').split(':')
    const ra = parseAddr(a); const rb = parseAddr(b || a)
    if (ra && rb) sheet.merges.push({ r1: ra.r, c1: ra.c, r2: rb.r, c2: rb.c })
  }
  const sv = kid(root, 'sheetViews')
  const pane = sv && kid(kid(sv, 'sheetView') || sv, 'pane')
  if (pane) {
    const x = +(at(pane, 'xSplit') || 0), y = +(at(pane, 'ySplit') || 0)
    if (x || y) sheet.frozen = { rows: y, cols: x }
  }
}
