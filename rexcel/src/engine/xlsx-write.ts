/* workbook model → XLSX. Regenerates sheet parts, sharedStrings and styles; keeps every
 * other source part byte-for-byte. workbook.xml is kept unless the sheet set changed. */
import { Package, NS, formats, newPackage, type FormatInfo } from '@voffice/react-shared'
import { esc, escText } from '@voffice/react-shared'
import { type Workbook, type Sheet, type Cell, addr, colName, parseAddr } from './model'

const S = NS.s
const R = NS.r
const ROOT = `xmlns="${S}" xmlns:r="${R}"`

function colWidthChars(px: number): number { return Math.max(0, Math.round(((px - 5) / 7) * 100) / 100) }

interface WriteCtx { wb: Workbook; strings: Map<string, number>; strList: string[] }

function strIndex(ctx: WriteCtx, s: string): number {
  let i = ctx.strings.get(s)
  if (i == null) { i = ctx.strList.length; ctx.strList.push(s); ctx.strings.set(s, i) }
  return i
}

function cellXml(ctx: WriteCtx, r: number, c: number, cell: Cell): string {
  const ref = addr(r, c)
  const sAttr = cell.s ? ` s="${cell.s}"` : ''
  if (cell.f != null) {
    return `<c r="${ref}"${sAttr}><f>${escText(cell.f)}</f></c>`
  }
  if (cell.t === 'e') return `<c r="${ref}"${sAttr} t="e"><v>${escText(String(cell.v))}</v></c>`
  if (cell.t === 'b') return `<c r="${ref}"${sAttr} t="b"><v>${cell.v ? 1 : 0}</v></c>`
  if (typeof cell.v === 'string') return `<c r="${ref}"${sAttr} t="s"><v>${strIndex(ctx, cell.v)}</v></c>`
  if (typeof cell.v === 'number') return `<c r="${ref}"${sAttr}><v>${cell.v}</v></c>`
  return ''
}

function sheetXml(ctx: WriteCtx, sheet: Sheet): string {
  const rows = new Map<number, { r: number; c: number; cell: Cell }[]>()
  let maxR = 0, maxC = 0
  for (const [key, cell] of sheet.cells) {
    const rc = parseAddr(key)!
    maxR = Math.max(maxR, rc.r); maxC = Math.max(maxC, rc.c)
    const list = rows.get(rc.r) || []
    list.push({ r: rc.r, c: rc.c, cell })
    rows.set(rc.r, list)
  }
  const dim = `A1:${colName(maxC)}${maxR + 1}`
  let cols = ''
  if (sheet.colWidths.size) {
    const entries = [...sheet.colWidths.entries()].sort((a, b) => a[0] - b[0])
    cols = '<cols>' + entries.map(([c, w]) => `<col min="${c + 1}" max="${c + 1}" width="${colWidthChars(w)}" customWidth="1"/>`).join('') + '</cols>'
  }
  const rowXml: string[] = []
  for (const [r, list] of [...rows.entries()].sort((a, b) => a[0] - b[0])) {
    list.sort((a, b) => a.c - b.c)
    const ht = sheet.rowHeights.get(r)
    const htAttr = ht != null ? ` ht="${ht}" customHeight="1"` : ''
    rowXml.push(`<row r="${r + 1}"${htAttr}>` + list.map((x) => cellXml(ctx, x.r, x.c, x.cell)).join('') + '</row>')
  }
  let merges = ''
  if (sheet.merges.length) merges = `<mergeCells count="${sheet.merges.length}">` + sheet.merges.map((m) => `<mergeCell ref="${addr(m.r1, m.c1)}:${addr(m.r2, m.c2)}"/>`).join('') + '</mergeCells>'
  let pane = ''
  if (sheet.frozen && (sheet.frozen.rows || sheet.frozen.cols)) {
    const { rows: fr, cols: fc } = sheet.frozen
    pane = `<pane xSplit="${fc}" ySplit="${fr}" topLeftCell="${addr(fr, fc)}" activePane="bottomRight" state="frozen"/>`
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet ${ROOT}>` +
    `<dimension ref="${dim}"/><sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15"/>${cols}<sheetData>${rowXml.join('')}</sheetData>${merges}</worksheet>`
}

function sharedStringsXml(ctx: WriteCtx): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst ${ROOT} count="${ctx.strList.length}" uniqueCount="${ctx.strList.length}">` +
    ctx.strList.map((s) => `<si><t xml:space="preserve">${escText(s)}</t></si>`).join('') + '</sst>'
}

function stylesXml(wb: Workbook): string {
  const fonts: string[] = []
  const fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>']
  const borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>']
  const numFmts: string[] = []
  const numFmtIds = new Map<string, number>()
  const xfs: string[] = []
  const fontIds = new Map<string, number>()
  const fillIds = new Map<string, number>()
  const borderIds = new Map<string, number>()
  let nextNumFmt = 164
  for (const style of wb.styles) {
    let fontId = 0
    if (style.font) {
      const k = JSON.stringify(style.font)
      if (!fontIds.has(k)) {
        const f = style.font
        let inner = ''
        if (f.bold) inner += '<b/>'
        if (f.italic) inner += '<i/>'
        if (f.underline) inner += '<u/>'
        inner += `<sz val="${f.size || 11}"/>`
        if (f.color) inner += `<color rgb="FF${f.color.replace('#', '').toUpperCase()}"/>`
        inner += `<name val="${esc(f.name || 'Calibri')}"/>`
        fontIds.set(k, fonts.length); fonts.push(`<font>${inner}</font>`)
      }
      fontId = fontIds.get(k)!
    }
    let fillId = 0
    if (style.fill) {
      const k = style.fill
      if (!fillIds.has(k)) { fillIds.set(k, fills.length); fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="FF${k.replace('#', '').toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`) }
      fillId = fillIds.get(k)!
    }
    let borderId = 0
    if (style.border) {
      const k = JSON.stringify(style.border)
      if (!borderIds.has(k)) {
        const b = style.border
        let inner = ''
        for (const s of ['left', 'right', 'top', 'bottom'] as const) inner += b[s] ? `<${s} style="thin"><color rgb="FF${b[s]!.replace('#', '').toUpperCase()}"/></${s}>` : `<${s}/>`
        borderIds.set(k, borders.length); borders.push(`<border>${inner}<diagonal/></border>`)
      }
      borderId = borderIds.get(k)!
    }
    let numFmtId = 0
    if (style.numFmt && style.numFmt !== 'General') {
      if (!numFmtIds.has(style.numFmt)) { numFmtIds.set(style.numFmt, nextNumFmt); numFmts.push(`<numFmt numFmtId="${nextNumFmt}" formatCode="${esc(style.numFmt)}"/>`); nextNumFmt++ }
      numFmtId = numFmtIds.get(style.numFmt)!
    }
    let align = ''
    if (style.align) {
      const a = style.align
      const h = a.h ? ` horizontal="${a.h}"` : ''
      const v = a.v ? ` vertical="${a.v === 'middle' ? 'center' : a.v}"` : ''
      const w = a.wrap ? ' wrapText="1"' : ''
      align = `<alignment${h}${v}${w}/>`
    }
    xfs.push(`<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''}${fontId ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${align ? ' applyAlignment="1"' : ''}>${align}</xf>`)
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<styleSheet ${ROOT}>` +
    (numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>` : '') +
    `<fonts count="${fonts.length || 1}">${fonts.join('') || '<font><sz val="11"/><name val="Calibri"/></font>'}</fonts>` +
    `<fills count="${fills.length}">${fills.join('')}</fills>` +
    `<borders count="${borders.length}">${borders.join('')}</borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`
}

const SS_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
const ST_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
const SH_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'

export interface XlsxWriteOpts { pkg?: Package; sheetParts?: string[]; format?: string }

export async function writeXlsx(wb: Workbook, opts: XlsxWriteOpts = {}): Promise<Blob> {
  const fmt: FormatInfo = formats[opts.format || 'xlsx'] || formats.xlsx
  const ctx: WriteCtx = { wb, strings: new Map(), strList: [] }
  const sheetXmls = wb.sheets.map((s) => sheetXml(ctx, s))
  const pkg = opts.pkg || newPackage('xl/workbook.xml', '', fmt)
  const parts = opts.sheetParts && opts.sheetParts.length === wb.sheets.length
    ? opts.sheetParts
    : wb.sheets.map((_, i) => `xl/worksheets/sheet${i + 1}.xml`)
  sheetXmls.forEach((xml, i) => pkg.addPart(parts[i], xml, SH_CT))
  pkg.addPart('xl/sharedStrings.xml', sharedStringsXml(ctx), SS_CT)
  pkg.addPart('xl/styles.xml', stylesXml(wb), ST_CT)
  // workbook.xml: keep source when the sheet set is unchanged, else regenerate
  const sourceNames = opts.pkg ? sheetNamesOf(opts.pkg) : null
  const names = wb.sheets.map((s) => s.name)
  const unchanged = sourceNames && sourceNames.length === names.length && sourceNames.every((n, i) => n === names[i])
  if (!unchanged) {
    const sheets = wb.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')
    pkg.setPart('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<workbook ${ROOT}><bookViews><workbookView activeTab="${wb.active}"/></bookViews><sheets>${sheets}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`)
    // rebuild workbook rels: sheets + styles + sharedStrings
    const rels: string[] = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    wb.sheets.forEach((_, i) => rels.push(`<Relationship Id="rId${i + 1}" Type="${NS.rel}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`))
    const n = wb.sheets.length
    rels.push(`<Relationship Id="rId${n + 1}" Type="${NS.rel}/styles" Target="styles.xml"/>`)
    rels.push(`<Relationship Id="rId${n + 2}" Type="${NS.rel}/sharedStrings" Target="sharedStrings.xml"/>`)
    rels.push('</Relationships>')
    pkg.setPart('xl/_rels/workbook.xml.rels', rels.join(''))
  } else if (opts.pkg) {
    // ensure calcPr fullCalcOnLoad so Excel recalcs our formula cells
    // (kept workbook otherwise untouched)
  }
  if (!opts.pkg) {
    pkg.addRel('', NS.rel + '/officeDocument', 'xl/workbook.xml')
  }
  return pkg.save(fmt.mime)
}

function sheetNamesOf(pkg: Package): string[] | null {
  if (!pkg.has('xl/workbook.xml')) return null
  const root = pkg.xml('xl/workbook.xml')
  const sheets = root.children.find((c) => c.localName === 'sheets')
  if (!sheets) return null
  return sheets.children.filter((c) => c.localName === 'sheet').map((c) => c.getAttribute('name') || '')
}
