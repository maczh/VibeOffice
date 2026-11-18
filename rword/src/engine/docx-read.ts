/* DOCX → model. Reads word/document.xml, styles, numbering and relationships;
 * images become blob URLs. The source Package is kept so untouched parts write back unchanged. */
import { Package, NS } from '@voffice/react-shared'
import type { El } from '@voffice/react-shared'
import {
  type Doc, type Para, type Run, type RPr, type PPr, type Table, type Cell, type Block,
  type StyleDef, type Section, defaultSection, emptyPara, uid,
} from './model'

const W = NS.w
const R = NS.r
const EMU_PER_PX = 914400 / 96
const TWIP_PER_PX = 1440 / 96

const kids = (el: El, local: string): El[] => el.children.filter((c) => c.localName === local)
const kid = (el: El, local: string): El | undefined => el.children.find((c) => c.localName === local)
const attr = (el: El | undefined, name: string): string | null => {
  if (!el) return null
  const v = el.getAttribute(name)
  if (v != null) return v
  // document.xml attributes are namespace-prefixed (w:val); match by local name
  const a = el.attributes.find((x) => x.localName === name)
  return a ? a.value : null
}
const on = (v: string | null): boolean => v == null ? false : v !== '0' && v !== 'false' && v !== 'none'

function readRPr(el: El | undefined): RPr | undefined {
  if (!el) return undefined
  const r: RPr = {}
  if (kid(el, 'b')) r.bold = on(attr(kid(el, 'b'), 'val') || '1')
  if (kid(el, 'i')) r.italic = on(attr(kid(el, 'i'), 'val') || '1')
  if (kid(el, 'u')) r.underline = true
  if (kid(el, 'strike')) r.strike = on(attr(kid(el, 'strike'), 'val') || '1')
  const fonts = kid(el, 'rFonts')
  if (fonts) r.font = attr(fonts, 'ascii') || attr(fonts, 'hAnsi') || undefined
  const sz = kid(el, 'sz')
  if (sz) r.size = +(attr(sz, 'val') || 22) / 2
  const color = kid(el, 'color')
  if (color && attr(color, 'val') && attr(color, 'val') !== 'auto') r.color = ('#' + attr(color, 'val')).toLowerCase()
  const hl = kid(el, 'highlight')
  if (hl) r.highlight = attr(hl, 'val') || undefined
  const va = kid(el, 'vertAlign')
  if (va) r.vertAlign = (attr(va, 'val') as RPr['vertAlign']) || undefined
  return Object.keys(r).length ? r : undefined
}

function readPPr(el: El | undefined): PPr | undefined {
  if (!el) return undefined
  const p: PPr = {}
  if (kid(el, 'pageBreakBefore')) p.pageBreakBefore = true
  const jc = kid(el, 'jc')
  if (jc) {
    const v = attr(jc, 'val') || 'left'
    p.align = v === 'both' ? 'justify' : v === 'start' ? 'left' : v === 'end' ? 'right' : (v as PPr['align'])
  }
  const st = kid(el, 'pStyle')
  if (st) p.style = attr(st, 'val') || undefined
  const numPr = kid(el, 'numPr')
  if (numPr) {
    const ilvl = kid(numPr, 'ilvl'); const numId = kid(numPr, 'numId')
    if (numId) p.numId = +(attr(numId, 'val') || 0)
    if (ilvl) p.ilvl = +(attr(ilvl, 'val') || 0)
  }
  const ind = kid(el, 'ind')
  if (ind) {
    const left = attr(ind, 'left') || attr(ind, 'start')
    if (left) p.indent = Math.round(+left / TWIP_PER_PX)
    const hang = attr(ind, 'hanging')
    if (hang) p.hanging = Math.round(+hang / TWIP_PER_PX)
  }
  const spacing = kid(el, 'spacing')
  if (spacing) {
    const before = attr(spacing, 'before'); const after = attr(spacing, 'after'); const line = attr(spacing, 'line')
    if (before) p.spaceBefore = Math.round(+before / TWIP_PER_PX)
    if (after) p.spaceAfter = Math.round(+after / TWIP_PER_PX)
    if (line && attr(spacing, 'lineRule') !== 'auto') p.line = Math.round(+line / TWIP_PER_PX)
    else if (line) p.line = +line / 240
  }
  return Object.keys(p).length ? p : undefined
}

function readRun(el: El, pkg: Package, part: string): Run {
  const rPr = readRPr(kid(el, 'rPr'))
  if (kid(el, 'br')) return { br: true, rPr }
  if (kid(el, 'tab')) return { tab: true, rPr }
  const drawing = kid(el, 'drawing')
  if (drawing) {
    const img = readImage(drawing, pkg, part)
    if (img) return { image: img, rPr }
    return { raw: serializeRaw(el), rPr }
  }
  const t = kid(el, 't')
  if (t) return { text: t.textContent, rPr }
  // anything else (fields, symbols, …): keep verbatim
  const txt = el.textContent
  if (txt) return { text: txt, rPr }
  return { raw: serializeRaw(el), rPr }
}

function serializeRaw(el: El): string {
  // lightweight verbatim serialization preserving prefixes/attrs
  const attrs = el.attributes.map((a) => ` ${a.name}="${a.value.replace(/"/g, '&quot;')}"`).join('')
  const inner = el.childNodes.map((n) => n.nodeType === 3 ? (n as { data: string }).data.replace(/&/g, '&amp;').replace(/</g, '&lt;') : serializeRaw(n as El)).join('')
  return `<${el.nodeName}${attrs}>${inner}</${el.nodeName}>`
}

function readImage(drawing: El, pkg: Package, part: string): Run['image'] | null {
  const extent = drawing.getElementsByTagNameNS('*', 'extent')[0]
  const blip = drawing.getElementsByTagNameNS(NS.a, 'blip')[0]
  if (!blip) return null
  const rId = blip.getAttributeNS(R, 'embed') || blip.getAttribute('r:embed')
  if (!rId) return null
  const rel = pkg.rel(part, rId)
  if (!rel) return null
  const target = rel.target.startsWith('/') ? rel.target.slice(1) : resolve(part, rel.target)
  const bytes = pkg.bytes(target)
  if (!bytes) return null
  const ext = target.split('.').pop() || 'png'
  const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : ext === 'svg' ? 'image/svg+xml' : 'image/' + ext
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }))
  const cx = +(attr(extent, 'cx') || 0); const cy = +(attr(extent, 'cy') || 0)
  return { rId, part: target, url, w: Math.round(cx / EMU_PER_PX) || 100, h: Math.round(cy / EMU_PER_PX) || 100 }
}

function resolve(part: string, target: string): string {
  const base = part.includes('/') ? part.slice(0, part.lastIndexOf('/') + 1) : ''
  const out: string[] = []
  for (const s of (base + target).split('/')) {
    if (s === '.' || s === '') continue
    if (s === '..') out.pop(); else out.push(s)
  }
  return out.join('/')
}

function readPara(el: El, pkg: Package, part: string): Para {
  const p: Para = emptyPara(readPPr(kid(el, 'pPr')))
  for (const child of el.children) {
    if (child.localName === 'r') p.runs.push(readRun(child, pkg, part))
    else if (child.localName === 'hyperlink') {
      const rId = child.getAttributeNS(R, 'id') || child.getAttribute('r:id') || ''
      const rel = rId ? pkg.rel(part, rId) : undefined
      const url = rel?.target || ''
      for (const r of kids(child, 'r')) {
        const run = readRun(r, pkg, part)
        run.hyperlink = { rId, url }
        p.runs.push(run)
      }
    } else if (child.localName === 'pPr') {
      // handled
    } else if (child.localName === 'bookmarkStart' || child.localName === 'bookmarkEnd' || child.localName === 'proofErr') {
      // ignore
    } else {
      p.runs.push({ raw: serializeRaw(child) })
    }
  }
  return p
}

function readTable(el: El, pkg: Package, part: string): Table {
  const grid = kid(el, 'tblGrid')
  const widths = grid ? kids(grid, 'gridCol').map((c) => Math.round(+(attr(c, 'w') || 1000) / TWIP_PER_PX)) : []
  const rows: Cell[][] = []
  for (const tr of kids(el, 'tr')) {
    const row: Cell[] = []
    for (const tc of kids(tr, 'tc')) {
      const tcPr = kid(tc, 'tcPr')
      const cell: Cell = { paras: [] }
      const span = tcPr ? kid(tcPr, 'gridSpan') : undefined
      if (span) cell.gridSpan = +(attr(span, 'val') || 1)
      const vm = tcPr ? kid(tcPr, 'vMerge') : undefined
      if (vm) cell.vMerge = attr(vm, 'val') === 'restart' ? 'restart' : 'continue'
      const shd = tcPr ? kid(tcPr, 'shd') : undefined
      if (shd && attr(shd, 'fill') && attr(shd, 'fill') !== 'auto') cell.fill = '#' + attr(shd, 'fill')
      const tcW = tcPr ? kid(tcPr, 'tcW') : undefined
      if (tcW) cell.width = Math.round(+(attr(tcW, 'w') || 0) / TWIP_PER_PX)
      for (const p of kids(tc, 'p')) cell.paras.push(readPara(p, pkg, part))
      if (!cell.paras.length) cell.paras.push(emptyPara())
      row.push(cell)
    }
    rows.push(row)
  }
  return { kind: 'tbl', id: uid('t'), rows, widths }
}

function readStyles(pkg: Package): { styles: StyleDef[]; defaultFont: string; defaultSize: number } {
  const styles: StyleDef[] = []
  let defaultFont = 'Calibri'; let defaultSize = 11
  if (!pkg.has('word/styles.xml')) return { styles, defaultFont, defaultSize }
  const root = pkg.xml('word/styles.xml')
  const docDefaults = kid(root, 'docDefaults')
  const rPrDefault = docDefaults ? kid(kid(docDefaults, 'rPrDefault') || docDefaults, 'rPr') : undefined
  const dr = readRPr(rPrDefault)
  if (dr?.font) defaultFont = dr.font
  if (dr?.size) defaultSize = dr.size
  for (const s of kids(root, 'style')) {
    const id = attr(s, 'styleId') || ''
    const nameEl = kid(s, 'name')
    const def: StyleDef = { id, name: attr(nameEl, 'val') || id }
    const basedOn = kid(s, 'basedOn')
    if (basedOn) def.basedOn = attr(basedOn, 'val') || undefined
    def.rPr = readRPr(kid(s, 'rPr'))
    def.pPr = readPPr(kid(s, 'pPr'))
    styles.push(def)
  }
  return { styles, defaultFont, defaultSize }
}

function readNumbering(pkg: Package): Record<number, string> {
  const map: Record<number, string> = {}
  if (!pkg.has('word/numbering.xml')) return map
  const root = pkg.xml('word/numbering.xml')
  const abstract: Record<string, string> = {}
  for (const an of kids(root, 'abstractNum')) {
    const id = attr(an, 'abstractNumId') || ''
    const lvl0 = kids(an, 'lvl').find((l) => attr(l, 'ilvl') === '0') || kids(an, 'lvl')[0]
    const fmt = lvl0 ? attr(kid(lvl0, 'numFmt'), 'val') || 'bullet' : 'bullet'
    abstract[id] = fmt
  }
  for (const num of kids(root, 'num')) {
    const numId = +(attr(num, 'numId') || 0)
    const absId = attr(kid(num, 'abstractNumId'), 'val') || ''
    map[numId] = abstract[absId] || 'bullet'
  }
  return map
}

function readSection(body: El): Section {
  const sectPr = kid(body, 'sectPr')
  const sec = defaultSection()
  if (!sectPr) return sec
  const pgSz = kid(sectPr, 'pgSz')
  if (pgSz) { sec.pageW = +(attr(pgSz, 'w') || sec.pageW); sec.pageH = +(attr(pgSz, 'h') || sec.pageH) }
  const pgMar = kid(sectPr, 'pgMar')
  if (pgMar) {
    sec.margins = {
      top: +(attr(pgMar, 'top') || 1440), right: +(attr(pgMar, 'right') || 1440),
      bottom: +(attr(pgMar, 'bottom') || 1440), left: +(attr(pgMar, 'left') || 1440),
    }
  }
  return sec
}

export interface ReadResult { doc: Doc; pkg: Package; part: string }

export async function readDocx(bytes: Uint8Array | ArrayBuffer): Promise<ReadResult> {
  const pkg = await Package.open(bytes)
  const part = pkg.main || 'word/document.xml'
  const root = pkg.xml(part)
  const body = kid(root, 'body') || root
  const blocks: Block[] = []
  for (const child of body.children) {
    if (child.localName === 'p') blocks.push(readPara(child, pkg, part))
    else if (child.localName === 'tbl') blocks.push(readTable(child, pkg, part))
    else if (child.localName === 'sectPr') { /* handled */ }
    else blocks.push({ ...emptyPara(), raw: serializeRaw(child) })
  }
  if (!blocks.length) blocks.push(emptyPara())
  const { styles, defaultFont, defaultSize } = readStyles(pkg)
  const doc: Doc = {
    blocks, section: readSection(body), styles, defaultFont, defaultSize,
    numbering: readNumbering(pkg),
    format: pkg.variant('docx'),
  }
  return { doc, pkg, part }
}
