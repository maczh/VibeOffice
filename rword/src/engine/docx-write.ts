/* model → DOCX. Regenerates the main document part; every other part of the source
 * package is written back unchanged. New pictures/hyperlinks get parts/relationships. */
import { Package, NS, formats, newPackage, xmlPart, type FormatInfo } from '@voffice/react-shared'
import { esc, escText } from '@voffice/react-shared'
import {
  type Doc, type Para, type Run, type RPr, type PPr, type Table, type Block,
  isPara, isTable, defaultSection,
} from './model'

const TWIP = 1440 / 96
const EMU = 914400 / 96
const tw = (px: number): number => Math.round(px * TWIP)

const ROOT_NS = 'xmlns:w="' + NS.w + '" xmlns:r="' + NS.r + '" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="' + NS.a + '" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'

function rPrXml(r: RPr | undefined): string {
  if (!r) return ''
  let s = ''
  if (r.font) s += `<w:rFonts w:ascii="${esc(r.font)}" w:hAnsi="${esc(r.font)}" w:cs="${esc(r.font)}"/>`
  if (r.bold) s += '<w:b/>'
  if (r.italic) s += '<w:i/>'
  if (r.strike) s += '<w:strike/>'
  if (r.underline) s += `<w:u w:val="${r.underlineStyle || 'single'}"/>`
  if (r.caps === 'all') s += '<w:caps/>'
  if (r.caps === 'small') s += '<w:smallCaps/>'
  if (r.hidden) s += '<w:vanish/>'
  if (r.color) s += `<w:color w:val="${r.color.replace('#', '').toUpperCase()}"/>`
  if (r.size) s += `<w:sz w:val="${Math.round(r.size * 2)}"/><w:szCs w:val="${Math.round(r.size * 2)}"/>`
  if (r.highlight) s += `<w:highlight w:val="${esc(r.highlight)}"/>`
  if (r.vertAlign) s += `<w:vertAlign w:val="${r.vertAlign}"/>`
  return s ? `<w:rPr>${s}</w:rPr>` : ''
}

function pPrXml(p: PPr | undefined): string {
  if (!p) return ''
  let s = ''
  if (p.pageBreakBefore) s += '<w:pageBreakBefore/>'
  if (p.style) s += `<w:pStyle w:val="${esc(p.style)}"/>`
  if (p.numId != null) s += `<w:numPr><w:ilvl w:val="${p.ilvl || 0}"/><w:numId w:val="${p.numId}"/></w:numPr>`
  if (p.spaceBefore != null || p.spaceAfter != null || p.line != null) {
    const before = p.spaceBefore != null ? ` w:before="${tw(p.spaceBefore)}"` : ''
    const after = p.spaceAfter != null ? ` w:after="${tw(p.spaceAfter)}"` : ''
    let line = ''
    if (p.line != null) line = p.line > 4 ? ` w:line="${tw(p.line)}" w:lineRule="exact"` : ` w:line="${Math.round(p.line * 240)}" w:lineRule="auto"`
    s += `<w:spacing${before}${after}${line}/>`
  }
  if (p.indent != null || p.hanging != null) {
    const left = p.indent != null ? ` w:left="${tw(p.indent)}"` : ''
    const hang = p.hanging != null ? ` w:hanging="${tw(p.hanging)}"` : ''
    s += `<w:ind${left}${hang}/>`
  }
  if (p.align) s += `<w:jc w:val="${p.align === 'justify' ? 'both' : p.align}"/>`
  if (p.keepNext) s += '<w:keepNext/>'
  if (p.keepLines) s += '<w:keepLines/>'
  if (p.outlineLvl != null) s += `<w:outlineLvl w:val="${p.outlineLvl}"/>`
  if (p.borders) s += '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>'
  if (p.shading) s += `<w:shd w:val="clear" w:color="auto" w:fill="${p.shading.replace('#', '').toUpperCase()}"/>`
  return s ? `<w:pPr>${s}</w:pPr>` : ''
}

interface WriteCtx {
  pkg: Package
  part: string
  picId: number
  linkIds: Map<string, string>
  newImages: { url: string; part: string; rId: string }[]
}

function ensureLink(ctx: WriteCtx, url: string, rId: string): string {
  if (rId && ctx.pkg.rel(ctx.part, rId)) return rId
  const existing = ctx.pkg.relsOf(ctx.part).find((r) => r.target === url && r.mode === 'External')
  if (existing) return existing.id
  const key = url
  if (ctx.linkIds.has(key)) return ctx.linkIds.get(key)!
  const id = ctx.pkg.addRel(ctx.part, NS.rel + '/hyperlink', url, 'External')
  ctx.linkIds.set(key, id)
  return id
}

async function ensureImage(ctx: WriteCtx, img: NonNullable<Run['image']>): Promise<string> {
  if (img.part && ctx.pkg.has(img.part) && img.rId && ctx.pkg.rel(ctx.part, img.rId)) return img.rId
  // a picture added while editing: fetch its bytes and store a new media part
  const res = await fetch(img.url)
  const bytes = new Uint8Array(await res.arrayBuffer())
  const ext = (img.part.split('.').pop() || 'png').toLowerCase()
  const name = ctx.pkg.nextMediaName(ext === 'jpg' ? 'jpeg' : ext)
  const folder = ctx.part.startsWith('word/') ? 'word/media/' : 'media/'
  const part = folder + name
  ctx.pkg.addPart(part, bytes)
  const rId = ctx.pkg.addRel(ctx.part, NS.rel + '/image', 'media/' + name)
  img.part = part
  img.rId = rId
  return rId
}

async function runXml(ctx: WriteCtx, r: Run): Promise<string> {
  if (r.raw != null) return r.raw
  const rPr = rPrXml(r.rPr)
  if (r.image) {
    const rId = await ensureImage(ctx, r.image)
    const cx = Math.round(r.image.w * EMU); const cy = Math.round(r.image.h * EMU)
    const id = ++ctx.picId
    return `<w:r>${rPr}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
      `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
  }
  let body = ''
  if (r.br) body += '<w:br/>'
  if (r.tab) body += '<w:tab/>'
  if (r.text != null) {
    const preserve = /^\s|\s$/.test(r.text) ? ' xml:space="preserve"' : ''
    body += `<w:t${preserve}>${escText(r.text)}</w:t>`
  }
  if (!body) body = '<w:t xml:space="preserve"></w:t>'
  const inner = `<w:r>${rPr}${body}</w:r>`
  if (r.hyperlink) {
    const id = ensureLink(ctx, r.hyperlink.url, r.hyperlink.rId)
    return `<w:hyperlink r:id="${id}">${inner}</w:hyperlink>`
  }
  return inner
}

async function paraXml(ctx: WriteCtx, p: Para): Promise<string> {
  if (p.raw != null && !p.runs.length) return p.raw
  const runs: string[] = []
  for (const r of p.runs) runs.push(await runXml(ctx, r))
  return `<w:p>${pPrXml(p.pPr)}${runs.join('')}</w:p>`
}

async function tableXml(ctx: WriteCtx, t: Table): Promise<string> {
  const grid = t.widths.map((w) => `<w:gridCol w:w="${tw(w)}"/>`).join('')
  const borders = t.rows.some((row) => row.some((c) => c.borders === false)) ? '' : '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') +
    '</w:tblBorders>'
  const jc = t.align ? `<w:jc w:val="${t.align}"/>` : ''
  const layout = t.autoFit === 'content' ? '<w:tblLayout w:type="autofit"/>' : t.autoFit === 'fixed' ? '<w:tblLayout w:type="fixed"/>' : ''
  const rows: string[] = []
  for (let ri = 0; ri < t.rows.length; ri++) {
    const row = t.rows[ri]
    const cells: string[] = []
    for (const c of row) {
      let tcPr = `<w:tcW w:w="${tw(c.width || 100)}" w:type="dxa"/>`
      if (c.gridSpan && c.gridSpan > 1) tcPr += `<w:gridSpan w:val="${c.gridSpan}"/>`
      if (c.vMerge) tcPr += c.vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>'
      if (c.fill) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${c.fill.replace('#', '').toUpperCase()}"/>`
      if (c.vAlign) tcPr += `<w:vAlign w:val="${c.vAlign === 'center' ? 'center' : c.vAlign === 'bottom' ? 'bottom' : 'top'}"/>`
      if (c.textDirection === 'vertical') tcPr += '<w:textDirection w:val="tbRl"/>'
      const paras: string[] = []
      for (const p of c.paras) paras.push(await paraXml(ctx, p))
      cells.push(`<w:tc><w:tcPr>${tcPr}</w:tcPr>${paras.join('')}</w:tc>`)
    }
    const trPr = ri === 0 && t.headerRows ? '<w:trPr><w:tblHeader/></w:trPr>' : ''
    rows.push(`<w:tr>${trPr}${cells.join('')}</w:tr>`)
  }
  return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}${jc}${layout}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows.join('')}</w:tbl>`
}

function sectXml(doc: Doc): string {
  const s = doc.section || defaultSection()
  const m = s.margins
  return `<w:sectPr><w:pgSz w:w="${s.pageW}" w:h="${s.pageH}"/>` +
    `<w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="720" w:footer="720" w:gutter="0"/>` +
    `<w:cols w:num="${Math.max(1, s.columns || 1)}" w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>`
}

async function bodyXml(ctx: WriteCtx, doc: Doc): Promise<string> {
  const out: string[] = []
  for (const b of doc.blocks) {
    if (isPara(b)) out.push(await paraXml(ctx, b))
    else if (isTable(b)) out.push(await tableXml(ctx, b))
  }
  return out.join('') + sectXml(doc)
}

const STYLES_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'
const SETTINGS_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'

function stylesXml(doc: Doc): string {
  const styles = doc.styles.map((s) => {
    let inner = `<w:name w:val="${esc(s.name)}"/>`
    if (s.basedOn) inner += `<w:basedOn w:val="${esc(s.basedOn)}"/>`
    inner += '<w:qFormat/>'
    if (s.pPr) inner += pPrXml(s.pPr).replace('<w:pPr>', '<w:pPr>').replace('</w:pPr>', '</w:pPr>')
    if (s.rPr) inner += rPrXml(s.rPr)
    return `<w:style w:type="paragraph" w:styleId="${esc(s.id)}">${inner}</w:style>`
  }).join('')
  const defRPr = `<w:rPr><w:rFonts w:ascii="${esc(doc.defaultFont)}" w:hAnsi="${esc(doc.defaultFont)}"/><w:sz w:val="${Math.round(doc.defaultSize * 2)}"/></w:rPr>`
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:styles ${ROOT_NS}>` +
    `<w:docDefaults><w:rPrDefault>${defRPr}</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
    styles + '</w:styles>'
}

const SETTINGS_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<w:settings ${ROOT_NS}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/></w:settings>`

export interface WriteOpts { pkg?: Package; part?: string; format?: string }

export async function writeDocx(doc: Doc, opts: WriteOpts = {}): Promise<Blob> {
  const fmt: FormatInfo = formats[opts.format || doc.format || 'docx'] || formats.docx
  const part = opts.part || opts.pkg?.main || 'word/document.xml'
  const pkg = opts.pkg || newPackage(part, '', fmt)
  const ctx: WriteCtx = { pkg, part, picId: 100, linkIds: new Map(), newImages: [] }
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document ${ROOT_NS}><w:body>${await bodyXml(ctx, doc)}</w:body></w:document>`
  pkg.setPart(part, xml)
  if (!opts.pkg) {
    // a brand-new document: add the styles and settings parts Word expects
    pkg.addPart('word/styles.xml', stylesXml(doc), STYLES_CT)
    pkg.addRel(part, NS.rel + '/styles', 'styles.xml')
    pkg.addPart('word/settings.xml', SETTINGS_XML, SETTINGS_CT)
    pkg.addRel(part, NS.rel + '/settings', 'settings.xml')
  } else if (!pkg.has('word/styles.xml')) {
    pkg.addPart('word/styles.xml', stylesXml(doc), STYLES_CT)
    pkg.addRel(part, NS.rel + '/styles', 'styles.xml')
  }
  return pkg.save(fmt.mime)
}
