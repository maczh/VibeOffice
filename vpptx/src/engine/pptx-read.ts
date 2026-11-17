/* PPTX → model. Reads presentation + slides; images become blob URLs; unsupported shapes keep raw XML. */
import { Package, NS } from '@voffice/shared'
import type { El } from '@voffice/shared'
import { type Presentation, type Slide, type Shape, type Para, type Run, type RPr, newSlide, uid } from './model'

const P = NS.p, A = NS.a, R = NS.r
const kids = (el: El, l: string): El[] => el.children.filter((c) => c.localName === l)
const kid = (el: El, l: string): El | undefined => el.children.find((c) => c.localName === l)
const at = (el: El | undefined, n: string): string | null => (el ? el.getAttribute(n) : null)
const deep = (el: El, l: string): El | undefined => el.getElementsByTagNameNS('*', l)[0]

function serializeRaw(el: El): string {
  const attrs = el.attributes.map((a) => ` ${a.name}="${a.value.replace(/"/g, '&quot;')}"`).join('')
  const inner = el.childNodes.map((n) => n.nodeType === 3 ? (n as { data: string }).data.replace(/&/g, '&amp;').replace(/</g, '&lt;') : serializeRaw(n as El)).join('')
  return `<${el.nodeName}${attrs}>${inner}</${el.nodeName}>`
}

function readRPr(el: El | undefined): RPr | undefined {
  if (!el) return undefined
  const r: RPr = {}
  const sz = at(el, 'sz'); if (sz) r.size = +sz / 100
  if (at(el, 'b') === '1') r.bold = true
  if (at(el, 'i') === '1') r.italic = true
  if (at(el, 'u') && at(el, 'u') !== 'none') r.underline = true
  const fill = kid(el, 'solidFill')
  const clr = fill && (kid(fill, 'srgbClr') || kid(fill, 'schemeClr'))
  if (clr && kid(fill, 'srgbClr')) r.color = ('#' + at(clr, 'val')).toLowerCase()
  const latin = kid(el, 'latin'); if (latin) r.font = at(latin, 'typeface') || undefined
  return Object.keys(r).length ? r : undefined
}

function readText(txBody: El | undefined): Para[] {
  if (!txBody) return []
  const paras: Para[] = []
  for (const p of kids(txBody, 'p')) {
    const para: Para = { runs: [] }
    const pPr = kid(p, 'pPr')
    const algn = at(pPr, 'algn')
    if (algn) para.align = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : 'left'
    if (pPr && (kid(pPr, 'buChar') || kid(pPr, 'buAutoNum'))) para.bullet = true
    for (const r of kids(p, 'r')) {
      const t = kid(r, 't')
      para.runs.push({ text: t?.textContent || '', rPr: readRPr(kid(r, 'rPr')) })
    }
    const endPr = kid(p, 'endParaRPr')
    if (!para.runs.length && endPr) para.runs.push({ text: '', rPr: readRPr(endPr) })
    paras.push(para)
  }
  return paras
}

function readXfrm(spPr: El | undefined): { x: number; y: number; w: number; h: number; rot?: number } {
  const xfrm = spPr && kid(spPr, 'xfrm')
  const off = xfrm && kid(xfrm, 'off'); const ext = xfrm && kid(xfrm, 'ext')
  return {
    x: +(at(off, 'x') || 0), y: +(at(off, 'y') || 0),
    w: +(at(ext, 'cx') || 0), h: +(at(ext, 'cy') || 0),
    rot: xfrm && at(xfrm, 'rot') ? +at(xfrm, 'rot')! / 60000 : undefined,
  }
}

function readFillLine(spPr: El | undefined): { fill: Shape['fill']; line: Shape['line'] } {
  let fill: Shape['fill'] = undefined
  let line: Shape['line'] = undefined
  if (!spPr) return { fill, line }
  if (kid(spPr, 'noFill')) fill = null
  const sf = kid(spPr, 'solidFill')
  if (sf) { const c = kid(sf, 'srgbClr'); if (c) fill = { color: ('#' + at(c, 'val')).toLowerCase() } }
  const ln = kid(spPr, 'ln')
  if (ln) {
    if (kid(ln, 'noFill')) line = null
    else {
      const lsf = kid(ln, 'solidFill'); const lc = lsf && kid(lsf, 'srgbClr')
      line = { color: lc ? ('#' + at(lc, 'val')).toLowerCase() : '#000000', width: +(at(ln, 'w') || 12700) / 12700 }
      const dash = kid(ln, 'prstDash'); if (dash) line.dash = at(dash, 'val') || undefined
    }
  }
  return { fill, line }
}

function readShape(el: El, pkg: Package, part: string): Shape | null {
  const tag = el.localName
  const nv = kid(el, 'nvSpPr') || kid(el, 'nvPicPr') || kid(el, 'nvCxnSpPr') || kid(el, 'nvGraphicFramePr')
  const cNvPr = nv && kid(nv, 'cNvPr')
  const name = at(cNvPr, 'name') || ''
  const spPr = kid(el, 'spPr')
  const geo = readXfrm(spPr)
  if (tag === 'pic') {
    const blip = deep(el, 'blip')
    const rId = blip ? (blip.getAttributeNS(R, 'embed') || at(blip, 'r:embed') || '') : ''
    const rel = rId ? pkg.rel(part, rId) : undefined
    let image: Shape['image']
    if (rel) {
      const target = rel.target.startsWith('/') ? rel.target.slice(1) : resolve(part, rel.target)
      const bytes = pkg.bytes(target)
      if (bytes) {
        const ext = target.split('.').pop() || 'png'
        const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : 'image/' + ext
        image = { part: target, rId, url: URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime })) }
      }
    }
    return { id: uid('sh'), kind: 'pic', ...geo, image, name }
  }
  if (tag === 'sp' || tag === 'cxnSp') {
    const prst = spPr && kid(spPr, 'prstGeom')
    const { fill, line } = readFillLine(spPr)
    const shape: Shape = {
      id: uid('sh'), kind: tag === 'cxnSp' ? 'cxnSp' : 'sp', ...geo,
      preset: at(prst, 'prst') || (tag === 'cxnSp' ? 'line' : 'rect'),
      fill, line, name,
      text: readText(kid(el, 'txBody')),
    }
    return shape
  }
  // graphicFrame (table/chart) or unknown: keep raw
  return { id: uid('sh'), kind: 'graphicFrame', ...geo, raw: serializeRaw(el), name }
}

function resolve(part: string, target: string): string {
  const base = part.includes('/') ? part.slice(0, part.lastIndexOf('/') + 1) : ''
  const out: string[] = []
  for (const s of (base + target).split('/')) { if (s === '.' || s === '') continue; if (s === '..') out.pop(); else out.push(s) }
  return out.join('/')
}

export interface PptxReadResult { pres: Presentation; pkg: Package; slideParts: string[] }

export async function readPptx(bytes: Uint8Array | ArrayBuffer): Promise<PptxReadResult> {
  const pkg = await Package.open(bytes)
  const presRoot = pkg.xml('ppt/presentation.xml')
  const sldSz = kid(presRoot, 'sldSz')
  const pres: Presentation = {
    slideW: +(at(sldSz, 'cx') || 12192000), slideH: +(at(sldSz, 'cy') || 6858000),
    slides: [], format: pkg.variant('pptx'),
  }
  const rels = pkg.relsOf('ppt/presentation.xml')
  const sldIdLst = kid(presRoot, 'sldIdLst')
  const slideParts: string[] = []
  const ids = sldIdLst ? kids(sldIdLst, 'sldId') : []
  for (const sid of ids) {
    const rid = sid.getAttributeNS(R, 'id') || at(sid, 'r:id') || ''
    const rel = rels.find((r) => r.id === rid)
    if (!rel) continue
    const part = rel.target.startsWith('/') ? rel.target.slice(1) : resolve('ppt/presentation.xml', rel.target)
    slideParts.push(part)
    const slide = newSlide()
    slide.part = part
    if (pkg.has(part)) {
      const root = pkg.xml(part)
      const cSld = kid(root, 'cSld')
      const bg = cSld && kid(cSld, 'bg')
      const bgFill = bg && deep(bg, 'srgbClr')
      if (bgFill) slide.background = ('#' + at(bgFill, 'val')).toLowerCase()
      const spTree = cSld && kid(cSld, 'spTree')
      if (spTree) for (const child of spTree.children) {
        if (['sp', 'pic', 'cxnSp', 'graphicFrame'].includes(child.localName)) {
          const sh = readShape(child, pkg, part)
          if (sh) slide.shapes.push(sh)
        }
      }
      // notes
      const notesRel = pkg.relsOf(part).find((r) => r.type.endsWith('/notesSlide'))
      if (notesRel) {
        const np = resolve(part, notesRel.target)
        if (pkg.has(np)) {
          const nroot = pkg.xml(np)
          const tx = deep(nroot, 'txBody')
          if (tx) slide.notes = kids(tx, 'p').map((p) => kids(p, 'r').map((r) => kid(r, 't')?.textContent || '').join('')).join('\n')
        }
      }
    }
    pres.slides.push(slide)
  }
  if (!pres.slides.length) { pres.slides.push(newSlide()); slideParts.push('ppt/slides/slide1.xml') }
  return { pres, pkg, slideParts }
}