/* model → PPTX. Regenerates slide parts; keeps other source parts. A brand-new deck gets a
 * minimal but Office-valid master / layout / theme / presentation part set. */
import { Package, NS, formats, newPackage, type FormatInfo } from '@voffice/react-shared'
import { esc, escText } from '@voffice/react-shared'
import { type Presentation, type Slide, type Shape, type Para } from './model'

const P = NS.p, A = NS.a, R = NS.r
const SLIDE_NS = `xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`
const hex = (c: string): string => c.replace('#', '').toUpperCase()

function fillXml(f: Shape['fill']): string {
  if (f === null) return '<a:noFill/>'
  if (f && f.color) return `<a:solidFill><a:srgbClr val="${hex(f.color)}"/></a:solidFill>`
  return ''
}
function lineXml(l: Shape['line']): string {
  if (l === null) return '<a:ln><a:noFill/></a:ln>'
  if (l && l.color) return `<a:ln w="${Math.round((l.width || 1) * 12700)}"><a:solidFill><a:srgbClr val="${hex(l.color)}"/></a:solidFill>${l.dash ? `<a:prstDash val="${l.dash}"/>` : ''}</a:ln>`
  return ''
}
function rPrXml(r: Para['runs'][number]['rPr']): string {
  if (!r) return ''
  const attrs: string[] = ['lang="en-US"']
  if (r.size) attrs.push(`sz="${Math.round(r.size * 100)}"`)
  if (r.bold) attrs.push('b="1"')
  if (r.italic) attrs.push('i="1"')
  if (r.underline) attrs.push('u="sng"')
  let inner = ''
  if (r.color) inner += `<a:solidFill><a:srgbClr val="${hex(r.color)}"/></a:solidFill>`
  if (r.font) inner += `<a:latin typeface="${esc(r.font)}"/>`
  return `<a:rPr ${attrs.join(' ')} dirty="0">${inner}</a:rPr>`
}
function paraXml(p: Para): string {
  const pPr = p.align || p.bullet ? `<a:pPr${p.align ? ` algn="${p.align === 'center' ? 'ctr' : p.align === 'right' ? 'r' : 'l'}"` : ''}>${p.bullet ? '<a:buChar char="•"/>' : '<a:buNone/>'}</a:pPr>` : ''
  const runs = p.runs.map((r) => `<a:r>${rPrXml(r.rPr)}<a:t>${escText(r.text)}</a:t></a:r>`).join('')
  return `<a:p>${pPr}${runs || '<a:endParaRPr lang="en-US"/>'}</a:p>`
}
function txBodyXml(text: Para[] | undefined): string {
  const paras = (text && text.length ? text : [{ runs: [] } as Para]).map(paraXml).join('')
  return `<p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/>${paras}</p:txBody>`
}
function xfrmXml(s: Shape): string {
  return `<a:xfrm${s.rot ? ` rot="${Math.round(s.rot * 60000)}"` : ''}><a:off x="${Math.round(s.x)}" y="${Math.round(s.y)}"/><a:ext cx="${Math.round(s.w)}" cy="${Math.round(s.h)}"/></a:xfrm>`
}

let shapeId = 1
function shapeXml(s: Shape, pkg: Package, part: string): string {
  const id = ++shapeId
  if (s.raw != null) return s.raw
  if (s.kind === 'pic' && s.image) {
    const rId = s.image.rId
    return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(s.name || 'Picture')}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr>${xfrmXml(s)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
  }
  const tag = s.kind === 'cxnSp' ? 'p:cxnSp' : 'p:sp'
  const nvTag = s.kind === 'cxnSp' ? 'p:nvCxnSpPr' : 'p:nvSpPr'
  const cNvTag = s.kind === 'cxnSp' ? 'p:cNvCxnSpPr' : 'p:cNvSpPr'
  const geom = `<a:prstGeom prst="${esc(s.preset || 'rect')}"><a:avLst/></a:prstGeom>`
  const tx = s.kind === 'cxnSp' ? '' : txBodyXml(s.text)
  return `<${tag}><${nvTag}><p:cNvPr id="${id}" name="${esc(s.name || 'Shape')}"/><${cNvTag}/><p:nvPr/></${nvTag}>` +
    `<p:spPr>${xfrmXml(s)}${geom}${fillXml(s.fill)}${lineXml(s.line)}</p:spPr>${tx}</${tag}>`
}

function slideXml(slide: Slide, pkg: Package, part: string): string {
  shapeId = 1
  const bg = slide.background ? `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex(slide.background)}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` : ''
  const shapes = slide.shapes.map((s) => shapeXml(s, pkg, part)).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld ${SLIDE_NS}><p:cSld>${bg}<p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    shapes + `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
}

/* ---------- minimal new-deck parts ---------- */
const THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:theme xmlns:a="${A}" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`

const MASTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldMaster ${SLIDE_NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`

const LAYOUT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sldLayout ${SLIDE_NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`

const CT = {
  pres: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  master: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  layout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
}

export interface PptxWriteOpts { pkg?: Package; slideParts?: string[]; format?: string; newImages?: { shapeId: string; part: string; rId: string; bytes: Uint8Array }[] }

export async function writePptx(pres: Presentation, opts: PptxWriteOpts = {}): Promise<Blob> {
  const fmt: FormatInfo = formats[opts.format || pres.format || 'pptx'] || formats.pptx
  const pkg = opts.pkg || newPackage('ppt/presentation.xml', '', fmt)
  const parts = opts.slideParts && opts.slideParts.length === pres.slides.length ? opts.slideParts : pres.slides.map((_, i) => `ppt/slides/slide${i + 1}.xml`)
  // register picture bytes for any image not already stored in the package
  for (let i = 0; i < pres.slides.length; i++) {
    const part = parts[i]
    for (const sh of pres.slides[i].shapes) {
      if (sh.kind === 'pic' && sh.image && sh.image.url && (!sh.image.part || !pkg.has(sh.image.part))) {
        const res = await fetch(sh.image.url)
        const bytes = new Uint8Array(await res.arrayBuffer())
        const ext = (sh.image.part.split('.').pop() || 'png').toLowerCase()
        const name = pkg.nextMediaName(ext === 'jpg' ? 'jpeg' : ext)
        const mediaPart = 'ppt/media/' + name
        pkg.addPart(mediaPart, bytes)
        sh.image.rId = pkg.addRel(part, NS.rel + '/image', '../media/' + name)
        sh.image.part = mediaPart
      }
    }
  }
  pres.slides.forEach((s, i) => {
    const part = parts[i]
    pkg.addPart(part, slideXml(s, pkg, part), CT.slide)
    if (!opts.pkg) pkg.addRel(part, NS.rel + '/slideLayout', '../slideLayouts/slideLayout1.xml')
  })
  if (!opts.pkg) {
    pkg.addPart('ppt/theme/theme1.xml', THEME, CT.theme)
    pkg.addPart('ppt/slideMasters/slideMaster1.xml', MASTER, CT.master)
    pkg.addPart('ppt/slideLayouts/slideLayout1.xml', LAYOUT, CT.layout)
    pkg.addRel('ppt/slideMasters/slideMaster1.xml', NS.rel + '/slideLayout', '../slideLayouts/slideLayout1.xml')
    pkg.addRel('ppt/slideMasters/slideMaster1.xml', NS.rel + '/theme', '../theme/theme1.xml')
    pkg.addRel('ppt/slideLayouts/slideLayout1.xml', NS.rel + '/slideMaster', '../slideMasters/slideMaster1.xml')
    const sldIds = pres.slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('')
    pkg.setPart('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation ${SLIDE_NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${sldIds}</p:sldIdLst><p:sldSz cx="${pres.slideW}" cy="${pres.slideH}"/><p:notesSz cx="${pres.slideH}" cy="${pres.slideW}"/></p:presentation>`)
    const rels: string[] = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    rels.push(`<Relationship Id="rId1" Type="${NS.rel}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`)
    pres.slides.forEach((_, i) => rels.push(`<Relationship Id="rId${i + 2}" Type="${NS.rel}/slide" Target="slides/slide${i + 1}.xml"/>`))
    rels.push('</Relationships>')
    pkg.setPart('ppt/_rels/presentation.xml.rels', rels.join(''))
    pkg.addRel('', NS.rel + '/officeDocument', 'ppt/presentation.xml')
  } else {
    // keep presentation unless slide count changed
    const srcCount = countSlides(opts.pkg)
    if (srcCount !== pres.slides.length) {
      const sldIds = pres.slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('')
      // rebuild sldIdLst only, keep rest of presentation.xml
      const root = pkg.xml('ppt/presentation.xml')
      const lst = root.children.find((c) => c.localName === 'sldIdLst')
      if (lst) { lst.childNodes = []; lst.children = [] }
      pkg.setPart('ppt/presentation.xml', rebuildPresentation(pkg, pres))
      const rels: string[] = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
      rels.push(`<Relationship Id="rId1" Type="${NS.rel}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`)
      pres.slides.forEach((_, i) => rels.push(`<Relationship Id="rId${i + 2}" Type="${NS.rel}/slide" Target="slides/slide${i + 1}.xml"/>`))
      rels.push('</Relationships>')
      pkg.setPart('ppt/_rels/presentation.xml.rels', rels.join(''))
    }
  }
  return pkg.save(fmt.mime)
}

function countSlides(pkg: Package): number {
  if (!pkg.has('ppt/presentation.xml')) return 0
  const root = pkg.xml('ppt/presentation.xml')
  const lst = root.children.find((c) => c.localName === 'sldIdLst')
  return lst ? lst.children.length : 0
}
function rebuildPresentation(pkg: Package, pres: Presentation): string {
  const sldIds = pres.slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:presentation ${SLIDE_NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${sldIds}</p:sldIdLst><p:sldSz cx="${pres.slideW}" cy="${pres.slideH}"/><p:notesSz cx="${pres.slideH}" cy="${pres.slideW}"/></p:presentation>`
}
