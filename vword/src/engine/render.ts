/* model ↔ HTML for the contenteditable surface. renderBlock produces the editing DOM;
 * syncBlock reads the DOM back into the model, computing each run's formatting from its
 * ancestor chain (so execCommand output round-trips). */
import {
  type Doc, type Block, type Para, type Run, type RPr, type Table, type Cell,
  isPara, emptyPara, uid,
} from './model'

const HL_NAMES: Record<string, string> = {
  '#ffff00': 'yellow', '#00ff00': 'green', '#00ffff': 'cyan', '#ff00ff': 'magenta',
  '#0000ff': 'blue', '#ff0000': 'red', '#000080': 'darkBlue', '#008080': 'darkCyan',
  '#008000': 'darkGreen', '#800080': 'darkMagenta', '#800000': 'darkRed', '#808000': 'darkYellow',
  '#808080': 'darkGray', '#c0c0c0': 'lightGray', '#000000': 'black', '#ffffff': 'white',
}

const escHtml = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

function rPrStyle(r: RPr | undefined, doc: Doc): string {
  const s: string[] = []
  const font = r?.font || doc.defaultFont
  const size = r?.size || doc.defaultSize
  s.push(`font-family:${font},Carlito,Calibri,sans-serif`)
  s.push(`font-size:${size}pt`)
  if (r?.bold) s.push('font-weight:bold')
  if (r?.italic) s.push('font-style:italic')
  const td: string[] = []
  if (r?.underline) td.push('underline')
  if (r?.strike) td.push('line-through')
  if (td.length) s.push('text-decoration:' + td.join(' '))
  if (r?.color) s.push('color:' + r.color)
  if (r?.highlight) s.push('background:' + r.highlight)
  if (r?.vertAlign === 'superscript') s.push('vertical-align:super;font-size:smaller')
  if (r?.vertAlign === 'subscript') s.push('vertical-align:sub;font-size:smaller')
  return s.join(';')
}

function pPrStyle(p: Para, doc: Doc): string {
  const s: string[] = []
  const st = p.pPr?.style ? doc.styles.find((x) => x.id === p.pPr!.style) : undefined
  const align = p.pPr?.align || st?.pPr?.align
  if (align) s.push('text-align:' + align)
  const before = p.pPr?.spaceBefore ?? st?.pPr?.spaceBefore
  const after = p.pPr?.spaceAfter ?? st?.pPr?.spaceAfter
  if (before != null) s.push('margin-top:' + before + 'px')
  if (after != null) s.push('margin-bottom:' + after + 'px')
  const indent = p.pPr?.indent ?? 0
  if (indent) s.push('padding-left:' + indent + 'px')
  if (p.pPr?.line) s.push(p.pPr.line > 4 ? 'line-height:' + p.pPr.line + 'px' : 'line-height:' + p.pPr.line)
  const rPr = st?.rPr
  if (rPr) s.push(rPrStyle(rPr, doc))
  return s.join(';')
}

function runHtml(r: Run, doc: Doc): string {
  if (r.raw != null) return `<span class="raw" contenteditable="false" data-raw="${encodeURIComponent(r.raw)}">[object]</span>`
  const style = rPrStyle(r.rPr, doc)
  if (r.image) {
    return `<img class="img" src="${r.image.url}" width="${r.image.w}" height="${r.image.h}" data-part="${r.image.part}" data-rid="${r.image.rId}" data-w="${r.image.w}" data-h="${r.image.h}" style="${style}">`
  }
  if (r.br) return '<br>'
  if (r.tab) return `<span class="tb" style="${style}">\t</span>`
  const text = escHtml(r.text ?? '')
  const span = `<span style="${style}">${text || ''}</span>`
  if (r.hyperlink) return `<a href="${escHtml(r.hyperlink.url)}" data-rid="${r.hyperlink.rId}" class="lnk">${span}</a>`
  return span
}

export function renderBlock(b: Block, doc: Doc): string {
  if (isPara(b)) {
    const inner = b.runs.map((r) => runHtml(r, doc)).join('')
    const marker = listMarker(b, doc)
    return `<div class="p" data-id="${b.id}" style="${pPrStyle(b, doc)}">${marker}${inner || '<br>'}</div>`
  }
  const t = b as Table
  const cols = t.widths.map((w) => `<col style="width:${w}px">`).join('')
  const rows = t.rows.map((row) => {
    const cells = row.map((c) => {
      const span = c.gridSpan && c.gridSpan > 1 ? ` colspan="${c.gridSpan}"` : ''
      const vm = c.vMerge === 'continue' ? ' data-vmerge="continue"' : c.vMerge === 'restart' ? ' data-vmerge="restart"' : ''
      const fill = c.fill ? `background:${c.fill};` : ''
      const paras = c.paras.map((p) => renderBlock(p, doc)).join('')
      return `<td${span}${vm} data-w="${c.width || 0}" style="${fill}border:1px solid #888;padding:2px 4px;vertical-align:top">${paras}</td>`
    }).join('')
    return `<tr>${cells}</tr>`
  }).join('')
  return `<table class="tbl" data-id="${t.id}" style="border-collapse:collapse;table-layout:fixed"><colgroup>${cols}</colgroup>${rows}</table>`
}

function listMarker(p: Para, doc: Doc): string {
  if (p.pPr?.numId == null) return ''
  // numbering is resolved by the component (needs sequence context); a bullet is static
  return `<span class="lm" contenteditable="false" data-num="${p.pPr.numId}" data-ilvl="${p.pPr.ilvl || 0}"></span>`
}

/* ---------- DOM → model ---------- */

function rgbToHex(c: string): string | undefined {
  if (!c) return undefined
  if (c.startsWith('#')) return c.length === 7 ? c : undefined
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c)
  if (!m) return undefined
  const h = (n: string) => (+n).toString(16).padStart(2, '0')
  return '#' + h(m[1]) + h(m[2]) + h(m[3])
}

function mergeStyle(r: RPr, el: HTMLElement): RPr {
  const out: RPr = { ...r }
  const tag = el.tagName.toLowerCase()
  if (tag === 'b' || tag === 'strong') out.bold = true
  if (tag === 'i' || tag === 'em') out.italic = true
  if (tag === 'u') out.underline = true
  if (tag === 's' || tag === 'strike' || tag === 'del') out.strike = true
  const st = el.style
  if (st.fontWeight && (st.fontWeight === 'bold' || +st.fontWeight >= 600)) out.bold = true
  if (st.fontWeight && (st.fontWeight === 'normal' || +st.fontWeight < 600)) out.bold = false
  if (st.fontStyle === 'italic') out.italic = true
  if (st.fontStyle === 'normal') out.italic = false
  const td = st.textDecorationLine || st.textDecoration || ''
  if (td.includes('underline')) out.underline = true
  if (td.includes('line-through')) out.strike = true
  if (st.fontFamily) out.font = st.fontFamily.split(',')[0].replace(/["']/g, '').trim()
  if (st.fontSize) {
    const px = parseFloat(st.fontSize)
    if (st.fontSize.endsWith('pt')) out.size = px
    else if (!isNaN(px)) out.size = Math.round(px * 0.75 * 2) / 2
  }
  const col = rgbToHex(st.color)
  if (col) out.color = col
  const bg = rgbToHex(st.backgroundColor)
  if (bg && bg !== '#ffffff') out.highlight = HL_NAMES[bg.toLowerCase()] || bg
  if (st.verticalAlign === 'super') out.vertAlign = 'superscript'
  if (st.verticalAlign === 'sub') out.vertAlign = 'subscript'
  return out
}

function syncRuns(node: Node, rPr: RPr, link: { rId: string; url: string } | undefined, out: Run[]): void {
  if (node.nodeType === 3) {
    const text = node.nodeValue || ''
    if (text) out.push({ text, rPr: cleanRPr(rPr), hyperlink: link })
    return
  }
  if (node.nodeType !== 1) return
  const el = node as HTMLElement
  const tag = el.tagName.toLowerCase()
  if (tag === 'br') { out.push({ br: true, rPr: cleanRPr(rPr) }); return }
  if (tag === 'img') {
    out.push({
      image: {
        url: el.getAttribute('src') || '', part: el.getAttribute('data-part') || '',
        rId: el.getAttribute('data-rid') || '', w: +(el.getAttribute('data-w') || el.getAttribute('width') || 100),
        h: +(el.getAttribute('data-h') || el.getAttribute('height') || 100),
      },
      rPr: cleanRPr(rPr),
    })
    return
  }
  if (el.classList.contains('lm')) return // list marker, not content
  if (el.classList.contains('raw')) {
    out.push({ raw: decodeURIComponent(el.getAttribute('data-raw') || '') })
    return
  }
  if (tag === 'span' && el.classList.contains('tb')) { out.push({ tab: true, rPr: cleanRPr(rPr) }); return }
  let nextRPr = rPr
  let nextLink = link
  if (tag === 'a') {
    nextLink = { rId: el.getAttribute('data-rid') || '', url: el.getAttribute('href') || '' }
    nextRPr = { ...rPr, color: rPr.color || '#0563c1', underline: true }
  } else {
    nextRPr = mergeStyle(rPr, el)
  }
  for (const child of Array.from(el.childNodes)) syncRuns(child, nextRPr, nextLink, out)
}

function cleanRPr(r: RPr): RPr | undefined {
  const out: RPr = {}
  if (r.bold) out.bold = true
  if (r.italic) out.italic = true
  if (r.underline) out.underline = true
  if (r.strike) out.strike = true
  if (r.font) out.font = r.font
  if (r.size) out.size = r.size
  if (r.color) out.color = r.color
  if (r.highlight) out.highlight = r.highlight
  if (r.vertAlign) out.vertAlign = r.vertAlign
  return Object.keys(out).length ? out : undefined
}

export function syncPara(el: HTMLElement, prev?: Para): Para {
  const runs: Run[] = []
  for (const child of Array.from(el.childNodes)) syncRuns(child, {}, undefined, runs)
  const p: Para = prev ? { ...prev, runs } : emptyPara()
  p.id = el.getAttribute('data-id') || p.id
  // read alignment/indent back from the element style (execCommand justify sets it)
  const align = el.style.textAlign
  if (align) p.pPr = { ...p.pPr, align: align as NonNullable<Para['pPr']>['align'] }
  return p
}

export function syncBlock(el: HTMLElement, doc: Doc, prev?: Block): Block {
  const id = el.getAttribute('data-id') || undefined
  if (el.tagName.toLowerCase() === 'table') {
    const prevT = prev && prev.kind === 'tbl' ? prev : undefined
    const widths: number[] = Array.from(el.querySelectorAll('col')).map((c) => parseInt((c as HTMLElement).style.width) || 100)
    const rows: Cell[][] = Array.from(el.querySelectorAll(':scope > tbody > tr, :scope > tr')).map((tr) =>
      Array.from((tr as HTMLElement).querySelectorAll(':scope > td, :scope > th')).map((td) => {
        const cellEl = td as HTMLElement
        const paras: Para[] = Array.from(cellEl.querySelectorAll(':scope > .p')).map((p) => syncPara(p as HTMLElement))
        if (!paras.length) paras.push(emptyPara())
        const cell: Cell = { paras }
        const span = cellEl.getAttribute('colspan')
        if (span && +span > 1) cell.gridSpan = +span
        const vm = cellEl.getAttribute('data-vmerge')
        if (vm === 'restart' || vm === 'continue') cell.vMerge = vm
        const bg = rgbToHex(cellEl.style.backgroundColor)
        if (bg) cell.fill = bg
        const w = +(cellEl.getAttribute('data-w') || 0)
        if (w) cell.width = w
        return cell
      }))
    const t: Table = { kind: 'tbl', id: id || uid('t'), rows, widths: widths.length ? widths : (prevT?.widths || [100]) }
    return t
  }
  const prevP = prev && prev.kind === 'p' ? prev : undefined
  return syncPara(el, prevP)
}