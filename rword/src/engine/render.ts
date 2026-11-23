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

/* The same CSS font stacks Quire uses in public/common/core.js, so the page paints with
   the browser's stand-ins for the font names kept in the document. */
const FONT_FALLBACK: Record<string, string> = {
  arial: '"Arimo","Liberation Sans",Helvetica,sans-serif',
  'arial black': '"Arial Black","Arimo",sans-serif',
  'arial narrow': '"Arial Narrow","Arimo",sans-serif',
  helvetica: '"Helvetica Neue","TeX Gyre Heros","Nimbus Sans",Arial,"Arimo",sans-serif',
  'times new roman': '"Tinos","Liberation Serif",Times,serif',
  times: '"Times New Roman","Tinos","TeX Gyre Termes",serif',
  'courier new': '"Cousine","Liberation Mono",Courier,monospace',
  courier: '"Courier New","Cousine",monospace',
  calibri: '"Carlito",Candara,"Segoe UI",sans-serif',
  'calibri light': '"Carlito",sans-serif',
  aptos: 'Calibri,"Carlito",sans-serif',
  georgia: '"Gelasio","DejaVu Serif","Times New Roman","Tinos",serif',
  cambria: '"Caladea",Georgia,serif',
  tahoma: '"DejaVu Sans Condensed",Verdana,Arial,"Arimo",sans-serif',
  verdana: '"DejaVu Sans",Tahoma,Arial,"Arimo",sans-serif',
  symbol: '"Segoe UI Symbol",serif',
  wingdings: '"Segoe UI Symbol",sans-serif',
}

export function fontStack(name: string): string {
  const family = String(name).replace(/"/g, '').trim()
  const fallback = FONT_FALLBACK[family.toLowerCase()] || 'Arial,"Arimo",sans-serif'
  /* A style attribute is double-quoted; use single quotes in the CSS value. */
  return `'${family}',${fallback.replace(/"/g, "'")}`
}

function rPrStyle(r: RPr | undefined, doc: Doc): string {
  const s: string[] = []
  const font = r?.font || doc.defaultFont
  const size = r?.size || doc.defaultSize
  s.push(`font-family:${fontStack(font)}`)
  s.push(`font-size:${size}pt`)
  if (r?.bold) s.push('font-weight:bold')
  if (r?.italic) s.push('font-style:italic')
  const td: string[] = []
  if (r?.underline) td.push('underline')
  if (r?.strike) td.push('line-through')
  if (td.length) s.push('text-decoration-line:' + td.join(' '))
  if (r?.underlineStyle === 'double') s.push('text-decoration-style:double')
  if (r?.underlineStyle === 'words') s.push('text-decoration-style:dotted')
  if (r?.caps === 'all') s.push('text-transform:uppercase')
  if (r?.caps === 'small') s.push('font-variant:small-caps')
  // Some Word files mark CJK runs as vanished while still expecting them to
  // be visible in the normal document view. Rendering them transparent makes
  // the text appear only when selected, so keep the run readable here.
  if (r?.color && !isInvisibleLightColor(r.color)) s.push('color:' + r.color)
  if (r?.highlight) s.push('background:' + r.highlight)
  if (r?.vertAlign === 'superscript') s.push('vertical-align:super;font-size:smaller')
  if (r?.vertAlign === 'subscript') s.push('vertical-align:sub;font-size:smaller')
  return s.join(';')
}

function isInvisibleLightColor(value: string): boolean {
  const color = value.trim().toLowerCase().replace(/\s+/g, '')
  return color === 'transparent' || color === '#fff' || color === '#ffffff' ||
    color === 'rgb(255,255,255)' || color === 'rgba(255,255,255,1)' || color === 'rgba(255,255,255,0)'
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
  if (p.pPr?.pageBreakBefore) s.push('break-before:page;page-break-before:always')
  if (p.pPr?.keepNext) s.push('page-break-after:avoid')
  if (p.pPr?.keepLines) s.push('page-break-inside:avoid')
  const borders = p.pPr?.borders
  if (borders) {
    for (const edge of ['top', 'bottom', 'left', 'right'] as const) {
      const b = borders[edge]
      if (!b || b.val === 'nil' || b.val === 'none') continue
      const width = Math.max(0.5, (b.sz ?? 0.5) * 1.33)
      const style = b.val === 'dashed' || b.val === 'dotted' || b.val === 'dotDash'
        ? (b.val === 'dotted' ? 'dotted' : 'dashed')
        : b.val === 'double' ? 'double' : 'solid'
      s.push(`border-${edge}:${width}px ${style} ${b.color && b.color !== 'auto' ? '#' + b.color : '#000'}`)
    }
  }
  if (p.pPr?.shading) s.push('background:' + p.pPr.shading)
  const rPr = st?.rPr
  if (rPr) s.push(rPrStyle(rPr, doc))
  return s.join(';')
}

function runHtml(r: Run, doc: Doc): string {
  if (r.raw != null) return `<span class="raw" contenteditable="false" data-raw="${encodeURIComponent(r.raw)}">[object]</span>`
  const style = rPrStyle(r.rPr, doc)
  if (r.image) {
    const img = r.image
    const css = [style]
    if (img.rotation) css.push(`transform:rotate(${img.rotation}deg)`)
    if (img.contrast != null) css.push(`filter:contrast(${img.contrast}%)`)
    if (img.brightness != null) css.push(`filter:brightness(${img.brightness}%)`)
    if (img.gray) css.push('filter:grayscale(1)')
    if (img.bw) css.push('filter:grayscale(1) contrast(1.6)')
    if (img.washout) css.push('filter:brightness(1.4) contrast(0.6)')
    if (img.offset && (img.offset.x || img.offset.y)) css.push(`position:relative;left:${img.offset.x || 0}px;top:${img.offset.y || 0}px`)
    if (img.wrap === 'square' || img.wrap === 'tight') css.push('float:left;margin:0 8px 4px 0')
    if (img.wrap === 'behind') css.push('position:absolute;z-index:-1')
    if (img.wrap === 'none' || img.wrap === 'topBottom' || img.wrap === 'through') css.push('float:left;margin:0 8px 4px 0')
    const crop = img.crop
    return `<img class="img" src="${img.url}" width="${img.w}" height="${img.h}" data-part="${img.part}" data-rid="${img.rId}"` +
      ` data-w="${img.w}" data-h="${img.h}"${crop ? ` data-crop="${crop.l},${crop.t},${crop.r},${crop.b}"` : ''}` +
      ` data-wrap="${img.wrap || 'inline'}" style="${css.join(';')}">`
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
    const style = b.pPr?.style ? doc.styles.find((s) => s.id === b.pPr!.style) : undefined
    const inner = b.runs.map((r) => runHtml({ ...r, rPr: { ...style?.rPr, ...r.rPr } }, doc)).join('')
    const marker = listMarker(b, doc)
    return `<div class="p" data-id="${b.id}" style="${pPrStyle(b, doc)}">${marker}${inner || '<br>'}</div>`
  }
  const t = b as Table
  const cols = t.widths.map((w) => `<col style="width:${w}px">`).join('')
  const tableStyle = [
    'border-collapse:collapse',
    `table-layout:${t.autoFit === 'content' ? 'auto' : 'fixed'}`,
    'break-inside:avoid', 'page-break-inside:avoid',
    t.align === 'center' ? 'margin-left:auto;margin-right:auto' : t.align === 'right' ? 'margin-left:auto' : '',
  ].filter(Boolean).join(';')
  const rows = t.rows.map((row, ri) => {
    const cells = row.map((c) => {
      const span = c.gridSpan && c.gridSpan > 1 ? ` colspan="${c.gridSpan}"` : ''
      const vm = c.vMerge === 'continue' ? ' data-vmerge="continue"' : c.vMerge === 'restart' ? ' data-vmerge="restart"' : ''
      const border = cellBorderStyle(c)
      const css = [
        c.fill ? `background:${c.fill}` : '', border, 'padding:2px 4px',
        `vertical-align:${c.vAlign === 'center' ? 'middle' : c.vAlign === 'bottom' ? 'bottom' : 'top'}`,
        c.textDirection === 'vertical' ? 'writing-mode:vertical-lr' : '',
        ri < (t.headerRows || 0) ? 'font-weight:bold' : '',
      ].filter(Boolean).join(';')
      const paras = c.paras.map((p) => renderBlock(p, doc)).join('')
      return `<td${span}${vm} data-w="${c.width || 0}" style="${css}">${paras}</td>`
    }).join('')
    return `<tr>${cells}</tr>`
  }).join('')
  return `<table class="tbl" data-id="${t.id}" style="${tableStyle}"><colgroup>${cols}</colgroup>${rows}</table>`
}

function listMarker(p: Para, doc: Doc): string {
  if (p.pPr?.numId == null) return ''
  // numbering is resolved by the component (needs sequence context); a bullet is static
  return `<span class="lm" contenteditable="false" data-num="${p.pPr.numId}" data-ilvl="${p.pPr.ilvl || 0}"></span>`
}

/** The cell's border CSS: a per-edge set when the cell has one, a plain frame otherwise. */
function cellBorderStyle(c: Cell): string {
  if (!c.borders) return 'border:1px solid #888'
  const parts: string[] = []
  for (const edge of ['top', 'bottom', 'left', 'right'] as const) {
    const b = c.borders[edge]
    if (!b || b.val === 'nil' || b.val === 'none') { parts.push(`border-${edge}:0`); continue }
    const width = Math.max(0.5, (b.sz ?? 0.5) * 1.33)
    const style = b.val === 'dashed' || b.val === 'dotted' || b.val === 'dotDash'
      ? (b.val === 'dotted' ? 'dotted' : 'dashed')
      : b.val === 'double' ? 'double' : 'solid'
    parts.push(`border-${edge}:${width}px ${style} ${b.color && b.color !== 'auto' ? '#' + b.color : '#000'}`)
  }
  return parts.join(';')
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
  if (col && !isInvisibleLightColor(col)) out.color = col
  const bg = rgbToHex(st.backgroundColor)
  if (bg && bg !== '#ffffff') out.highlight = HL_NAMES[bg.toLowerCase()] || bg
  if (st.verticalAlign === 'super') out.vertAlign = 'superscript'
  if (st.verticalAlign === 'sub') out.vertAlign = 'subscript'
  if (st.textTransform === 'uppercase') out.caps = 'all'
  if (st.fontVariantCaps === 'small-caps') out.caps = 'small'
  if (st.textDecorationStyle === 'double') out.underlineStyle = 'double'
  if (st.textDecorationStyle === 'dotted') out.underlineStyle = 'words'
  if (st.color === 'rgba(0, 0, 0, 0)' || st.color === 'transparent') out.hidden = true
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
    const crop = (el.getAttribute('data-crop') || '').split(',').map(Number)
    out.push({
      image: {
        url: el.getAttribute('src') || '', part: el.getAttribute('data-part') || '',
        rId: el.getAttribute('data-rid') || '', w: +(el.getAttribute('data-w') || el.getAttribute('width') || 100),
        h: +(el.getAttribute('data-h') || el.getAttribute('height') || 100),
        ...(crop.length === 4 && crop.some((n) => n > 0) ? { crop: { l: crop[0], r: crop[2], t: crop[1], b: crop[3] } } : {}),
        wrap: (el.getAttribute('data-wrap') || 'inline') as NonNullable<NonNullable<Run['image']>['wrap']>,
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
  if (r.caps) out.caps = r.caps
  if (r.hidden) out.hidden = r.hidden
  if (r.underlineStyle) out.underlineStyle = r.underlineStyle
  return Object.keys(out).length ? out : undefined
}

export function syncPara(el: HTMLElement, prev?: Para): Para {
  const runs: Run[] = []
  for (const child of Array.from(el.childNodes)) syncRuns(child, {}, undefined, runs)
  /* The parsed runs are the content of record: keep them even when no previous
   * model paragraph is supplied (table cells are synced without one), or every
   * surface sync would wipe all text out of the table. */
  const p: Para = prev ? { ...prev, runs } : { ...emptyPara(), runs }
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
    /* Match each cell paragraph with its previous model entry (by data-id) so
     * paragraph formatting (indent, spacing, numbering…) survives the round trip. */
    const prevParas = new Map<string, Para>()
    for (const row of prevT?.rows || []) for (const cell of row) for (const para of cell.paras) prevParas.set(para.id, para)
    const widths: number[] = Array.from(el.querySelectorAll('col')).map((c) => parseInt((c as HTMLElement).style.width) || 100)
    const rows: Cell[][] = Array.from(el.querySelectorAll(':scope > tbody > tr, :scope > tr')).map((tr) =>
      Array.from((tr as HTMLElement).querySelectorAll(':scope > td, :scope > th')).map((td) => {
        const cellEl = td as HTMLElement
        const paras: Para[] = Array.from(cellEl.querySelectorAll(':scope > .p')).map((p) => {
          const element = p as HTMLElement
          return syncPara(element, prevParas.get(element.getAttribute('data-id') || ''))
        })
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
        const va = cellEl.style.verticalAlign
        if (va === 'middle' || va === 'bottom') cell.vAlign = va === 'middle' ? 'center' : 'bottom'
        if (cellEl.style.writingMode.includes('vertical')) cell.textDirection = 'vertical'
        {
          const edges: NonNullable<Cell['borders']> = {}
          for (const edge of ['top', 'bottom', 'left', 'right'] as const) {
            const side = `border${edge[0].toUpperCase()}${edge.slice(1)}`
            const style = cellEl.style as unknown as Record<string, string>
            const width = parseFloat(style[side + 'Width'] || '0')
            if (!width) continue
            const edgeStyle = style[side + 'Style'] || 'solid'
            const colour = rgbToHex(style[side + 'Color']) || '#000000'
            edges[edge] = { val: edgeStyle === 'none' ? 'nil' : edgeStyle, sz: Math.round(width / 1.33 * 100) / 100, color: colour.replace('#', '') }
          }
          if (Object.keys(edges).length) cell.borders = edges
        }
        return cell
      }))
    const t: Table = {
      kind: 'tbl', id: id || uid('t'), rows,
      widths: widths.length ? widths : (prevT?.widths || [100]),
      ...(prevT?.align ? { align: prevT.align } : {}),
      ...(prevT?.autoFit ? { autoFit: prevT.autoFit } : {}),
      ...(prevT?.headerRows ? { headerRows: prevT.headerRows } : {}),
    }
    return t
  }
  const prevP = prev && prev.kind === 'p' ? prev : undefined
  return syncPara(el, prevP)
}
