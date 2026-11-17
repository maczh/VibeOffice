/* Static HTML rendering of a slide (thumbnails, slideshow, print). */
import { type Presentation, type Slide, type Shape, px } from './model'
import { presetSvg } from './shapes'

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

export function shapeHtml(s: Shape): string {
  const x = px(s.x), y = px(s.y), w = px(s.w), h = px(s.h)
  const rot = s.rot ? `transform:rotate(${s.rot}deg);` : ''
  const base = `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${h}px;${rot}`
  if (s.raw != null) return `<div style="${base}border:1px dashed #999;background:#eee"></div>`
  if (s.kind === 'pic' && s.image) return `<div style="${base}"><img src="${s.image.url}" style="width:100%;height:100%;object-fit:fill" draggable="false"></div>`
  const svg = presetSvg(s.preset || 'rect', w, h, s.fill ? s.fill.color : null, s.line ? s.line.color : null, s.line?.width || 1)
  const text = (s.text || []).map((p) => {
    const align = p.align || 'left'
    const runs = p.runs.map((r) => {
      const st: string[] = []
      if (r.rPr?.bold) st.push('font-weight:bold')
      if (r.rPr?.italic) st.push('font-style:italic')
      if (r.rPr?.underline) st.push('text-decoration:underline')
      if (r.rPr?.size) st.push(`font-size:${r.rPr.size}px`)
      if (r.rPr?.color) st.push(`color:${r.rPr.color}`)
      if (r.rPr?.font) st.push(`font-family:${r.rPr.font},sans-serif`)
      return `<span style="${st.join(';')}">${esc(r.text)}</span>`
    }).join('')
    return `<div style="text-align:${align};min-height:1em">${p.bullet ? '• ' : ''}${runs}</div>`
  }).join('')
  const hasText = (s.text || []).some((p) => p.runs.some((r) => r.text))
  return `<div style="${base}">` +
    `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" style="position:absolute;inset:0;overflow:visible">${svg}</svg>` +
    (hasText || s.kind === 'sp' ? `<div style="position:absolute;inset:4px;overflow:hidden;color:#000;font-size:18px">${text}</div>` : '') +
    `</div>`
}

export function slideHtml(slide: Slide, pres: Presentation, widthPx: number): string {
  const scale = widthPx / px(pres.slideW)
  const inner = slide.shapes.map(shapeHtml).join('')
  return `<div style="position:relative;width:${px(pres.slideW)}px;height:${px(pres.slideH)}px;background:${slide.background || '#fff'};transform:scale(${scale});transform-origin:top left">${inner}</div>`
}