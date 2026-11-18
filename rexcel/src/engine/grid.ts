/* Canvas grid renderer: headers, gridlines, styled cells, selection, merges. */
import { type Sheet, type Workbook, type Cell, addr, colName } from './model'
import { formatValue } from './numfmt'

export const HDR_W = 46
export const HDR_H = 22

export interface Layout { xs: number[]; ys: number[]; nw: number; nh: number }
export interface Sel { r: number; c: number; r2: number; c2: number }

export function computeLayout(sheet: Sheet, maxC = 200, maxR = 2000): Layout {
  const xs: number[] = [HDR_W]
  for (let c = 0; c < maxC; c++) xs.push(xs[c] + (sheet.colWidths.get(c) ?? sheet.defaultColWidth))
  const ys: number[] = [HDR_H]
  for (let r = 0; r < maxR; r++) ys.push(ys[r] + (sheet.rowHeights.get(r) ?? sheet.defaultRowHeight))
  return { xs, ys, nw: maxC, nh: maxR }
}

export function cellRect(layout: Layout, sheet: Sheet, r: number, c: number): { x: number; y: number; w: number; h: number } {
  const m = sheet.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2)
  if (m) {
    return { x: layout.xs[m.c1], y: layout.ys[m.r1], w: layout.xs[m.c2 + 1] - layout.xs[m.c1], h: layout.ys[m.r2 + 1] - layout.ys[m.r1] }
  }
  return { x: layout.xs[c], y: layout.ys[r], w: layout.xs[c + 1] - layout.xs[c], h: layout.ys[r + 1] - layout.ys[r] }
}

export function pointToCell(layout: Layout, x: number, y: number): { r: number; c: number } | null {
  if (x < HDR_W || y < HDR_H) return null
  let c = 0
  while (c < layout.nw && layout.xs[c + 1] < x) c++
  let r = 0
  while (r < layout.nh && layout.ys[r + 1] < y) r++
  return { r, c }
}

export interface DrawState {
  sheet: Sheet
  wb: Workbook
  layout: Layout
  sel: Sel
  scrollX: number
  scrollY: number
  zoom: number
  showGrid: boolean
  copyRange?: Sel | null
}

const css = (name: string, fallback: string): string => {
  const v = getComputedStyle(document.documentElement).getPropertyValue(name).trim()
  return v || fallback
}

export function drawGrid(canvas: HTMLCanvasElement, st: DrawState): void {
  const dpr = window.devicePixelRatio || 1
  const W = canvas.clientWidth, H = canvas.clientHeight
  if (canvas.width !== W * dpr || canvas.height !== H * dpr) { canvas.width = W * dpr; canvas.height = H * dpr }
  const ctx = canvas.getContext('2d')!
  ctx.setTransform(dpr * st.zoom, 0, 0, dpr * st.zoom, 0, 0)
  ctx.clearRect(0, 0, W / st.zoom, H / st.zoom)
  const { layout, sheet, wb, sel } = st
  const ox = HDR_W - st.scrollX
  const oy = HDR_H - st.scrollY
  const vw = W / st.zoom, vh = H / st.zoom

  const gridLine = css('--sheet-grid', '#c0c0c0')
  const hdrA = css('--hdr-a', '#fff'), hdrB = css('--hdr-b', '#E4ECF7'), hdrLine = css('--hdr-line', '#9EB6CE'), hdrInk = css('--hdr-ink', '#000')
  const selFill = css('--cell-sel', 'rgba(49,74,160,0.2)'), selBorder = css('--cell-sel-border', '#000')

  // visible cell range
  const c0 = Math.max(0, lowerBound(layout.xs, st.scrollX + HDR_W) - 1)
  const c1 = Math.min(layout.nw - 1, lowerBound(layout.xs, st.scrollX + vw) + 1)
  const r0 = Math.max(0, lowerBound(layout.ys, st.scrollY + HDR_H) - 1)
  const r1 = Math.min(layout.nh - 1, lowerBound(layout.ys, st.scrollY + vh) + 1)

  ctx.save()
  ctx.beginPath()
  ctx.rect(HDR_W, HDR_H, vw - HDR_W, vh - HDR_H)
  ctx.clip()
  ctx.translate(ox, oy)

  // cell fills + borders + text
  ctx.textBaseline = 'middle'
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      const rect = cellRect(layout, sheet, r, c)
      const cell = sheet.cells.get(addr(r, c))
      const style = cell?.s != null ? wb.styles[cell.s] : undefined
      if (style?.fill) { ctx.fillStyle = style.fill; ctx.fillRect(rect.x, rect.y, rect.w, rect.h) }
      if (st.showGrid) {
        ctx.strokeStyle = gridLine; ctx.lineWidth = 1 / st.zoom
        ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w, rect.h)
      }
      if (style?.border) {
        ctx.strokeStyle = '#000'; ctx.lineWidth = 1 / st.zoom
        const b = style.border
        ctx.beginPath()
        if (b.top) { ctx.moveTo(rect.x, rect.y + 0.5); ctx.lineTo(rect.x + rect.w, rect.y + 0.5) }
        if (b.bottom) { ctx.moveTo(rect.x, rect.y + rect.h - 0.5); ctx.lineTo(rect.x + rect.w, rect.y + rect.h - 0.5) }
        if (b.left) { ctx.moveTo(rect.x + 0.5, rect.y); ctx.lineTo(rect.x + 0.5, rect.y + rect.h) }
        if (b.right) { ctx.moveTo(rect.x + rect.w - 0.5, rect.y); ctx.lineTo(rect.x + rect.w - 0.5, rect.y + rect.h) }
        ctx.stroke()
      }
      if (cell && cell.v != null && cell.v !== '') {
        const f = style?.font
        const size = (f?.size || 11)
        ctx.font = `${f?.italic ? 'italic ' : ''}${f?.bold ? 'bold ' : ''}${size}px ${f?.name || 'Calibri'}, Carlito, sans-serif`
        ctx.fillStyle = f?.color || css('--ink', '#000')
        const text = cell.t === 'e' ? String(cell.v) : formatValue(cell.v, style?.numFmt)
        const align = style?.align?.h || (typeof cell.v === 'number' ? 'right' : 'left')
        const pad = 3
        let tx = rect.x + pad
        ctx.textAlign = 'left'
        if (align === 'right') { tx = rect.x + rect.w - pad; ctx.textAlign = 'right' }
        else if (align === 'center') { tx = rect.x + rect.w / 2; ctx.textAlign = 'center' }
        ctx.save()
        ctx.beginPath(); ctx.rect(rect.x, rect.y, rect.w, rect.h); ctx.clip()
        ctx.fillText(text, tx, rect.y + rect.h / 2)
        ctx.restore()
      }
    }
  }

  // selection overlay
  const sr = normSel(sel)
  const sx = layout.xs[sr.c1], sy = layout.ys[sr.r1]
  const sw = layout.xs[sr.c2 + 1] - sx, sh = layout.ys[sr.r2 + 1] - sy
  ctx.fillStyle = selFill
  ctx.fillRect(sx, sy, sw, sh)
  ctx.strokeStyle = selBorder; ctx.lineWidth = 2 / st.zoom
  ctx.strokeRect(sx + 1, sy + 1, sw - 2, sh - 2)
  // active cell white (no fill) outline
  const ar = cellRect(layout, sheet, sel.r, sel.c)
  ctx.strokeStyle = selBorder; ctx.lineWidth = 2 / st.zoom
  ctx.strokeRect(ar.x + 1, ar.y + 1, ar.w - 2, ar.h - 2)

  ctx.restore()

  // headers
  ctx.save()
  const grad = ctx.createLinearGradient(0, 0, 0, HDR_H)
  grad.addColorStop(0, hdrA); grad.addColorStop(1, hdrB)
  ctx.fillStyle = grad
  ctx.fillRect(0, 0, vw, HDR_H)
  ctx.fillRect(0, 0, HDR_W, vh)
  ctx.strokeStyle = hdrLine; ctx.lineWidth = 1 / st.zoom
  ctx.font = `${11}px ${css('--ui-font', 'Tahoma')}`
  ctx.fillStyle = hdrInk
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle'
  // column letters
  for (let c = c0; c <= c1; c++) {
    const x = layout.xs[c] + ox
    const w = layout.xs[c + 1] - layout.xs[c]
    const inSel = c >= sr.c1 && c <= sr.c2
    if (inSel) { ctx.fillStyle = css('--hdr-sel-b', '#F8C55E'); ctx.fillRect(x, 0, w, HDR_H); ctx.fillStyle = hdrInk }
    ctx.strokeRect(x + 0.5, 0.5, w, HDR_H)
    ctx.fillText(colName(c), x + w / 2, HDR_H / 2)
  }
  // row numbers
  ctx.textAlign = 'center'
  for (let r = r0; r <= r1; r++) {
    const y = layout.ys[r] + oy
    const h = layout.ys[r + 1] - layout.ys[r]
    const inSel = r >= sr.r1 && r <= sr.r2
    if (inSel) { ctx.fillStyle = css('--hdr-sel-b', '#F8C55E'); ctx.fillRect(0, y, HDR_W, h); ctx.fillStyle = hdrInk }
    ctx.strokeRect(0.5, y + 0.5, HDR_W, h)
    ctx.fillText(String(r + 1), HDR_W / 2, y + h / 2)
  }
  // corner
  ctx.fillStyle = grad
  ctx.fillRect(0, 0, HDR_W, HDR_H)
  ctx.strokeRect(0.5, 0.5, HDR_W, HDR_H)
  ctx.restore()
}

function lowerBound(arr: number[], v: number): number {
  let lo = 0, hi = arr.length - 1
  while (lo < hi) { const mid = (lo + hi) >> 1; if (arr[mid] < v) lo = mid + 1; else hi = mid }
  return lo
}
export function normSel(s: Sel): { r1: number; c1: number; r2: number; c2: number } {
  return { r1: Math.min(s.r, s.r2), c1: Math.min(s.c, s.c2), r2: Math.max(s.r, s.r2), c2: Math.max(s.c, s.c2) }
}
