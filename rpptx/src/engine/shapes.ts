/* Preset geometry → SVG path/shape markup for a w×h box (user units = px). */

export function presetSvg(preset: string, w: number, h: number, fill: string | null, line: string | null, lineW: number): string {
  const f = fill || 'none'
  const stroke = line || 'none'
  const sw = line ? lineW : 0
  const common = `fill="${f}" stroke="${stroke}" stroke-width="${sw}"`
  const P = (d: string) => `<path d="${d}" ${common}/>`
  const r = (x: number, y: number, ww: number, hh: number) => `<rect x="${x}" y="${y}" width="${ww}" height="${hh}" ${common}/>`
  const e = (cx: number, cy: number, rx: number, ry: number) => `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" ${common}/>`
  const poly = (pts: [number, number][]) => `<polygon points="${pts.map((p) => p.join(',')).join(' ')}" ${common}/>`
  const m = sw / 2
  switch (preset) {
    case 'ellipse': case 'oval': return e(w / 2, h / 2, w / 2 - m, h / 2 - m)
    case 'triangle': return poly([[w / 2, m], [w - m, h - m], [m, h - m]])
    case 'rtTriangle': return poly([[m, m], [m, h - m], [w - m, h - m]])
    case 'diamond': return poly([[w / 2, m], [w - m, h / 2], [w / 2, h - m], [m, h / 2]])
    case 'parallelogram': return poly([[w * 0.25, m], [w - m, m], [w * 0.75, h - m], [m, h - m]])
    case 'trapezoid': return poly([[w * 0.25, m], [w * 0.75, m], [w - m, h - m], [m, h - m]])
    case 'pentagon': return poly([[w / 2, m], [w - m, h * 0.4], [w * 0.8, h - m], [w * 0.2, h - m], [m, h * 0.4]])
    case 'hexagon': return poly([[w * 0.25, m], [w * 0.75, m], [w - m, h / 2], [w * 0.75, h - m], [w * 0.25, h - m], [m, h / 2]])
    case 'octagon': return poly([[w * 0.3, m], [w * 0.7, m], [w - m, h * 0.3], [w - m, h * 0.7], [w * 0.7, h - m], [w * 0.3, h - m], [m, h * 0.7], [m, h * 0.3]])
    case 'star5': {
      const pts: [number, number][] = []
      for (let i = 0; i < 10; i++) {
        const rad = i % 2 === 0 ? Math.min(w, h) / 2 - m : Math.min(w, h) / 4.5
        const a = (Math.PI / 5) * i - Math.PI / 2
        pts.push([w / 2 + rad * Math.cos(a) * (w / Math.min(w, h)), h / 2 + rad * Math.sin(a) * (h / Math.min(w, h))])
      }
      return poly(pts)
    }
    case 'heart': return P(`M ${w / 2} ${h * 0.9} C ${w * 0.1} ${h * 0.55} ${w * 0.05} ${h * 0.2} ${w * 0.3} ${h * 0.15} C ${w * 0.45} ${h * 0.12} ${w / 2} ${h * 0.3} ${w / 2} ${h * 0.35} C ${w / 2} ${h * 0.3} ${w * 0.55} ${h * 0.12} ${w * 0.7} ${h * 0.15} C ${w * 0.95} ${h * 0.2} ${w * 0.9} ${h * 0.55} ${w / 2} ${h * 0.9} Z`)
    case 'rightArrow': return poly([[m, h * 0.3], [w * 0.6, h * 0.3], [w * 0.6, m], [w - m, h / 2], [w * 0.6, h - m], [w * 0.6, h * 0.7], [m, h * 0.7]])
    case 'leftArrow': return poly([[w - m, h * 0.3], [w * 0.4, h * 0.3], [w * 0.4, m], [m, h / 2], [w * 0.4, h - m], [w * 0.4, h * 0.7], [w - m, h * 0.7]])
    case 'upArrow': return poly([[w * 0.3, h - m], [w * 0.3, h * 0.4], [m, h * 0.4], [w / 2, m], [w - m, h * 0.4], [w * 0.7, h * 0.4], [w * 0.7, h - m]])
    case 'downArrow': return poly([[w * 0.3, m], [w * 0.3, h * 0.6], [m, h * 0.6], [w / 2, h - m], [w - m, h * 0.6], [w * 0.7, h * 0.6], [w * 0.7, m]])
    case 'chevron': return poly([[m, m], [w * 0.7, m], [w - m, h / 2], [w * 0.7, h - m], [m, h - m], [w * 0.3, h / 2]])
    case 'homePlate': return poly([[m, m], [w * 0.7, m], [w - m, h / 2], [w * 0.7, h - m], [m, h - m]])
    case 'roundRect': return `<rect x="${m}" y="${m}" width="${w - sw}" height="${h - sw}" rx="${Math.min(w, h) * 0.12}" ${common}/>`
    case 'line': case 'straightConnector1': return `<line x1="${m}" y1="${h - m}" x2="${w - m}" y2="${m}" stroke="${line || '#000'}" stroke-width="${Math.max(1, sw)}"/>`
    case 'bentConnector3': return P(`M ${m} ${h - m} L ${w / 2} ${h - m} L ${w / 2} ${m} L ${w - m} ${m}`)
    case 'cloud': return P(`M ${w * 0.25} ${h * 0.7} a ${w * 0.15} ${h * 0.18} 0 0 1 ${w * 0.05} ${-h * 0.35} a ${w * 0.18} ${h * 0.2} 0 0 1 ${w * 0.3} ${-h * 0.15} a ${w * 0.18} ${h * 0.2} 0 0 1 ${w * 0.3} ${h * 0.15} a ${w * 0.15} ${h * 0.18} 0 0 1 ${w * 0.05} ${h * 0.35} Z`)
    case 'rect': default: return r(m, m, w - sw, h - sw)
  }
}

export const PRESETS = ['rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid', 'pentagon', 'hexagon', 'octagon', 'star5', 'heart', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'chevron', 'homePlate', 'line', 'cloud']
