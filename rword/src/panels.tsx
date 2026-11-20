/* The custom panels the original Quire opens from toolbar dropdowns: the colour
 * picker, the border grid, the table/column grid pickers, the line-style lists
 * and the shadow grid. They reuse the original's class names so the classic
 * looks.css / ui.css styles render them identically. */
import { useState, type ReactNode } from 'react'
import { getIcon } from './icons'
import { plainAmp, t } from './i18n'
import { useMenuClose } from './quireChrome'
import { shapesInGroup } from './geometryData'

/** The Office 2003 standard colour grid (public/common/core.js → L.color.STANDARD). */
export const STANDARD_COLORS = [
  '#000000', '#993300', '#333300', '#003300', '#003366', '#000080', '#333399', '#333333',
  '#800000', '#FF6600', '#808000', '#008000', '#008080', '#0000FF', '#666699', '#808080',
  '#FF0000', '#FF9900', '#99CC00', '#339966', '#33CCCC', '#3366FF', '#800080', '#969696',
  '#FF00FF', '#FFCC00', '#FFFF00', '#00FF00', '#00FFFF', '#00CCFF', '#993366', '#C0C0C0',
  '#FF99CC', '#FFCC99', '#FFFF99', '#CCFFCC', '#CCFFFF', '#99CCFF', '#CC99FF', '#FFFFFF',
]

/** The highlight colours (public/quire/js/render.js → R.HIGHLIGHT). */
export const HIGHLIGHT_COLORS: [string, string][] = [
  ['yellow', 'Yellow'], ['green', 'Bright Green'], ['cyan', 'Turquoise'], ['magenta', 'Pink'],
  ['blue', 'Blue'], ['red', 'Red'], ['darkBlue', 'Dark Blue'], ['darkCyan', 'Teal'],
  ['darkGreen', 'Green'], ['darkMagenta', 'Violet'], ['darkRed', 'Dark Red'], ['darkYellow', 'Dark Yellow'],
  ['darkGray', 'Gray-50%'], ['lightGray', 'Gray-25%'], ['black', 'Black'],
]

const HIGHLIGHT_HEX: Record<string, string> = {
  yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF',
  red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000',
  darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
  lightGray: '#C0C0C0', black: '#000000', white: '#FFFFFF',
}

/** The colour names Office shows for the standard palette (common/ui.js → CNAMES). */
const COLOR_NAMES: Record<string, string> = {
  '#000000': 'Black', '#993300': 'Brown', '#333300': 'Olive Green', '#003300': 'Dark Green',
  '#003366': 'Dark Teal', '#000080': 'Dark Blue', '#333399': 'Indigo', '#333333': 'Gray-80%',
  '#800000': 'Dark Red', '#FF6600': 'Orange', '#808000': 'Dark Yellow', '#008000': 'Green',
  '#008080': 'Teal', '#0000FF': 'Blue', '#666699': 'Blue-Gray', '#808080': 'Gray-50%',
  '#FF0000': 'Red', '#FF9900': 'Light Orange', '#99CC00': 'Lime', '#339966': 'Sea Green',
  '#33CCCC': 'Aqua', '#3366FF': 'Light Blue', '#800080': 'Violet', '#969696': 'Gray-40%',
  '#FF00FF': 'Pink', '#FFCC00': 'Gold', '#FFFF00': 'Yellow', '#00FF00': 'Bright Green',
  '#00FFFF': 'Turquoise', '#00CCFF': 'Sky Blue', '#993366': 'Plum', '#C0C0C0': 'Gray-25%',
  '#FF99CC': 'Rose', '#FFCC99': 'Tan', '#FFFF99': 'Light Yellow', '#CCFFCC': 'Light Green',
  '#CCFFFF': 'Light Turquoise', '#99CCFF': 'Pale Blue', '#CC99FF': 'Lavender', '#FFFFFF': 'White',
}

/** The name a colour shows as its tooltip, in the UI's language. */
export function colorName(hex: string): string {
  const name = COLOR_NAMES[hex.toUpperCase()]
  return name ? t(name) : hex.toUpperCase()
}

export interface ColorPick { none?: boolean; auto?: boolean; hex?: string }

/** The colour panel (ui.colorMenu): a wide automatic/none button, the standard grid. */
export function QColorPanel({ mode, onPick, onMore }: {
  mode: 'plain' | 'fill' | 'line' | 'font'
  onPick: (value: ColorPick) => void
  onMore?: () => void
}): ReactNode {
  const [recent, setRecent] = useState<string[]>([])
  const close = useMenuClose()
  /* the original closes the menu before it applies a colour */
  const pick = (hex: string) => {
    close()
    setRecent((r) => [hex, ...r.filter((c) => c !== hex)].slice(0, 8))
    onPick({ hex })
  }
  const wide = (value: ColorPick, more?: () => void) => () => { close(); more ? more() : onPick(value) }
  return <div className="color-panel">
    {mode === 'fill' && <button className="cp-wide" type="button" onClick={wide({ none: true })}>{t('No Fill')}</button>}
    {mode === 'line' && <button className="cp-wide" type="button" onClick={wide({ none: true })}>{t('No Line')}</button>}
    {(mode === 'font' || mode === 'plain') && <button className="cp-wide" type="button" onClick={wide({ auto: true })}>{t('Automatic')}</button>}
    <div className="cp-sep" />
    {recent.length > 0 && <>
      <div className="cp-row">
        {recent.map((c) => <button key={c} className="cp-sw" type="button" style={{ background: c }} aria-label={c} onClick={() => pick(c)} />)}
      </div>
      <div className="cp-sep" />
    </>}
    <div className="cp-grid">
      {STANDARD_COLORS.map((c) => <button key={c} className="cp-sw" type="button" style={{ background: c }} aria-label={c} data-tip={colorName(c)} onClick={() => pick(c)} />)}
    </div>
    <div className="cp-sep" />
    <button className="cp-wide" type="button" onClick={wide({}, () => onMore?.())}>
      {mode === 'fill' ? t('More Fill Colors...') : mode === 'line' ? t('More Line Colors...') : t('More Colors...')}
    </button>
  </div>
}

/** The highlight panel: the fifteen highlight colours plus a None button. */
export function QHighlightPanel({ onPick }: { onPick: (name: string) => void }): ReactNode {
  const close = useMenuClose()
  const pick = (name: string) => { close(); onPick(name) }
  return <div className="color-panel">
    <div className="cp-grid hl">
      {HIGHLIGHT_COLORS.map(([key, name]) => <button key={key} className="cp-sw" type="button"
        style={{ background: HIGHLIGHT_HEX[key] }} data-tip={plainAmp(t(name))} aria-label={plainAmp(t(name))}
        onClick={() => pick(key)} />)}
    </div>
    <button className="cp-wide" type="button" onClick={() => pick('none')}>{t('None')}</button>
  </div>
}

/** The border grid: the thirteen border presets the original offers. */
export const BORDER_ITEMS: [string, string, string][] = [
  ['outside', 'bordersOut', 'Outside Border'], ['all', 'bordersAll', 'All Borders'],
  ['top', 'bordersTop', 'Top Border'], ['bottom', 'bordersBottom', 'Bottom Border'],
  ['left', 'bordersLeft', 'Left Border'], ['right', 'bordersRight', 'Right Border'],
  ['insideH', 'bordersInH', 'Inside Horizontal Border'], ['insideV', 'bordersInV', 'Inside Vertical Border'],
  ['inside', 'bordersIn', 'Inside Border'], ['diagDown', 'bordersDiagDown', 'Diagonal Down Border'],
  ['diagUp', 'bordersDiagUp', 'Diagonal Up Border'], ['hline', 'hline', 'Horizontal Line'],
  ['none', 'bordersNone', 'No Border'],
]

export function QBorderGrid({ onPick }: { onPick: (key: string) => void }): ReactNode {
  const close = useMenuClose()
  return <div className="border-grid">
    {BORDER_ITEMS.map(([key, icon, name]) => <button key={key} type="button" data-tip={plainAmp(t(name))}
      aria-label={plainAmp(t(name))} onClick={() => { close(); onPick(key) }}
      dangerouslySetInnerHTML={{ __html: getIcon(icon) }} />)}
  </div>
}

/** The table grid picker: a growing grid whose label shows the size. */
export function QTableGridPicker({ onPick }: { onPick: (rows: number, cols: number) => void }): ReactNode {
  const [rows, setRows] = useState(4)
  const [cols, setCols] = useState(5)
  const [cur, setCur] = useState<[number, number]>([0, 0])
  const close = useMenuClose()
  const cells: ReactNode[] = []
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const on = r < cur[0] && c < cur[1]
      cells.push(<span key={`${r},${c}`} className={on ? 'on' : undefined}
        onMouseEnter={() => {
          setCur([r + 1, c + 1])
          if (r + 2 > rows && rows < 12) setRows(r + 2)
          if (c + 2 > cols && cols < 10) setCols(c + 2)
        }} />)
    }
  }
  return <div className="tbl-grid-pick">
    {/* the original inserts the table when a cell is clicked */}
    <div className="tgp-cells" style={{ gridTemplateColumns: `repeat(${cols}, 18px)` }}
      onClick={() => { if (cur[0]) { close(); onPick(cur[0], cur[1]) } }}>{cells}</div>
    <div className="tgp-label">{cur[0] ? t('{0} x {1} Table', String(cur[0]), String(cur[1])) : t('Cancel')}</div>
  </div>
}

/** The column picker: a row of columns; the label counts them. */
export function QColumnPicker({ onPick }: { onPick: (columns: number) => void }): ReactNode {
  const [cur, setCur] = useState(0)
  const close = useMenuClose()
  return <div className="tbl-grid-pick">
    <div className="tgp-cells" style={{ gridTemplateColumns: 'repeat(4, 22px)' }}>
      {[0, 1, 2, 3].map((i) => <span key={i} style={{ height: 30 }}
        onMouseEnter={() => setCur(i + 1)}
        onClick={() => { close(); onPick(i + 1) }} />)}
    </div>
    <div className="tgp-label">{cur ? t('{0} Columns', String(cur)) : t('Cancel')}</div>
  </div>
}

/** A line-style list (the original's listPanel): a row of previews with labels. */
export function QStyleList({ rows }: { rows: { html: string; label?: string; run: () => void }[] }): ReactNode {
  const close = useMenuClose()
  return <div className="style-list">
    {rows.map((row, i) => <button key={i} type="button" onClick={() => { close(); row.run() }}
      dangerouslySetInnerHTML={{ __html: row.html }} />)}
  </div>
}

/** The shadow grid: eight shadow presets plus "no shadow". */
export function QShadowGrid({ onPick }: { onPick: (offset: [number, number] | null) => void }): ReactNode {
  const shadows: ([number, number] | null)[] = [null, [3, 3], [-3, 3], [3, -3], [-3, -3], [0, 4], [4, 0], [6, 6]]
  const close = useMenuClose()
  return <div className="shape-grid" style={{ gridTemplateColumns: 'repeat(4, 30px)' }}>
    {shadows.map((s, i) => <button key={i} type="button" style={{ width: 30, height: 30 }}
      aria-label={s ? t('Shadow Style') : t('No Shadow')} onClick={() => { close(); onPick(s) }}
      dangerouslySetInnerHTML={{ __html: s
        ? `<svg width="24" height="24"><rect x="${6 + s[0] / 1.5}" y="${6 + s[1] / 1.5}" width="12" height="12" fill="#888"/><rect x="6" y="6" width="12" height="12" fill="#fff" stroke="#000"/></svg>`
        : '<svg width="24" height="24"><rect x="6" y="6" width="12" height="12" fill="#fff" stroke="#000"/><path d="M3 21L21 3" stroke="#d33"/></svg>' }} />)}
  </div>
}

/** A line preview used by the line-weight and dash-style lists. */
export function lineSVG(width: number, dash?: string): string {
  return `<svg width="90" height="14" viewBox="0 0 90 14"><path d="M2 7h86" stroke="#000" stroke-width="${width}"${dash ? ` stroke-dasharray="${dash}"` : ''} fill="none"/></svg>`
}

/* ------------------------------------------------------------ the shape gallery */

/** The AutoShapes gallery: one group of the original's shapes, drawn as they are.
 * Each button carries the original's own preview markup, so the gallery's
 * pictures are the original's, pixel for pixel. */
export function QShapeGallery({ group, onPick }: { group: string; onPick: (geom: string) => void }): ReactNode {
  const shapes = shapesInGroup(group)
  const close = useMenuClose()
  return <div className="shape-grid">
    {shapes.map((s) => <button key={s.geom} type="button" data-tip={plainAmp(t(s.label))}
      aria-label={plainAmp(t(s.label))} onClick={() => { close(); onPick(s.geom) }}
      dangerouslySetInnerHTML={{ __html: s.preview }} />)}
  </div>
}
