/* The commands the toolbar dropdowns open: border presets, font colour, highlight,
 * line spacing, table and column grids, and the drawing-object panels. They act
 * on the same model the rest of the editor uses, so the result survives a rebuild
 * and a save. */
import { type Border, type BorderSet, type Cell, type ImageRef, type Para, isPara } from '../engine/model'
import { SHAPE_BY_GEOM } from '../geometryData'
import { type CommandEnv, type CommandFn } from './documentCommands'

/* ------------------------------------------------------------------ borders */

/** The pen the border buttons draw with: style, width in points and colour. */
export interface BorderPen { val: string; sz: number; color: string }
export const borderPen: BorderPen = { val: 'single', sz: 0.5, color: 'auto' }

const EDGES = ['top', 'bottom', 'left', 'right'] as const
type Edge = typeof EDGES[number]
type EdgeKey = Edge | 'between' | 'insideH' | 'insideV' | 'diagDown' | 'diagUp'

function applyBorderSet(current: BorderSet | undefined, key: string, pen: BorderPen): BorderSet | undefined {
  const next: BorderSet = { ...(current || {}) }
  const has = (edge: EdgeKey) => {
    const b = next[edge]
    return !!(b && b.val && b.val !== 'nil')
  }
  const set = (edge: Edge, on: boolean) => {
    if (on) next[edge] = { val: pen.val, sz: pen.sz, color: pen.color, space: edge === 'left' || edge === 'right' ? 4 : 1 }
    else delete next[edge]
  }
  switch (key) {
    case 'none':
      return undefined
    case 'outside':
    case 'all': {
      const on = !(has('top') && has('bottom') && has('left') && has('right'))
      EDGES.forEach((edge) => set(edge, on))
      if (key === 'all') next.between = on ? { val: pen.val, sz: pen.sz, color: pen.color } : undefined
      break
    }
    case 'inside':
    case 'insideH':
      next.between = has('between') ? undefined : { val: pen.val, sz: pen.sz, color: pen.color }
      break
    case 'insideV':
      next.insideV = has('insideV') ? undefined : { val: pen.val, sz: pen.sz, color: pen.color }
      break
    case 'diagDown':
    case 'diagUp':
      next[key] = has(key) ? undefined : { val: pen.val, sz: pen.sz, color: pen.color }
      break
    default:
      if ((EDGES as readonly string[]).includes(key)) {
        const edge = key as Edge
        set(edge, !has(edge))
      }
  }
  return Object.keys(next).length ? next : undefined
}

/** Apply a border preset to the paragraph at the caret, or to the table cell. */
export function applyBorder(env: CommandEnv, key: string): string | undefined {
  const pen: BorderPen = { val: borderPen.val, sz: borderPen.sz, color: borderPen.color }
  if (env.loc?.table) {
    const cell = env.loc.table.cell
    cell.borders = applyBorderSet(cell.borders, key, pen)
    return undefined
  }
  const para = env.loc?.para
  if (!para) return 'Place the cursor in a paragraph first.'
  if (key === 'hline') {
    para.pPr = {
      ...para.pPr,
      borders: { ...(para.pPr?.borders || {}), bottom: { val: 'single', sz: 1.5, color: 'auto', space: 1 } },
    }
    return undefined
  }
  para.pPr = { ...para.pPr, borders: applyBorderSet(para.pPr?.borders, key, pen) }
  return undefined
}

/** The shading colour of the paragraph or cell at the caret. */
export function applyShading(env: CommandEnv, hex: string | null): string | undefined {
  if (env.loc?.table) {
    env.loc.table.cell.fill = hex ?? undefined
    return undefined
  }
  const para = env.loc?.para
  if (!para) return 'Place the cursor in a paragraph first.'
  para.pPr = { ...para.pPr, shading: hex ?? undefined }
  return undefined
}

/* ----------------------------------------------------------- run formatting */

/** Apply a character format to every run of the paragraph at the caret. */
export function formatRuns(env: CommandEnv, patch: Record<string, unknown>): string | undefined {
  const para = env.loc?.para
  if (!para) return 'Place the cursor in a paragraph first.'
  const runs = para.runs.filter((r) => r.text)
  if (!runs.length) {
    para.runs.push({ text: '', rPr: patch as never })
    return undefined
  }
  for (const run of runs) run.rPr = { ...run.rPr, ...patch } as never
  return undefined
}

/** The font colour of the caret: `auto` restores the inherited colour. */
export function applyFontColor(env: CommandEnv, hex: string | null): string | undefined {
  return formatRuns(env, hex ? { color: hex } : { color: undefined })
}

/** The highlight colour of the caret; `none` removes it. */
export function applyHighlight(env: CommandEnv, name: string): string | undefined {
  return formatRuns(env, name === 'none' ? { highlight: undefined } : { highlight: name })
}

/* -------------------------------------------------------------- line spacing */

/** The line spacing of the caret, as the line-spacing menu shows it. */
export function currentLineSpacing(env: CommandEnv): number {
  return env.loc?.para?.pPr?.line || 1
}

/** Set the line spacing of the caret. */
export function setLineSpacing(env: CommandEnv, value: number): string | undefined {
  const para = env.loc?.para
  if (!para) return 'Place the cursor in a paragraph first.'
  para.pPr = { ...para.pPr, line: value }
  return undefined
}

/* ------------------------------------------------------------------- tables */

/** Insert a table of the given size at the caret. */
export function insertTableAt(env: CommandEnv, rows: number, cols: number): string | undefined {
  const doc = env.doc
  const table = {
    kind: 'tbl' as const,
    id: `t${Date.now().toString(36)}`,
    rows: Array.from({ length: rows }, () => Array.from({ length: cols }, () => ({
      paras: [{ kind: 'p' as const, id: `p${Math.random().toString(36).slice(2, 8)}`, runs: [] }],
    }))),
    widths: Array.from({ length: cols }, () => 96),
  }
  const at = env.loc?.blockIndex ?? doc.blocks.length - 1
  doc.blocks.splice(at + 1, 0, table)
  return undefined
}

/** Set the number of newspaper-style columns in the section. */
export function setColumns(env: CommandEnv, count: number): string | undefined {
  env.doc.section.columns = Math.max(1, Math.min(4, count))
  return undefined
}

/* ------------------------------------------------------------------ drawing */

/** The picture the drawing and picture panels act on. */
export function selectedImage(env: CommandEnv): ImageRef | null {
  const runs = env.loc?.para?.runs || []
  const found = runs.find((r) => r.image)?.image
  if (found) return found
  for (const block of env.doc.blocks) {
    if (!isPara(block)) continue
    const image = block.runs.find((r) => r.image)?.image
    if (image) return image
  }
  return null
}

/** The text-wrapping modes the original offers, with their labels. */
export const WRAP_MODES: [string, string][] = [
  ['inline', 'In Line with Text'], ['square', 'Square'], ['tight', 'Tight'],
  ['behind', 'Behind Text'], ['none', 'In Front of Text'], ['topBottom', 'Top and Bottom'],
  ['through', 'Through'],
]

/** The LettersArt shapes the original offers, with their labels. */
export const WORDART_SHAPES: [string, string][] = [
  ['textPlain', 'Plain Text'], ['textArchUp', 'Arch Up'], ['textArchDown', 'Arch Down'],
  ['textCircle', 'Circle'], ['textWave1', 'Wave'], ['textSlantUp', 'Slant Up'],
  ['textSlantDown', 'Slant Down'], ['textTriangle', 'Triangle Up'],
  ['textChevron', 'Chevron Up'], ['textCanUp', 'Can Up'],
]

/** The border styles the original offers, with the dash pattern each draws. */
export const BORDER_STYLES: [string, string | null][] = [
  ['single', null], ['dotted', '1 2'], ['dashed', '4 3'], ['dotDash', '4 3 1 3'],
  ['double', null], ['thick', null], ['triple', null], ['wave', null], ['threeDEmboss', null],
]

/** The line weights the original offers, in points. */
export const LINE_WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6]

/** The dash styles the original offers for drawing objects. */
export const DASH_STYLES: [string, string | null][] = [
  ['solid', null], ['sysDot', '1 1'], ['sysDash', '3 1'], ['dash', '4 3'],
  ['dashDot', '4 3 1 3'], ['lgDash', '8 3'], ['lgDashDot', '8 3 1 3'], ['lgDashDotDot', '8 3 1 3 1 3'],
]

/** The AutoText entries the original offers in the header and footer toolbar.
 *  The names are the original's own (public/quire/js/autocorrect.js → AUTOTEXT['Header/Footer']),
 *  so they translate through the same dictionary; the text is the field sequence they insert. */
export const HEADER_FOOTER_AUTOTEXT: [string, string][] = [
  ['- PAGE -', '- «PAGE» -'], ['Author, Page #, Date', '«AUTHOR», Page «PAGE», «DATE»'],
  ['Confidential, Page #, Date', 'Confidential, Page «PAGE», «DATE»'],
  ['Created by', 'Created by «AUTHOR»'], ['Created on', 'Created on «CREATEDATE»'],
  ['Filename', '«FILENAME»'], ['Filename and path', '«FILENAMEPATH»'],
  ['Last printed', 'Last printed «PRINTDATE»'], ['Last saved by', 'Last saved by «LASTSAVEDBY»'],
  ['Page X of Y', 'Page «PAGE» of «NUMPAGES»'],
]

/** Insert an AutoText entry into the header or the footer. */
export function insertAutoText(env: CommandEnv, text: string): string | undefined {
  const hf = (env.doc.headerFooter ||= {})
  hf.header = (hf.header || '') + text
  return undefined
}

/* -------------------------------------------------------------- the commands */

export const toolbarCommands: Record<string, CommandFn> = {
  applyBorder: (env) => applyBorder(env, env.arg || 'outside'),
  applyShading: (env) => applyShading(env, env.arg === 'none' ? null : env.arg || null),
  applyFontColor: (env) => applyFontColor(env, env.arg === 'auto' ? null : env.arg || null),
  applyHighlight: (env) => applyHighlight(env, env.arg || 'none'),
  setLineSpacing: (env) => setLineSpacing(env, Number(env.arg) || 1),
  insertTableAt: (env) => insertTableAt(env, Number(env.arg?.split('x')[0]) || 3, Number(env.arg?.split('x')[1]) || 3),
  setColumns: (env) => setColumns(env, Number(env.arg) || 1),
  setBorderStyle: (env) => { borderPen.val = env.arg || 'single'; return undefined },
  setBorderWeight: (env) => { borderPen.sz = Number(env.arg) || 0.5; return undefined },
  setBorderColor: (env) => { borderPen.color = env.arg || 'auto'; return undefined },
  setWrap: (env) => {
    const img = selectedImage(env)
    if (!img) return 'Select a picture first.'
    img.wrap = (env.arg || 'inline') as ImageRef['wrap']
    return undefined
  },
  setWordArtShape: (env) => {
    const para = env.loc?.para
    if (!para) return 'Place the cursor on the page first.'
    const run = para.runs.find((r) => r.text && r.rPr?.size && r.rPr.size >= 24)
    if (run) run.rPr = { ...run.rPr, warp: env.arg }
    return undefined
  },
  insertAutoText: (env) => insertAutoText(env, env.arg || ''),
  setCellFill: (env) => applyShading(env, env.arg || null),
  setOutlineLevel: (env) => {
    const para = env.loc?.para
    if (!para) return 'Place the cursor in a paragraph first.'
    const level = Number(env.arg) || 0
    para.pPr = { ...para.pPr, style: level ? `Heading${Math.min(3, level)}` : 'Normal', outlineLvl: level ? level - 1 : undefined }
    return undefined
  },
  goToRecord: (env) => {
    const merge = env.doc.merge
    if (!merge?.rows?.length) return 'No data source loaded.'
    const index = Math.max(0, Math.min(merge.rows.length - 1, (Number(env.arg) || 1) - 1))
    merge.index = index
    return `Record ${index + 1} of ${merge.rows.length}.`
  },
  setPicMode: (env) => {
    const img = selectedImage(env)
    if (!img) return 'Select a picture first.'
    img.gray = env.arg === 'gray' ? true : undefined
    img.bw = env.arg === 'bw' ? true : undefined
    img.washout = env.arg === 'washout' ? true : undefined
    return undefined
  },
  insertShape: (env) => {
    const para = env.loc?.para
    if (!para) return 'Place the cursor on the page first.'
    para.runs.push({
      image: { rId: '', part: 'word/media/shape.svg', url: shapeSvg(env.arg || 'rect'), w: 120, h: 60, wrap: 'inline' },
    })
    return undefined
  },
  setFillColor: (env) => {
    if (env.arg && env.arg !== 'none') shapeFill = env.arg
    else if (env.arg === 'none') shapeFill = 'none'
    return `Fill colour ${shapeFill}.`
  },
  setLineColor: (env) => {
    if (env.arg && env.arg !== 'none') shapeLine = env.arg
    else if (env.arg === 'none') shapeLine = 'none'
    return `Line colour ${shapeLine}.`
  },
  setLineStyle: (env) => { lineStyleState = env.arg || 'solid'; return `Line style ${lineStyleState}.` },
  setDashStyle: (env) => { dashStyleState = env.arg || 'solid'; return `Dash style ${dashStyleState}.` },
  setArrowStyle: (env) => { arrowStyleState = env.arg || 'none'; return `Arrow style ${arrowStyleState}.` },
  setShadowStyle: (env) => { shadowStyleState = env.arg || 'none'; return `Shadow ${shadowStyleState}.` },
  /* the review display modes the reviewing toolbar's dropdown offers */
  markupFinal: (env) => { env.doc.markupMode = 'final'; return undefined },
  markupOriginal: (env) => { env.doc.markupMode = 'original'; return undefined },
  markupOriginalMarkup: (env) => { env.doc.markupMode = 'originalMarkup'; return undefined },
  /* the nudge commands the draw menu offers */
  nudgeUp: (env) => nudge(env, 0, -4),
  nudgeDown: (env) => nudge(env, 0, 4),
  nudgeLeft: (env) => nudge(env, -4, 0),
  nudgeRight: (env) => nudge(env, 4, 0),
  setDefaults: () => undefined,
}

/** Move the picture at the caret by a few pixels. */
function nudge(env: CommandEnv, dx: number, dy: number): string | undefined {
  const img = selectedImage(env)
  if (!img) return 'Select an object first.'
  img.offset = { x: (img.offset?.x || 0) + dx, y: (img.offset?.y || 0) + dy }
  return undefined
}

/* The drawing tools insert real SVG shapes as picture runs, so they appear on the
 * page and round-trip through DOCX like any other image. */
let shapeFill = '#FFFFFF'
let shapeLine = '#1D4F9E'
let lineStyleState = 'solid'
let dashStyleState = 'solid'
let arrowStyleState = 'none'
let shadowStyleState = 'none'

/** The SVG for a drawing shape, in the colours the drawing toolbar set. */
function shapeSvg(kind: string): string {
  const fill = shapeFill === 'none' ? 'none' : shapeFill
  const line = shapeLine === 'none' ? 'none' : shapeLine
  /* a gallery preset: the original's own geometry, painted in the toolbar's colours */
  const preset = SHAPE_BY_GEOM[kind]
  if (preset) {
    const body = preset.parts.map((p) => `<path d="${p.d}" fill="${p.fill === 'none' ? 'none' : fill}"${p.rule ? ` fill-rule="${p.rule}"` : ''} stroke="${p.stroke ? line : 'none'}" stroke-width="2"/>`).join('')
    return 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="120" height="60" viewBox="0 0 100 100" preserveAspectRatio="xMidYMid meet">${body}</svg>`)
  }
  const body = kind === 'oval' ? `<ellipse cx="60" cy="30" rx="58" ry="28" fill="${fill}" stroke="${line}" stroke-width="2"/>`
    : kind === 'arrow' ? `<line x1="4" y1="30" x2="106" y2="30" stroke="${line}" stroke-width="2"/><polygon points="106,30 94,24 94,36" fill="${line}"/>`
      : kind === 'dblArrow' ? `<line x1="4" y1="30" x2="116" y2="30" stroke="${line}" stroke-width="2"/><polygon points="4,30 16,24 16,36" fill="${line}"/><polygon points="116,30 104,24 104,36" fill="${line}"/>`
        : `<rect x="2" y="2" width="116" height="56" fill="${fill}" stroke="${line}" stroke-width="2"/>`
  return 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="120" height="60" viewBox="0 0 120 60">${body}</svg>`)
}
