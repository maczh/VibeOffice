/* VWord document model: a JSON-serialisable tree of blocks (paragraphs and tables).
 * Runs carry character formatting; unmodelable content keeps its raw XML for lossless write-back. */

export interface RPr {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  font?: string
  /** point size */
  size?: number
  color?: string
  highlight?: string
  vertAlign?: 'superscript' | 'subscript'
}

export interface ImageRef {
  rId: string
  part: string
  url: string
  /** display size in px */
  w: number
  h: number
}

export interface Run {
  text?: string
  br?: boolean
  tab?: boolean
  image?: ImageRef
  hyperlink?: { rId: string; url: string }
  rPr?: RPr
  /** verbatim XML for content we do not model (fields, drawings, …) */
  raw?: string
}

export interface PPr {
  align?: 'left' | 'center' | 'right' | 'justify'
  style?: string
  numId?: number
  ilvl?: number
  /** left indent in px */
  indent?: number
  hanging?: number
  pageBreakBefore?: boolean
  spaceBefore?: number
  spaceAfter?: number
  line?: number
}

export interface Para {
  kind: 'p'
  id: string
  runs: Run[]
  pPr?: PPr
  raw?: string
}

export interface Cell {
  paras: Para[]
  gridSpan?: number
  vMerge?: 'restart' | 'continue'
  fill?: string
  width?: number
  raw?: string
}

export interface Table {
  kind: 'tbl'
  id: string
  rows: Cell[][]
  /** column widths in px */
  widths: number[]
  raw?: string
}

export type Block = Para | Table

export interface Section {
  /** twips (1/1440 inch) */
  pageW: number
  pageH: number
  margins: { top: number; right: number; bottom: number; left: number }
}

export interface StyleDef { id: string; name: string; basedOn?: string; rPr?: RPr; pPr?: PPr }

export interface Doc {
  blocks: Block[]
  section: Section
  styles: StyleDef[]
  defaultFont: string
  defaultSize: number
  /** numId → level-0 number format (bullet/decimal/…), from numbering.xml */
  numbering?: Record<number, string>
  format: string
}

let seq = 0
export const uid = (p = 'b'): string => `${p}${++seq}_${Math.random().toString(36).slice(2, 7)}`

export const emptyPara = (pPr?: PPr): Para => ({ kind: 'p', id: uid('p'), runs: [], pPr })

export const defaultSection = (): Section => ({
  pageW: 12240, pageH: 15840,
  margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
})

export const newDoc = (): Doc => ({
  blocks: [emptyPara()],
  section: defaultSection(),
  styles: [
    { id: 'Normal', name: 'Normal' },
    { id: 'Heading1', name: 'Heading 1', rPr: { bold: true, size: 16 }, pPr: { spaceBefore: 12, spaceAfter: 6 } },
    { id: 'Heading2', name: 'Heading 2', rPr: { bold: true, size: 13 }, pPr: { spaceBefore: 10, spaceAfter: 4 } },
    { id: 'Heading3', name: 'Heading 3', rPr: { bold: true, size: 12 }, pPr: { spaceBefore: 8, spaceAfter: 4 } },
  ],
  defaultFont: 'Calibri',
  defaultSize: 11,
  format: 'docx',
})

export const isPara = (b: Block): b is Para => b.kind === 'p'
export const isTable = (b: Block): b is Table => b.kind === 'tbl'

/** plain text of a block (for word count / find) */
export function blockText(b: Block): string {
  if (isTable(b)) return b.rows.map((r) => r.map((c) => c.paras.map(paraText).join(' ')).join('\t')).join('\n')
  return paraText(b)
}
export function paraText(p: Para): string {
  return p.runs.map((r) => r.text ?? (r.tab ? '\t' : r.br ? '\n' : '')).join('')
}
export function docText(d: Doc): string { return d.blocks.map(blockText).join('\n') }

export function wordCount(d: Doc): number {
  const t = docText(d).trim()
  return t ? t.split(/\s+/).length : 0
}

/** deep clone via JSON (model is plain data; image blob URLs survive as strings) */
export const cloneDoc = (d: Doc): Doc => JSON.parse(JSON.stringify(d))
