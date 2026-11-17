/* VPptx presentation model. Shapes carry EMU geometry; text is paragraphs of runs.
 * Unmodelable shapes keep raw XML for lossless write-back. */

export interface RPr { bold?: boolean; italic?: boolean; underline?: boolean; size?: number; color?: string; font?: string }
export interface Run { text: string; rPr?: RPr }
export interface Para { align?: 'left' | 'center' | 'right'; runs: Run[]; bullet?: boolean; lvl?: number }
export interface Fill { color: string; alpha?: number }
export interface Line { color: string; width?: number; dash?: string }

export interface Shape {
  id: string
  kind: 'sp' | 'pic' | 'cxnSp' | 'graphicFrame'
  /** EMU */
  x: number; y: number; w: number; h: number
  rot?: number
  preset?: string
  fill?: Fill | null
  line?: Line | null
  text?: Para[]
  image?: { part: string; rId: string; url: string }
  raw?: string
  name?: string
}

export interface Slide {
  id: string
  shapes: Shape[]
  notes?: string
  background?: string
  part?: string
  raw?: string
}

export interface Presentation {
  slideW: number
  slideH: number
  slides: Slide[]
  format: string
}

let seq = 0
export const uid = (p = 's'): string => `${p}${++seq}_${Math.random().toString(36).slice(2, 7)}`

export const EMU_PER_PX = 914400 / 96
export const px = (emu: number): number => emu / EMU_PER_PX
export const emu = (pxv: number): number => Math.round(pxv * EMU_PER_PX)

export const newSlide = (): Slide => ({ id: uid('sl'), shapes: [] })

export const newPresentation = (): Presentation => ({
  slideW: 12192000, slideH: 6858000,
  slides: [newSlide()],
  format: 'pptx',
})

export function newTextShape(x: number, y: number, w: number, h: number, text = ''): Shape {
  return {
    id: uid('sh'), kind: 'sp', x, y, w, h, preset: 'rect', fill: null, line: null,
    text: [{ align: 'left', runs: text ? [{ text }] : [] }],
  }
}

export const clonePresentation = (p: Presentation): Presentation => JSON.parse(JSON.stringify(p))