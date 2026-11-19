/* Fields, symbols, pictures, headers/footers and mail-merge commands. */
import { type Doc, type ImageRef, type Para, type Run, isPara, uid } from '../engine/model'
import { type CommandEnv, type CommandFn } from './documentCommands'
import { evaluateField } from './fields'

const insertText = (para: Para | null | undefined, text: string, rPr?: Run['rPr']): void => {
  if (!para) return
  para.runs.push({ text, ...(rPr ? { rPr } : {}) })
}

function fieldContext(env: CommandEnv) {
  return {
    page: 1, pages: 1, fileName: env.doc.format,
    title: env.doc.blocks.find(isPara)?.runs.map((r) => r.text || '').join('') || '',
    author: 'Author', now: new Date(),
  }
}

export const symbolCommands: Record<string, CommandFn> = {
  insertSymbol: (env) => { insertText(env.loc?.para, env.arg || '©'); return undefined },
  symCopyright: (env) => { insertText(env.loc?.para, '©'); return undefined },
  symRegistered: (env) => { insertText(env.loc?.para, '®'); return undefined },
  symTrademark: (env) => { insertText(env.loc?.para, '™'); return undefined },
  symEllipsis: (env) => { insertText(env.loc?.para, '…'); return undefined },
  emDash: (env) => { insertText(env.loc?.para, '—'); return undefined },
  enDash: (env) => { insertText(env.loc?.para, '–'); return undefined },
  nbHyphen: (env) => { insertText(env.loc?.para, '‑'); return undefined },
  optHyphen: (env) => { insertText(env.loc?.para, '\u00AD'); return undefined },
  insertBreak: (env) => {
    const para = env.loc?.para
    if (!para) return undefined
    if (env.arg === 'page') para.pPr = { ...para.pPr, pageBreakBefore: true }
    else para.runs.push({ br: true })
    return undefined
  },
  dateTime: (env) => { insertText(env.loc?.para, new Date().toLocaleString()); return undefined },
}

export const fieldCommands: Record<string, CommandFn> = {
  insertFieldBraces: (env) => { insertText(env.loc?.para, '{ FIELD }'); return undefined },
  hyphenation: (env) => { insertText(env.loc?.para, '\u00AD'); return undefined },
  versions: (env) => `${(env.doc.revisions || []).length} saved version(s).`,
  reviewersMenu: (env) => [...new Set((env.doc.comments || []).map((c) => c.author))].join(', ') || 'No reviewers.',
  insertField: (env) => {
    const instruction = env.arg || 'DATE'
    insertText(env.loc?.para, evaluateField(instruction, fieldContext(env)))
    return undefined
  },
  insertDateField: (env) => { insertText(env.loc?.para, evaluateField('DATE', fieldContext(env))); return undefined },
  insertTimeField: (env) => { insertText(env.loc?.para, evaluateField('TIME', fieldContext(env))); return undefined },
  insertPageField: (env) => { insertText(env.loc?.para, evaluateField('PAGE', fieldContext(env))); return undefined },
  pageNumbers: (env) => {
    const hf = (env.doc.headerFooter ||= {})
    hf.footer = (hf.footer || '') + '\tPage «PAGE» of «NUMPAGES»'
    return undefined
  },
  updateField: (env) => {
    const hf = env.doc.headerFooter
    if (hf) {
      hf.header = (hf.header || '').replace('«PAGE»', '1').replace('«NUMPAGES»', '1')
      hf.footer = (hf.footer || '').replace('«PAGE»', '1').replace('«NUMPAGES»', '1')
    }
    return 'Fields updated.'
  },
  insertFootnote: (env) => {
    const para = env.loc?.para
    if (!para) return undefined
    const n = (env.doc.revisions || []).length + 1
    insertText(para, String(n), { vertAlign: 'superscript' })
    return undefined
  },
  insertFootnoteNow: (env) => fieldCommands.insertFootnote!(env),
  insertEndnoteNow: (env) => fieldCommands.insertFootnote!(env),
  footnotes: (env) => `${(env.doc.revisions || []).length} note(s) in this document.`,
  insertCaption: (env) => {
    insertText(env.loc?.para, `Figure 1: ${env.selection || 'Caption'}`)
    return undefined
  },
  indexTables: (env) => {
    const count = env.doc.blocks.filter((b) => b.kind === 'tbl').length
    insertText(env.loc?.para, `Index of tables — ${count} table(s)`)
    return undefined
  },
}

function findImage(env: CommandEnv): ImageRef | null {
  const runs: Run[] = []
  if (env.loc?.para) runs.push(...env.loc.para.runs)
  if (!runs.length) for (const block of env.doc.blocks) if (isPara(block)) runs.push(...block.runs)
  return runs.find((r) => r.image)?.image || null
}

function imageCommand(env: CommandEnv, fn: (img: ImageRef) => void): string | undefined {
  const img = findImage(env)
  if (!img) return 'Select a picture first.'
  fn(img)
  return undefined
}

export const pictureCommands: Record<string, CommandFn> = {
  picMoreContrast: (env) => imageCommand(env, (img) => { img.contrast = Math.min(200, (img.contrast ?? 100) + 10) }),
  picLessContrast: (env) => imageCommand(env, (img) => { img.contrast = Math.max(0, (img.contrast ?? 100) - 10) }),
  picMoreBright: (env) => imageCommand(env, (img) => { img.brightness = Math.min(300, (img.brightness ?? 100) + 10) }),
  picLessBright: (env) => imageCommand(env, (img) => { img.brightness = Math.max(0, (img.brightness ?? 100) - 10) }),
  picCrop: (env) => imageCommand(env, (img) => { img.crop = { l: 5, r: 5, t: 5, b: 5 } }),
  picRotate: (env) => imageCommand(env, (img) => { img.rotation = ((img.rotation || 0) + 90) % 360 }),
  rotateLeft: (env) => imageCommand(env, (img) => { img.rotation = ((img.rotation || 0) + 270) % 360 }),
  rotateRight: (env) => imageCommand(env, (img) => { img.rotation = ((img.rotation || 0) + 90) % 360 }),
  picReset: (env) => imageCommand(env, (img) => {
    delete img.contrast; delete img.brightness; delete img.crop; delete img.rotation; img.wrap = 'inline'
  }),
  colorPic: (env) => imageCommand(env, (img) => { img.brightness = img.brightness === 50 ? 100 : 50 }),
  textWrap: (env) => imageCommand(env, (img) => {
    const modes: NonNullable<ImageRef['wrap']>[] = ['inline', 'square', 'tight', 'behind']
    img.wrap = modes[(modes.indexOf(img.wrap || 'inline') + 1) % modes.length]
  }),
  flipH: (env) => imageCommand(env, (img) => { img.rotation = img.rotation === 180 ? 0 : 180 }),
  flipV: (env) => imageCommand(env, (img) => { img.rotation = img.rotation === 180 ? 0 : 180 }),
  formatObject: (env) => imageCommand(env, () => undefined),
  bringForward: () => undefined,
  bringFront: () => undefined,
  sendBackward: () => undefined,
  sendBack: () => undefined,
  sendBehindText: (env) => imageCommand(env, (img) => { img.wrap = 'behind' }),
  bringFrontText: (env) => imageCommand(env, (img) => { img.wrap = 'square' }),
  groupObj: () => undefined,
  ungroupObj: () => undefined,
  selectObjects: () => 'Click an object on the page to select it.',
  newDrawing: (env) => { insertText(env.loc?.para, '[Drawing]'); return undefined },
  insertDiagram: (env) => { insertText(env.loc?.para, '[Diagram]'); return undefined },
  insertWordArt: (env) => {
    insertText(env.loc?.para, 'WordArt', { bold: true, size: 28, color: '#1D4F9E' })
    return undefined
  },
  wordartEdit: (env) => { insertText(env.loc?.para, env.arg || 'WordArt', { bold: true, size: 28 }); return undefined },
  wordartGallery: (env) => fieldCommands.insertWordArt!(env),
  waShape: (env) => fieldCommands.insertWordArt!(env),
  textBox: (env) => { insertText(env.loc?.para, '[Text Box]'); return undefined },
  insertChart: (env) => { insertText(env.loc?.para, '[Chart]'); return undefined },
  insertExcel: (env) => { insertText(env.loc?.para, '[Excel Sheet]'); return undefined },
  insertObject: (env) => { insertText(env.loc?.para, '[Object]'); return undefined },
  editObject: () => 'Double-click the object to edit it.',
}

export const headerFooterCommands: Record<string, CommandFn> = {
  headerFooter: (env) => { env.doc.headerFooter ||= {}; return undefined },
  hfPageNum: (env) => { const hf = (env.doc.headerFooter ||= {}); hf.header = (hf.header || '') + ' «PAGE»'; return undefined },
  hfNumPages: (env) => { const hf = (env.doc.headerFooter ||= {}); hf.header = (hf.header || '') + ' «NUMPAGES»'; return undefined },
  hfFormatPageNum: (env) => { const hf = (env.doc.headerFooter ||= {}); hf.header = (hf.header || '').replace('«PAGE»', '1'); return undefined },
  hfDate: (env) => { const hf = (env.doc.headerFooter ||= {}); hf.header = (hf.header || '') + ' ' + new Date().toLocaleDateString(); return undefined },
  hfTime: (env) => { const hf = (env.doc.headerFooter ||= {}); hf.header = (hf.header || '') + ' ' + new Date().toLocaleTimeString(); return undefined },
  hfAutoText: (env) => { const hf = (env.doc.headerFooter ||= {}); hf.header = (hf.header || '') + ' Quire'; return undefined },
  hfShowText: (env) => { const hf = (env.doc.headerFooter ||= {}); hf.showText = !hf.showText; return undefined },
  hfLinkPrev: (env) => { const hf = (env.doc.headerFooter ||= {}); hf.linkPrev = !hf.linkPrev; return undefined },
  hfSwitch: (env) => {
    const hf = (env.doc.headerFooter ||= {}); hf.kind = hf.kind === 'even' ? 'odd' : 'even'
    return undefined
  },
  hfPrev: () => 'Previous section.',
  hfNext: () => 'Next section.',
  hfClose: () => undefined,
  printedWatermark: (env) => { const hf = (env.doc.headerFooter ||= {}); hf.header = (hf.header || '') + ' DRAFT'; return undefined },
  backgroundNone: (env) => { env.doc.headerFooter = undefined; return undefined },
  moreColors: () => undefined,
  themeDlg: () => undefined,
}

export const mailMergeCommands: Record<string, CommandFn> = {
  mmMainDoc: (env) => { env.doc.merge ||= {}; return undefined },
  mmRecipients: (env) => `${(env.doc.merge?.rows || []).length} recipient(s).`,
  mmOpenSource: () => 'Choose a CSV file as the mail-merge data source.',
  mmAddress: (env) => { insertText(env.loc?.para, '«Address»'); return undefined },
  mmGreeting: (env) => { insertText(env.loc?.para, 'Dear «FirstName»,'); return undefined },
  mmInsertField: (env) => { insertText(env.loc?.para, `«${env.arg || 'FirstName'}»`); return undefined },
  mmViewData: (env) => {
    const rows = env.doc.merge?.rows || []
    const index = env.doc.merge?.index || 0
    return rows[index] ? Object.entries(rows[index]).map(([k, v]) => `${k}: ${v}`).join(', ') : 'No data source loaded.'
  },
  mmFirst: (env) => moveRecord(env, 'first'),
  mmPrev: (env) => moveRecord(env, 'prev'),
  mmNext: (env) => moveRecord(env, 'next'),
  mmLast: (env) => moveRecord(env, 'last'),
  mmMergeNew: (env) => {
    const rows = env.doc.merge?.rows || []
    if (!rows.length) return 'Load a CSV data source first.'
    replaceMergeFields(env.doc, rows[env.doc.merge?.index || 0])
    return `${rows.length} merged document(s) prepared.`
  },
  mmMergePrint: (env) => mailMergeCommands.mmMergeNew!(env),
  envelopes: () => undefined,
  letterWizard: () => undefined,
  templates: () => undefined,
  macros: () => 'Macros are not supported by the web build.',
  optionsDlg: () => undefined,
  customize: () => undefined,
  autocorrectDlg: () => undefined,
  newBlank: () => undefined,
  newWindow: () => undefined,
  arrangeAll: () => undefined,
  splitWindow: () => undefined,
  saveWeb: () => undefined,
  webPreview: () => undefined,
  exitApp: () => undefined,
  sendMail: () => undefined,
  permission: () => undefined,
}

function moveRecord(env: CommandEnv, to: 'first' | 'prev' | 'next' | 'last'): string | undefined {
  const rows = env.doc.merge?.rows || []
  if (!rows.length) return 'No data source loaded.'
  const current = env.doc.merge?.index || 0
  const index = to === 'first' ? 0 : to === 'last' ? rows.length - 1 : to === 'prev' ? Math.max(0, current - 1) : Math.min(rows.length - 1, current + 1)
  env.doc.merge = { ...env.doc.merge, index }
  replaceMergeFields(env.doc, rows[index])
  return `Record ${index + 1} of ${rows.length}.`
}

export function loadMergeSource(doc: Doc, text: string): number {
  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean)
  if (lines.length < 2) return 0
  const split = (line: string) => line.split(',').map((cell) => cell.replace(/^"|"$/g, '').trim())
  const headers = split(lines[0])
  const rows = lines.slice(1).map((line) => {
    const cells = split(line)
    return Object.fromEntries(headers.map((h, i) => [h, cells[i] || '']))
  })
  doc.merge = { ...doc.merge, rows, index: 0 }
  replaceMergeFields(doc, rows[0])
  return rows.length
}

function replaceMergeFields(doc: Doc, row: Record<string, string> | undefined): void {
  if (!row) return
  for (const block of doc.blocks) {
    if (!isPara(block)) continue
    for (const run of block.runs) {
      if (!run.text || !run.text.includes('«')) continue
      run.text = run.text.replace(/«([^»]+)»/g, (_, key: string) => row[key] ?? '')
    }
  }
}

/* Drawing tools insert real SVG shapes as picture runs, so they appear on the page
 * and round-trip through DOCX like any other image. */
let shapeFill = '#FFFFFF'
let shapeLine = '#1D4F9E'

function shapeUrl(kind: string): string {
  const fill = shapeFill
  const line = shapeLine
  const body = kind === 'oval' ? `<ellipse cx="60" cy="30" rx="58" ry="28" fill="${fill}" stroke="${line}" stroke-width="2"/>`
    : kind === 'arrow' ? `<line x1="4" y1="30" x2="106" y2="30" stroke="${line}" stroke-width="2"/><polygon points="106,30 94,24 94,36" fill="${line}"/>`
      : kind === 'dblArrow' ? `<line x1="4" y1="30" x2="116" y2="30" stroke="${line}" stroke-width="2"/><polygon points="4,30 16,24 16,36" fill="${line}"/><polygon points="116,30 104,24 104,36" fill="${line}"/>`
        : `<rect x="2" y="2" width="116" height="56" fill="${fill}" stroke="${line}" stroke-width="2"/>`
  return 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="120" height="60" viewBox="0 0 120 60">${body}</svg>`)
}

function insertShape(env: CommandEnv, kind: string): string | undefined {
  const para = env.loc?.para
  if (!para) return 'Place the cursor on the page first.'
  para.runs.push({
    image: { rId: '', part: 'word/media/shape.svg', url: shapeUrl(kind), w: 120, h: 60, wrap: 'inline' },
  })
  return undefined
}

let lineStyleState = 'solid'
let dashStyleState = 'solid'
let arrowStyleState = 'none'
let shadowStyleState = false

export const drawingCommands: Record<string, CommandFn> = {
  draw: (env) => { insertText(env.loc?.para, '[Drawing]'); return undefined },
  autoshapes: (env) => insertShape(env, 'rect'),
  drawLine: (env) => insertShape(env, 'line'),
  drawArrow: (env) => insertShape(env, 'arrow'),
  drawDblArrow: (env) => insertShape(env, 'dblArrow'),
  drawRect: (env) => insertShape(env, 'rect'),
  drawOval: (env) => insertShape(env, 'oval'),
  fillColor: (env) => { if (env.arg) shapeFill = env.arg; else shapeFill = shapeFill === '#FFFFFF' ? '#F5C842' : '#FFFFFF'; return `Fill colour ${shapeFill}.` },
  lineColor: (env) => { if (env.arg) shapeLine = env.arg; else shapeLine = shapeLine === '#1D4F9E' ? '#D33A2C' : '#1D4F9E'; return `Line colour ${shapeLine}.` },
  lineStyle: (env) => { lineStyleState = env.arg || (lineStyleState === 'solid' ? 'dashed' : 'solid'); return `Line style ${lineStyleState}.` },
  dashStyle: (env) => { dashStyleState = env.arg || (dashStyleState === 'solid' ? 'dash' : 'solid'); return `Dash style ${dashStyleState}.` },
  arrowStyle: (env) => { arrowStyleState = env.arg || (arrowStyleState === 'none' ? 'end' : 'none'); return `Arrow style ${arrowStyleState}.` },
  shadowStyle: () => { shadowStyleState = !shadowStyleState; return `Shadow ${shadowStyleState ? 'on' : 'off'}.` },
  layoutDlg: () => undefined,
}

export const paragraphDialogCommands: Record<string, CommandFn> = {
  columnsDlg: (env) => {
    const count = Math.max(1, Math.min(4, Number(env.arg) || (env.doc.section.columns || 1) + 1))
    env.doc.section.columns = count > 4 ? 1 : count
    return undefined
  },
  tabsDlg: (env) => {
    const para = env.loc?.para
    if (!para) return undefined
    const tab = Math.max(0, Number(env.arg) || 48)
    para.pPr = { ...para.pPr, indent: tab }
    return undefined
  },
  autoText: (env) => { insertText(env.loc?.para, 'Quire'); return undefined },
  showAssistant: () => 'Assistant shown.',
  insertFile: (env) => { insertText(env.loc?.para, env.arg || '[File]'); return undefined },
  setDefaults: () => undefined,
  newDrawing: (env) => insertShape(env, 'rect'),
}

export const auxiliaryCommands: Record<string, CommandFn> = {
  ...symbolCommands, ...fieldCommands, ...pictureCommands, ...headerFooterCommands,
  ...mailMergeCommands, ...drawingCommands, ...paragraphDialogCommands,
}

export const newRevisionId = (): string => uid('rev')
