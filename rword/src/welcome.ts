/* The Quire welcome/sample document, built with the React document model. Mirrors
 * A.buildSample() in public/quire/js/app.js, including the styles and the text that
 * exercises fields, tabs and footnotes. */
import { emptyPara, newDoc, uid, type Doc, type Para, type Run, type RPr, type PPr, type Table } from './engine/model'

const text = (value: string, rPr?: RPr): Run => ({ text: value, rPr })

function para(runs: Run[] | string, pPr?: PPr): Para {
  const list = typeof runs === 'string' ? [text(runs)] : runs
  return { kind: 'p', id: uid('p'), runs: list, pPr }
}

function cell(value: string, header: boolean) {
  return { paras: [para(value ? [text(value, header ? { bold: true } : undefined)] : [])], width: 192, ...(header ? { fill: 'D9E2F3' } : {}) }
}

export function buildWelcomeDoc(): Doc {
  const doc = newDoc()
  doc.defaultFont = 'Times New Roman'
  doc.defaultSize = 12
  const replaced = new Set(['Title', 'Subtitle', 'Heading1', 'Heading2'])
  doc.styles = [
    ...doc.styles.filter((s) => !replaced.has(s.id)).map((s) => ({ ...s, rPr: { ...s.rPr, font: 'Times New Roman' } })),
    { id: 'Title', name: 'Title', rPr: { font: 'Arial', bold: true, size: 16, color: '000000' }, pPr: { align: 'center', spaceBefore: 12, spaceAfter: 3 } },
    { id: 'Subtitle', name: 'Subtitle', rPr: { font: 'Arial', italic: true, size: 12 }, pPr: { align: 'center', spaceAfter: 3 } },
    { id: 'Heading1', name: 'Heading 1', rPr: { font: 'Arial', bold: true, size: 16 }, pPr: { spaceBefore: 12, spaceAfter: 3 } },
    { id: 'Heading2', name: 'Heading 2', rPr: { font: 'Arial', bold: true, italic: true, size: 14 }, pPr: { spaceBefore: 12, spaceAfter: 3 } },
  ]
  doc.numbering = { 1: 'bullet' }
  const blocks: Doc['blocks'] = [
    para('Welcome to Quire', { style: 'Title' }),
    para('A word processor in the spirit of Word — it opens and saves real .docx files.', { style: 'Subtitle' }),
    para('Getting around', { style: 'Heading1' }),
    para([
      text('Type anywhere on this page. Everything works the way you remember: '),
      text('bold', { bold: true }), text(', '),
      text('italic', { italic: true }), text(', '),
      text('underline', { underline: true }), text(', '),
      text('highlighting', { highlight: 'yellow' }), text(', '),
      text('colors', { color: 'C00000' }),
      text(', styles, bullets, numbering, tables, headers and footers, footnotes, comments, tracked changes, fields and more.'),
    ], { align: 'justify' }),
    para('Open a document with File ▸ Open, or drop a .docx onto the window. Save with File ▸ Save (Ctrl+S); File ▸ Save As offers Word Document, Word Template, Web Page, Plain Text, Markdown and PDF.', { numId: 1, ilvl: 0, indent: 42 }),
    para('Switch views with the buttons in the lower-left corner: Normal, Web Layout, Print Layout, Outline and Reading Layout.', { numId: 1, ilvl: 0, indent: 42 }),
    para('Right-click for the shortcut menu; the task pane (Ctrl+F1) holds Styles and Formatting, Reveal Formatting, Clip Art and the Mail Merge wizard.', { numId: 1, ilvl: 0, indent: 42 }),
    para('A small table', { style: 'Heading2' }),
  ]
  const rows = [
    ['Feature', 'Menu', 'Shortcut'],
    ['Find and Replace', 'Edit', 'Ctrl+H'],
    ['Insert footnote', 'Insert ▸ Reference', 'Ctrl+Alt+F'],
    ['Track changes', 'Tools', 'Ctrl+Shift+E'],
  ]
  const table: Table = {
    kind: 'tbl', id: uid('tbl'), widths: [192, 192, 192],
    rows: rows.map((r, ri) => r.map((v) => cell(v, ri === 0))),
  }
  blocks.push(table)
  blocks.push(para([text('Tabs, leaders and fields'), text('1', { vertAlign: 'superscript' })], { style: 'Heading2' }))
  blocks.push(para('Chapter One' + ' '.repeat(58) + '·  7'))
  blocks.push(para('Chapter Two' + ' '.repeat(58) + '·  19'))
  const date = new Intl.DateTimeFormat('en-US', { month: 'long', day: 'numeric', year: 'numeric' }).format(new Date())
  blocks.push(para([text('This document was last printed on '), text(date)]))
  blocks.push(para(''))
  doc.blocks = blocks
  return doc
}

export const isWelcomeEmpty = (doc: Doc): boolean => doc.blocks.length === 1 && !doc.blocks[0]
