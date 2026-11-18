import assert from 'node:assert/strict'
import { test } from 'node:test'

test('React office engines round-trip minimal OOXML documents', async () => {
  const word = await import('@voffice/rword/engine')
  const doc = word.newDoc()
  doc.blocks[0].runs = [{ text: 'React Quire' }]
  const docx = await word.writeDocx(doc)
  const reopenedDoc = await word.readDocx(new Uint8Array(await docx.arrayBuffer()))
  assert.match(word.docText(reopenedDoc.doc), /React Quire/)

  const excel = await import('@voffice/rexcel/engine')
  const workbook = excel.newWorkbook()
  excel.setCell(workbook.sheets[0], 0, 0, { v: 'React Ledger', t: 's' })
  const xlsx = await excel.writeXlsx(workbook)
  const reopenedWorkbook = await excel.readXlsx(new Uint8Array(await xlsx.arrayBuffer()))
  assert.equal(reopenedWorkbook.wb.sheets[0].cells.get('A1')?.v, 'React Ledger')

  const pptx = await import('@voffice/rpptx/engine')
  const presentation = pptx.newPresentation()
  presentation.slides[0].shapes.push(pptx.newTextShape(0, 0, 1000, 1000, 'React Lectern'))
  const deck = await pptx.writePptx(presentation)
  const reopenedDeck = await pptx.readPptx(new Uint8Array(await deck.arrayBuffer()))
  assert.match(reopenedDeck.pres.slides[0].shapes.flatMap((s) => s.text ?? []).flatMap((p) => p.runs.map((r) => r.text)).join(''), /React Lectern/)
})
