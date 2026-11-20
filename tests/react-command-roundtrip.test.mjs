import assert from 'node:assert/strict'
import { test } from 'node:test'

test('command-created formatting and table state survives DOCX round-trip', async () => {
  const word = await import('@voffice/rword/engine')
  const doc = word.newDoc()
  doc.blocks[0].runs = [{
    text: 'Formatted',
    rPr: { caps: 'all', hidden: true, underline: true, underlineStyle: 'double', vertAlign: 'superscript' },
  }]
  doc.blocks[0].pPr = {
    keepNext: true,
    outlineLvl: 0,
    borders: { bottom: { val: 'single', sz: 1.5, color: 'auto', space: 1 } },
  }
  doc.blocks.push({
    kind: 'tbl', id: 'cmd-table', widths: [144, 144], align: 'center', autoFit: 'content', headerRows: 1,
    rows: [
      [
        { paras: [{ kind: 'p', id: 'cmd-p1', runs: [{ text: 'A' }] }], vAlign: 'center' },
        { paras: [{ kind: 'p', id: 'cmd-p2', runs: [{ text: 'B' }] }], textDirection: 'vertical' },
      ],
      [
        { paras: [{ kind: 'p', id: 'cmd-p3', runs: [{ text: '1' }] }] },
        { paras: [{ kind: 'p', id: 'cmd-p4', runs: [{ text: '2' }] }] },
      ],
    ],
  })

  const blob = await word.writeDocx(doc)
  const { doc: reopened } = await word.readDocx(new Uint8Array(await blob.arrayBuffer()))
  const para = reopened.blocks[0]
  assert.equal(para.runs[0].rPr.caps, 'all')
  assert.equal(para.runs[0].rPr.hidden, true)
  assert.equal(para.runs[0].rPr.underlineStyle, 'double')
  assert.equal(para.runs[0].rPr.vertAlign, 'superscript')
  assert.equal(para.pPr.keepNext, true)
  assert.equal(para.pPr.outlineLvl, 0)
  /* paragraph borders are kept per edge, the way the toolbar sets them */
  assert.equal(para.pPr.borders.bottom.val, 'single')
  assert.equal(para.pPr.borders.bottom.sz, 1.5)
  assert.equal(para.pPr.borders.bottom.color, 'auto')
  assert.equal(para.pPr.borders.bottom.space, 1)

  const table = reopened.blocks[1]
  assert.equal(table.kind, 'tbl')
  assert.equal(table.align, 'center')
  assert.equal(table.autoFit, 'content')
  assert.equal(table.headerRows, 1)
  assert.equal(table.rows[0][0].vAlign, 'center')
  assert.equal(table.rows[0][1].textDirection, 'vertical')
})
