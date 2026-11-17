/* Engine regression tests for the Vue rewrite (headless, against the built ./engine bundles).
 * Run: node --test tests/vue/engines.test.mjs   (after npm run build) */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import * as W from '@voffice/vword/engine'
import * as X from '@voffice/vexcel/engine'
import * as P from '@voffice/vpptx/engine'

/* ---------- VExcel formula engine + recalc ---------- */
test('formula: arithmetic, precedence, concat, errors', () => {
  const wb = X.newWorkbook()
  const s = wb.sheets[0]
  const set = (a, cell) => s.cells.set(a, cell)
  set('A1', { v: 5, t: 'n' }); set('A2', { v: 10, t: 'n' })
  set('A3', { f: 'SUM(A1:A2)' })
  set('B1', { f: 'A3*2+1' })
  set('C1', { f: 'IF(A1>3,"big","small")' })
  set('D1', { f: '"a"&"b"' })
  set('E1', { f: '1/0' })
  set('F1', { f: '2^3^2' }) // right-assoc: 2^(3^2)=512
  X.recalc(wb)
  assert.equal(s.cells.get('A3').v, 15)
  assert.equal(s.cells.get('B1').v, 31)
  assert.equal(s.cells.get('C1').v, 'big')
  assert.equal(s.cells.get('D1').v, 'ab')
  assert.equal(s.cells.get('E1').v, '#DIV/0!')
  assert.equal(s.cells.get('F1').v, 512)
})

test('formula: VLOOKUP exact + circular reference', () => {
  const wb = X.newWorkbook()
  const s = wb.sheets[0]
  s.cells.set('A10', { v: 1, t: 'n' }); s.cells.set('B10', { v: 'one', t: 's' })
  s.cells.set('A11', { v: 2, t: 'n' }); s.cells.set('B11', { v: 'two', t: 's' })
  s.cells.set('F1', { f: 'VLOOKUP(2,A10:B11,2,FALSE)' })
  s.cells.set('G1', { f: 'H1' }); s.cells.set('H1', { f: 'G1' })
  X.recalc(wb)
  assert.equal(s.cells.get('F1').v, 'two')
  assert.equal(s.cells.get('G1').v, '#REF!')
  assert.equal(s.cells.get('H1').v, '#REF!')
})

test('formula: cross-sheet reference', () => {
  const wb = X.newWorkbook()
  wb.sheets.push(X.newSheet('Data'))
  wb.sheets[1].cells.set('A1', { v: 42, t: 'n' })
  wb.sheets[0].cells.set('A1', { f: 'Data!A1+1' })
  X.recalc(wb)
  assert.equal(wb.sheets[0].cells.get('A1').v, 43)
})

test('numfmt: general, grouped, percent, date', () => {
  assert.equal(X.formatValue(1234.5, 'General'), '1234.5')
  assert.equal(X.formatValue(1234.5, '#,##0.00'), '1,234.50')
  assert.equal(X.formatValue(0.25, '0%'), '25%')
  assert.equal(X.formatValue(1234.5, '$#,##0.00'), '$1,234.50')
  const serial = X.excelDate(new Date(2020, 0, 1))
  assert.equal(X.formatValue(serial, 'm/d/yy'), '1/1/20')
  assert.equal(X.formatValue(serial, 'yyyy-mm-dd'), '2020-01-01')
})

/* ---------- VExcel XLSX / CSV round trip ---------- */
test('xlsx round trip preserves values, formula, style, merge', async () => {
  const wb = X.newWorkbook()
  const s = wb.sheets[0]
  s.cells.set('A1', { v: 5, t: 'n' })
  s.cells.set('A2', { v: 10, t: 'n' })
  s.cells.set('A3', { f: 'SUM(A1:A2)' })
  s.cells.set('B1', { v: 'hello', t: 's', s: X.styleIndex(wb, { font: { bold: true, size: 14 }, fill: '#ff0000' }) })
  s.merges.push({ r1: 0, c1: 2, r2: 0, c2: 3 })
  s.colWidths.set(0, 120)
  const blob = await X.writeXlsx(wb)
  const res = await X.readXlsx(new Uint8Array(await blob.arrayBuffer()))
  X.recalc(res.wb)
  const rs = res.wb.sheets[0]
  assert.equal(rs.cells.get('A1').v, 5)
  assert.equal(rs.cells.get('A3').v, 15)
  assert.equal(rs.cells.get('B1').v, 'hello')
  assert.equal(res.wb.styles[rs.cells.get('B1').s].font.bold, true)
  assert.equal(res.wb.styles[rs.cells.get('B1').s].fill, '#ff0000')
  assert.equal(rs.merges.length, 1)
  assert.equal(rs.colWidths.get(0), 120)
})

test('csv parse and write', () => {
  const wb = X.parseCsv('name,qty\napple,3\n"pear, x",4')
  const s = wb.sheets[0]
  assert.equal(s.cells.get('A1').v, 'name')
  assert.equal(s.cells.get('B2').v, 3)
  assert.equal(s.cells.get('A3').v, 'pear, x')
  const out = X.writeCsv(wb, s)
  assert.ok(out.includes('"pear, x"'))
})

/* ---------- VWord DOCX round trip ---------- */
test('docx round trip: text, bold, table', async () => {
  const doc = W.newDoc()
  doc.blocks = [
    { kind: 'p', id: 'p1', runs: [{ text: 'Hello ' }, { text: 'Bold', rPr: { bold: true, size: 16, color: '#ff0000' } }], pPr: { align: 'center' } },
    { kind: 'tbl', id: 't1', widths: [100, 100], rows: [[{ paras: [{ kind: 'p', id: 'c1', runs: [{ text: 'a' }] }] }, { paras: [{ kind: 'p', id: 'c2', runs: [{ text: 'b' }] }] }]] },
  ]
  const blob = await W.writeDocx(doc)
  const res = await W.readDocx(new Uint8Array(await blob.arrayBuffer()))
  const p = res.doc.blocks[0]
  assert.equal(p.runs.map((r) => r.text).join(''), 'Hello Bold')
  assert.equal(p.runs[1].rPr.bold, true)
  assert.equal(p.runs[1].rPr.size, 16)
  assert.equal(p.pPr.align, 'center')
  const tbl = res.doc.blocks[1]
  assert.equal(tbl.kind, 'tbl')
  assert.equal(tbl.rows[0].length, 2)
})

test('docx corpus round trip preserves word count', async () => {
  const bytes = readFileSync(new URL('../../tests/quire/lines/linecases.docx', import.meta.url))
  const a = await W.readDocx(bytes)
  const before = W.wordCount(a.doc)
  const blob = await W.writeDocx(a.doc, { pkg: a.pkg, part: a.part })
  const b = await W.readDocx(new Uint8Array(await blob.arrayBuffer()))
  assert.equal(W.wordCount(b.doc), before)
  assert.ok(before > 0)
})

/* ---------- VPptx PPTX round trip ---------- */
test('pptx round trip: shapes, preset, text; new-deck part set', async () => {
  const pres = P.newPresentation()
  pres.slides[0].shapes.push(
    { id: 's1', kind: 'sp', x: 100000, y: 100000, w: 4000000, h: 2000000, preset: 'roundRect', fill: { color: '#4472c4' }, line: { color: '#000000', width: 1 }, text: [{ align: 'center', runs: [{ text: 'Title', rPr: { bold: true, size: 28 } }] }] },
    { id: 's2', kind: 'sp', x: 100000, y: 2500000, w: 3000000, h: 1500000, preset: 'ellipse', fill: null, line: null, text: [] },
  )
  const blob = await P.writePptx(pres)
  const res = await P.readPptx(new Uint8Array(await blob.arrayBuffer()))
  assert.equal(res.pres.slides.length, 1)
  const sh = res.pres.slides[0].shapes
  assert.equal(sh.length, 2)
  assert.equal(sh[0].preset, 'roundRect')
  assert.equal(sh[0].fill.color, '#4472c4')
  assert.equal(sh[0].text[0].runs[0].text, 'Title')
  assert.equal(sh[0].text[0].runs[0].rPr.bold, true)
  assert.equal(sh[1].preset, 'ellipse')
  // a brand-new deck must carry the master/layout/theme parts
  assert.ok(res.pkg.has('ppt/slideMasters/slideMaster1.xml'))
  assert.ok(res.pkg.has('ppt/slideLayouts/slideLayout1.xml'))
  assert.ok(res.pkg.has('ppt/theme/theme1.xml'))
})

test('pptx: picture bytes stored and re-read', async () => {
  const pres = P.newPresentation()
  // 1x1 transparent PNG
  const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
  pres.slides[0].shapes.push({ id: 'p1', kind: 'pic', x: 0, y: 0, w: 914400, h: 914400, image: { part: '', rId: '', url: png } })
  const blob = await P.writePptx(pres)
  const res = await P.readPptx(new Uint8Array(await blob.arrayBuffer()))
  const pic = res.pres.slides[0].shapes.find((s) => s.kind === 'pic')
  assert.ok(pic)
  assert.ok(pic.image.url.startsWith('blob:'))
  assert.ok(res.pkg.has(pic.image.part))
})