/* Shared core tests: run against the built library (node --test voffice-shared/test/). */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { zip, xml, Package, newPackage, formats, resolveTarget, xmlPart, uuid } from '../dist/index.js'

test('crc32 known vector', () => {
  assert.equal(zip.crc32(new TextEncoder().encode('123456789')), 0xcbf43926)
})

test('zip round trip with deflate and store', async () => {
  const big = 'hello voffice '.repeat(500)
  const blob = await zip.writeZip([
    { name: 'a.txt', data: big },
    { name: 'dir/b.bin', data: new Uint8Array([0, 1, 2, 255]) },
    { name: 'stored.xml', data: '<x/>', store: true },
  ])
  const files = await zip.readZip(new Uint8Array(await blob.arrayBuffer()))
  assert.equal(await files.get('a.txt').text(), big)
  assert.deepEqual([...await files.get('dir/b.bin').bytes()], [0, 1, 2, 255])
  assert.equal(await files.get('stored.xml').text(), '<x/>')
})

test('zip rejects non-zip and detects OLE', async () => {
  await assert.rejects(() => zip.readZip(new TextEncoder().encode('not a zip at all, really not')), /not a valid Office Open XML/)
  const ole = new Uint8Array(64)
  ole.set([0xd0, 0xcf, 0x11, 0xe0], 0)
  await assert.rejects(() => zip.readZip(ole), (e) => e.code === 'ole')
})

test('xml parse, namespaces, entities, serialize', () => {
  const doc = xml.parseXml('<w:document xmlns:w="http://x/w" xmlns:r="http://x/r"><w:p w:val="a&amp;b"><w:t>hi &lt;there&gt;</w:t></w:p></w:document>')
  assert.equal(doc.localName, 'document')
  const p = doc.children[0]
  assert.equal(p.getAttribute('w:val'), 'a&b')
  assert.equal(p.namespaceURI, 'http://x/w')
  assert.equal(p.children[0].textContent, 'hi <there>')
  const out = xml.serializeXml(doc)
  assert.ok(out.includes('xmlns:w="http://x/w"'))
  const again = xml.parseXml(out)
  assert.equal(again.children[0].children[0].textContent, 'hi <there>')
})

test('xml escText protects control chars and _x escapes', () => {
  assert.equal(xml.escText('a_x0041_b'), 'a_x005F_x0041_b')
  assert.equal(xml.unx('a_x0041_b'), 'aAb')
  assert.ok(xml.escText('a\x01b').includes('_x0001_'))
})

test('opc: new package, edit one part, others kept byte-for-byte', async () => {
  const pkg = newPackage('word/document.xml', xmlPart(xml.mkEl('w:document', { 'xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' })), formats.docx)
  const junk = new Uint8Array([1, 2, 3, 4, 5])
  pkg.addPart('word/media/image1.png', junk, 'image/png')
  const blob1 = await pkg.save()
  const p2 = await Package.open(new Uint8Array(await blob1.arrayBuffer()))
  assert.equal(p2.main, 'word/document.xml')
  assert.equal(p2.variant('docx'), 'docx')
  assert.deepEqual([...p2.bytes('word/media/image1.png')], [1, 2, 3, 4, 5])
  // edit the main part only
  p2.setPart('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>')
  const blob2 = await p2.save()
  const p3 = await Package.open(new Uint8Array(await blob2.arrayBuffer()))
  assert.deepEqual([...p3.bytes('word/media/image1.png')], [1, 2, 3, 4, 5], 'unedited part must survive unchanged')
  assert.ok(p3.text('word/document.xml').includes('<w:body/>'))
  assert.equal(p3.type('word/media/image1.png'), 'image/png')
})

test('opc: rels and target resolution', async () => {
  const pkg = newPackage('word/document.xml', '<w:document/>', formats.docx)
  const id = pkg.addRel('word/document.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'media/image1.png')
  assert.equal(id, 'rId1')
  const id2 = pkg.addRel('word/document.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', 'https://example.com/', 'External')
  assert.equal(id2, 'rId2')
  assert.equal(resolveTarget('word/document.xml', 'media/image1.png'), 'word/media/image1.png')
  assert.equal(resolveTarget('word/slides/slide1.xml', '../media/image1.png'), 'word/media/image1.png')
  assert.equal(resolveTarget('word/document.xml', '/ppt/media/x.png'), 'ppt/media/x.png')
  const blob = await pkg.save()
  const p2 = await Package.open(new Uint8Array(await blob.arrayBuffer()))
  const rels = p2.relsOf('word/document.xml')
  assert.equal(rels.length, 2)
  assert.equal(rels[1].mode, 'External')
  assert.deepEqual(p2.related('word/document.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'), ['word/media/image1.png'])
})

test('uuid shape', () => {
  assert.match(uuid(), /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/)
})
