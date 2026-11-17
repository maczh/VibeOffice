/* Online connectors against a local mock server; the S3 SigV4 signature is cross-checked
 * with an independent node:crypto implementation of the same algorithm. */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import http from 'node:http'
import { createHmac, createHash } from 'node:crypto'
import { fetchOnline, saveOnline, listOnline } from '../dist/index.js'

const sha = (d) => createHash('sha256').update(d).digest('hex')
const hmac = (k, d) => createHmac('sha256', k).update(d).digest()

function mockServer(handler) {
  return new Promise((resolve) => {
    const srv = http.createServer(handler)
    srv.listen(0, '127.0.0.1', () => resolve({ srv, base: `http://127.0.0.1:${srv.address().port}` }))
  })
}
const close = (srv) => new Promise((r) => srv.close(r))

test('webdav GET with Basic auth', async () => {
  let auth = null
  const { srv, base } = await mockServer((req, res) => {
    auth = req.headers.authorization
    res.writeHead(200, { 'Content-Type': 'application/octet-stream' })
    res.end('FILEBYTES')
  })
  const f = await fetchOnline({ kind: 'webdav', url: base + '/docs/a.docx', username: 'u', password: 'p' })
  assert.equal(await f.blob.text(), 'FILEBYTES')
  assert.equal(f.name, 'a.docx')
  assert.equal(auth, 'Basic ' + Buffer.from('u:p').toString('base64'))
  await close(srv)
})

test('webdav PUT and PROPFIND list', async () => {
  let putBody = null, putMethod = null
  const propfind = `<?xml version="1.0"?><multistatus xmlns="DAV:">
    <response><href>/docs/</href><propstat><prop><resourcetype><collection/></resourcetype></prop></propstat></response>
    <response><href>/docs/a.docx</href><propstat><prop><displayname>a.docx</displayname><getcontentlength>9</getcontentlength><resourcetype/></prop></propstat></response>
    <response><href>/docs/sub/</href><propstat><prop><resourcetype><collection/></resourcetype></prop></propstat></response>
  </multistatus>`
  const { srv, base } = await mockServer(async (req, res) => {
    if (req.method === 'PUT') {
      putMethod = 'PUT'
      const chunks = []
      for await (const c of req) chunks.push(c)
      putBody = Buffer.concat(chunks).toString()
      res.writeHead(201); res.end()
    } else if (req.method === 'PROPFIND') {
      assert.equal(req.headers.depth, '1')
      res.writeHead(207, { 'Content-Type': 'application/xml' }); res.end(propfind)
    } else { res.writeHead(405); res.end() }
  })
  await saveOnline({ kind: 'webdav', url: base + '/docs/a.docx', token: 'T' }, new Blob(['NEWBYTES']), 'a.docx')
  assert.equal(putMethod, 'PUT')
  assert.equal(putBody, 'NEWBYTES')
  const entries = await listOnline({ kind: 'webdav', url: base + '/docs/' })
  assert.equal(entries.length, 2)
  assert.deepEqual(entries.map((e) => [e.name, e.isDir, e.size]), [['a.docx', false, 9], ['sub', true, undefined]])
  await close(srv)
})

test('s3 GET/PUT with SigV4 cross-checked by node:crypto', async () => {
  const creds = { accessKeyId: 'AKIDEXAMPLE', secretAccessKey: 'wJalrXUtnFEMI/K7MDENG+bPxRfiCYEXAMPLEKEY' }
  const region = 'us-east-1'
  let captured = null
  const { srv, base } = await mockServer(async (req, res) => {
    const chunks = []
    for await (const c of req) chunks.push(c)
    captured = { method: req.method, url: req.url, headers: req.headers, body: Buffer.concat(chunks) }
    if (req.method === 'GET' && req.url.startsWith('/bucket/key.txt')) { res.writeHead(200); res.end('S3DATA') }
    else if (req.method === 'PUT') { res.writeHead(200); res.end() }
    else { res.writeHead(404); res.end() }
  })
  const src = { kind: 's3', endpoint: base, bucket: 'bucket', key: 'key.txt', region, ...creds, forcePathStyle: true }
  const f = await fetchOnline(src)
  assert.equal(await f.blob.text(), 'S3DATA')
  // cross-check the GET signature independently
  const check = (cap, body) => {
    const h = cap.headers
    const amzDate = h['x-amz-date']
    const dateStamp = amzDate.slice(0, 8)
    const u = new URL('http://x' + cap.url)
    const canonicalHeaders = ['host', 'x-amz-content-sha256', 'x-amz-date']
      .filter((k) => h[k] !== undefined).sort()
      .map((k) => `${k}:${String(h[k]).trim()}\n`).join('')
    const signedHeaders = ['host', 'x-amz-content-sha256', 'x-amz-date'].filter((k) => h[k] !== undefined).sort().join(';')
    const q = [...u.searchParams.entries()].sort(([a], [b]) => a < b ? -1 : 1).map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`).join('&')
    const payloadHash = sha(body)
    assert.equal(h['x-amz-content-sha256'], payloadHash)
    const canonicalRequest = [cap.method, u.pathname, q, canonicalHeaders, signedHeaders, payloadHash].join('\n')
    const scope = `${dateStamp}/${region}/s3/aws4_request`
    const sts = ['AWS4-HMAC-SHA256', amzDate, scope, sha(canonicalRequest)].join('\n')
    const kSigning = hmac(hmac(hmac(hmac('AWS4' + creds.secretAccessKey, dateStamp), region), 's3'), 'aws4_request')
    const sig = createHmac('sha256', kSigning).update(sts).digest('hex')
    assert.ok(h.authorization.includes(`Signature=${sig}`), 'SigV4 signature must match independent implementation')
    assert.ok(h.authorization.includes(`Credential=${creds.accessKeyId}/${scope}`))
  }
  check(captured, Buffer.alloc(0))
  await saveOnline(src, new Blob(['NEW S3 BYTES']), 'key.txt')
  assert.equal(captured.method, 'PUT')
  assert.equal(captured.body.toString(), 'NEW S3 BYTES')
  check(captured, captured.body)
  await close(srv)
})

test('s3 list parses ListObjectsV2', async () => {
  const xmlBody = `<?xml version="1.0"?><ListBucketResult xmlns="http://s3.amazonaws.com/doc/2006-03-01/">
    <Contents><Key>docs/a.docx</Key><Size>12</Size><LastModified>2026-01-01T00:00:00Z</LastModified></Contents>
    <CommonPrefixes><Prefix>docs/sub/</Prefix></CommonPrefixes></ListBucketResult>`
  const { srv, base } = await mockServer((req, res) => {
    assert.ok(req.url.includes('list-type=2'))
    res.writeHead(200, { 'Content-Type': 'application/xml' }); res.end(xmlBody)
  })
  const entries = await listOnline({ kind: 's3', endpoint: base, bucket: 'b', key: '', region: 'us-east-1', accessKeyId: 'A', secretAccessKey: 'S', forcePathStyle: true }, 'docs/')
  assert.deepEqual(entries.map((e) => [e.name, e.isDir]), [['a.docx', false], ['sub', true]])
  await close(srv)
})

test('http GET with Content-Disposition name and POST multipart', async () => {
  let post = null
  const { srv, base } = await mockServer(async (req, res) => {
    if (req.method === 'POST') {
      const chunks = []
      for await (const c of req) chunks.push(c)
      post = { type: req.headers['content-type'], body: Buffer.concat(chunks).toString() }
      res.writeHead(200); res.end('ok')
    } else {
      res.writeHead(200, { 'Content-Disposition': 'attachment; filename="report.docx"' })
      res.end('HTTPFILE')
    }
  })
  const f = await fetchOnline({ kind: 'http', url: base + '/get' })
  assert.equal(f.name, 'report.docx')
  assert.equal(await f.blob.text(), 'HTTPFILE')
  await saveOnline({ kind: 'post', url: base + '/upload', field: 'doc', fields: { tag: 'x' } }, new Blob(['UPLOAD']), 'report.docx')
  assert.ok(post.type.startsWith('multipart/form-data'))
  assert.ok(post.body.includes('name="doc"'))
  assert.ok(post.body.includes('filename="report.docx"'))
  assert.ok(post.body.includes('UPLOAD'))
  assert.ok(post.body.includes('name="tag"'))
  await close(srv)
})

test('presigned open and save', async () => {
  let put = null
  const { srv, base } = await mockServer(async (req, res) => {
    if (req.method === 'PUT') { put = true; res.writeHead(200); res.end() }
    else { res.writeHead(200); res.end('PRESIGNED') }
  })
  const f = await fetchOnline({ kind: 'presigned', url: base + '/obj?sig=1' })
  assert.equal(await f.blob.text(), 'PRESIGNED')
  await saveOnline({ kind: 'presigned', url: base + '/obj?sig=1', putUrl: base + '/obj?sig=2' }, new Blob(['X']), 'obj')
  assert.ok(put)
  await assert.rejects(() => saveOnline({ kind: 'presigned', url: base + '/obj' }, new Blob(['X']), 'obj'), /read-only/)
  await close(srv)
})
