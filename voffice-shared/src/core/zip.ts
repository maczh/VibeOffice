/* ZIP container engine (read + write), TypeScript rewrite of public/common/zip.js.
 * Uses the browser's Compression Streams when present and ships a small pure-JS inflater.
 *
 * The pure-JS inflate below is derived from puff contrib/puff/puff.c (zlib licence): Mark Adler's
 * puff.c translated to JavaScript and altered (typed arrays, a growing output buffer).
 * This is an altered version, not the original puff. Its notice:
 *
 *   Copyright (C) 2002-2013 Mark Adler, all rights reserved
 *   version 2.3, 21 Jan 2013
 *
 *   This software is provided 'as-is', without any express or implied
 *   warranty.  In no event will the author be held liable for any damages
 *   arising from the use of this software.
 *
 *   Permission is granted to anyone to use this software for any purpose,
 *   including commercial applications, and to alter it and redistribute it
 *   freely, subject to the following restrictions:
 *
 *   1. The origin of this software must not be misrepresented; you must not
 *      claim that you wrote the original software. If you use this software
 *      in a product, an acknowledgment in the product documentation would be
 *      appreciated but is not required.
 *   2. Altered source versions must be plainly marked as such, and must not be
 *      misrepresented as being the original software.
 *   3. This notice may not be removed or altered from any source distribution.
 *
 *   Mark Adler    madler@alumni.caltech.edu
 */

const te = new TextEncoder()
const td = new TextDecoder('utf-8')

/* ---------- CRC-32 ---------- */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256)
  for (let n = 0; n < 256; n++) {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    t[n] = c >>> 0
  }
  return t
})()

export function crc32(buf: Uint8Array): number {
  let c = 0xffffffff
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

/* ---------- pure-JS inflate (RFC 1951) ---------- */
const LBASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258]
const LEXT = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0]
const DBASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577]
const DEXT = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13]
const CLORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]

interface Huff { counts: Uint16Array; syms: Uint16Array }

function buildHuff(lengths: Uint8Array, n: number): Huff {
  const counts = new Uint16Array(16), offs = new Uint16Array(16), syms = new Uint16Array(n)
  for (let i = 0; i < n; i++) counts[lengths[i]]++
  counts[0] = 0
  for (let i = 1; i < 15; i++) offs[i + 1] = offs[i] + counts[i]
  for (let i = 0; i < n; i++) if (lengths[i]) syms[offs[lengths[i]]++] = i
  return { counts, syms }
}

let FIXED_L: Huff | null = null, FIXED_D: Huff | null = null
function fixedTables(): void {
  if (FIXED_L) return
  const l = new Uint8Array(288)
  for (let i = 0; i < 144; i++) l[i] = 8
  for (let i = 144; i < 256; i++) l[i] = 9
  for (let i = 256; i < 280; i++) l[i] = 7
  for (let i = 280; i < 288; i++) l[i] = 8
  FIXED_L = buildHuff(l, 288)
  FIXED_D = buildHuff(new Uint8Array(30).fill(5), 30)
}

export function inflateRawJS(data: Uint8Array, sizeHint?: number): Uint8Array {
  let pos = 0, bitbuf = 0, bitcnt = 0
  let out = new Uint8Array(Math.max(1024, sizeHint || data.length * 4)), op = 0
  const need = (n: number) => {
    if (op + n <= out.length) return
    let sz = out.length * 2
    while (sz < op + n) sz *= 2
    const nb = new Uint8Array(sz)
    nb.set(out.subarray(0, op))
    out = nb
  }
  const bits = (n: number): number => {
    while (bitcnt < n) {
      if (pos >= data.length) throw new Error('inflate: unexpected end of data')
      bitbuf |= data[pos++] << bitcnt
      bitcnt += 8
    }
    const v = bitbuf & ((1 << n) - 1)
    bitbuf >>>= n
    bitcnt -= n
    return v
  }
  const decode = (h: Huff): number => {
    let code = 0, first = 0, index = 0
    for (let len = 1; len < 16; len++) {
      code |= bits(1)
      const count = h.counts[len]
      if (code - count < first) return h.syms[index + (code - first)]
      index += count
      first += count
      first <<= 1
      code <<= 1
    }
    throw new Error('inflate: bad code')
  }
  const codes = (lt: Huff, dt: Huff) => {
    for (;;) {
      let sym = decode(lt)
      if (sym < 256) { need(1); out[op++] = sym }
      else if (sym === 256) return
      else {
        sym -= 257
        if (sym >= 29) throw new Error('inflate: bad length')
        const len = LBASE[sym] + bits(LEXT[sym])
        const ds = decode(dt)
        const dist = DBASE[ds] + bits(DEXT[ds])
        if (dist > op) throw new Error('inflate: distance too far')
        need(len)
        for (let i = 0; i < len; i++) { out[op] = out[op - dist]; op++ }
      }
    }
  }
  let last: number
  do {
    last = bits(1)
    const type = bits(2)
    if (type === 0) {
      bitbuf = 0
      bitcnt = 0
      const len = data[pos] | (data[pos + 1] << 8)
      pos += 4
      need(len)
      out.set(data.subarray(pos, pos + len), op)
      op += len
      pos += len
    } else if (type === 1) {
      fixedTables()
      codes(FIXED_L!, FIXED_D!)
    } else if (type === 2) {
      const nlen = bits(5) + 257, ndist = bits(5) + 1, ncode = bits(4) + 4
      const cl = new Uint8Array(19)
      for (let i = 0; i < ncode; i++) cl[CLORDER[i]] = bits(3)
      const clh = buildHuff(cl, 19)
      const lens = new Uint8Array(nlen + ndist)
      let i = 0
      while (i < nlen + ndist) {
        const sym = decode(clh)
        if (sym < 16) lens[i++] = sym
        else {
          let rep = 0, val = 0
          if (sym === 16) {
            if (!i) throw new Error('inflate: repeat with no prior')
            val = lens[i - 1]
            rep = 3 + bits(2)
          } else if (sym === 17) rep = 3 + bits(3)
          else rep = 11 + bits(7)
          while (rep--) lens[i++] = val
        }
      }
      codes(buildHuff(lens.subarray(0, nlen), nlen), buildHuff(lens.subarray(nlen), ndist))
    } else throw new Error('inflate: bad block type')
  } while (!last)
  return out.slice(0, op)
}

async function streamThrough(data: Uint8Array, stream: GenericTransformStream): Promise<Uint8Array> {
  const res = new Response(new Blob([data as BlobPart]).stream().pipeThrough(stream))
  return new Uint8Array(await res.arrayBuffer())
}

export async function inflateRaw(data: Uint8Array, size?: number): Promise<Uint8Array> {
  if (typeof DecompressionStream === 'function') {
    try { return await streamThrough(data, new DecompressionStream('deflate-raw')) } catch { /* fall back */ }
  }
  return inflateRawJS(data, size)
}

async function deflateRaw(data: Uint8Array): Promise<Uint8Array | null> {
  if (typeof CompressionStream === 'function') {
    try { return await streamThrough(data, new CompressionStream('deflate-raw')) } catch { /* store */ }
  }
  return null
}

/* ---------- reader ---------- */
export interface ZipEntry {
  name: string
  method: number
  csize: number
  usize: number
  loff: number
  crc?: number
  bytes(): Promise<Uint8Array>
  text(): Promise<string>
}

export interface ZipFiles extends Map<string, ZipEntry> {
  repaired?: boolean
}

function makeEntry(u8: Uint8Array, dv: DataView, name: string, method: number, csize: number, usize: number, loff: number, dataStart: number | undefined, crc?: number): ZipEntry {
  let cache: Uint8Array | null = null
  return {
    name, method, csize, usize, loff, crc,
    async bytes() {
      if (cache) return cache
      let start = dataStart
      if (start == null) {
        const lh = loff
        if (lh + 30 > u8.length || dv.getUint32(lh, true) !== 0x04034b50) throw new Error('Damaged entry ' + name)
        start = lh + 30 + dv.getUint16(lh + 26, true) + dv.getUint16(lh + 28, true)
      }
      const raw = u8.subarray(start, Math.min(u8.length, start + csize))
      let out: Uint8Array
      if (method === 0) out = raw.slice()
      else if (method === 8) out = await inflateRaw(raw, usize)
      else throw new Error('Unsupported compression method ' + method + ' in ' + name)
      cache = out
      return out
    },
    async text() { return td.decode(await this.bytes()) },
  }
}

const decodeName = (u8: Uint8Array, flags: number) =>
  flags & 0x800 ? td.decode(u8) : Array.from(u8, (b) => String.fromCharCode(b)).join('')

function addEntry(files: ZipFiles, entry: ZipEntry): void {
  if (entry.name.endsWith('/')) return
  files.set(entry.name, entry)
  const norm = entry.name.replace(/\\/g, '/')
  if (norm !== entry.name && !files.has(norm)) files.set(norm, entry)
}

function readCentral(u8: Uint8Array, dv: DataView): ZipFiles | null {
  let eocd = -1
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break }
  }
  if (eocd < 0) return null
  let count = dv.getUint16(eocd + 10, true)
  let cdOff = dv.getUint32(eocd + 16, true)
  if (cdOff === 0xffffffff || count === 0xffff) {
    const loc = eocd - 20
    if (loc >= 0 && dv.getUint32(loc, true) === 0x07064b50) {
      const z64 = Number(dv.getBigUint64(loc + 8, true))
      count = Number(dv.getBigUint64(z64 + 32, true))
      cdOff = Number(dv.getBigUint64(z64 + 48, true))
    }
  }
  const files: ZipFiles = new Map()
  let p = cdOff
  for (let n = 0; n < count; n++) {
    if (p + 46 > u8.length || dv.getUint32(p, true) !== 0x02014b50) break
    const flags = dv.getUint16(p + 8, true)
    const method = dv.getUint16(p + 10, true)
    const crc = dv.getUint32(p + 16, true)
    let csize = dv.getUint32(p + 20, true)
    let usize = dv.getUint32(p + 24, true)
    const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true)
    let loff = dv.getUint32(p + 42, true)
    const name = decodeName(u8.subarray(p + 46, p + 46 + nlen), flags)
    let xp = p + 46 + nlen
    const xend = Math.min(u8.length, xp + xlen)
    while (xp + 4 <= xend) {
      const id = dv.getUint16(xp, true), sz = dv.getUint16(xp + 2, true)
      if (id === 1) {
        let q = xp + 4
        if (usize === 0xffffffff) { usize = Number(dv.getBigUint64(q, true)); q += 8 }
        if (csize === 0xffffffff) { csize = Number(dv.getBigUint64(q, true)); q += 8 }
        if (loff === 0xffffffff) { loff = Number(dv.getBigUint64(q, true)) }
      }
      xp += 4 + sz
    }
    p += 46 + nlen + xlen + clen
    addEntry(files, makeEntry(u8, dv, name, method, csize, usize, loff, undefined, crc))
  }
  return files.size ? files : null
}

/** Repair path: walk the local file headers when the central directory is missing or damaged. */
function readLocal(u8: Uint8Array, dv: DataView): ZipFiles {
  const files: ZipFiles = new Map()
  let p = 0
  while (p + 30 <= u8.length) {
    if (dv.getUint32(p, true) !== 0x04034b50) {
      let q = p + 1
      while (q + 4 <= u8.length && !(u8[q] === 0x50 && u8[q + 1] === 0x4b && u8[q + 2] === 0x03 && u8[q + 3] === 0x04)) q++
      if (q + 30 > u8.length) break
      p = q
      continue
    }
    const flags = dv.getUint16(p + 6, true), method = dv.getUint16(p + 8, true)
    let csize = dv.getUint32(p + 18, true), usize = dv.getUint32(p + 22, true)
    const nlen = dv.getUint16(p + 26, true), xlen = dv.getUint16(p + 28, true)
    const name = decodeName(u8.subarray(p + 30, p + 30 + nlen), flags)
    const start = p + 30 + nlen + xlen
    let next = start + csize
    if ((flags & 8) || csize === 0xffffffff || (csize === 0 && method === 8)) {
      let q = start, found = false
      while (q + 16 <= u8.length) {
        if (u8[q] === 0x50 && u8[q + 1] === 0x4b) {
          const sig = dv.getUint32(q, true)
          if (sig === 0x08074b50 && dv.getUint32(q + 8, true) === q - start) { csize = q - start; usize = dv.getUint32(q + 12, true); next = q + 16; found = true; break }
          if (sig === 0x04034b50 || sig === 0x02014b50) { csize = q - start; usize = 0; next = q; found = true; break }
        }
        q++
      }
      if (!found) { csize = u8.length - start; usize = 0; next = u8.length }
    }
    if (start > u8.length) break
    addEntry(files, makeEntry(u8, dv, name, method, csize, usize, p, start))
    p = Math.max(next, p + 30)
  }
  return files
}

export interface ZipError extends Error {
  code?: string
  encrypted?: boolean
  bytes?: Uint8Array
}

export async function readZip(buffer: Uint8Array | ArrayBuffer): Promise<ZipFiles> {
  const u8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer)
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength)
  if (u8.length < 22) throw new Error('This file is not a valid Office Open XML (ZIP) package.')
  if (dv.getUint32(0, true) !== 0x04034b50 && dv.getUint32(0, true) === 0xe011cfd0) {
    const probe = new TextDecoder('utf-16le').decode(u8.subarray(0, Math.min(u8.length, 1 << 20)))
    const enc = /EncryptedPackage/.test(probe)
    const err: ZipError = new Error(enc
      ? 'This file is password-protected.'
      : 'This is a binary Office 97–2003 file, not an Office Open XML package.')
    err.code = 'ole'
    err.encrypted = enc
    err.bytes = u8
    throw err
  }
  let files: ZipFiles | null = null
  try { files = readCentral(u8, dv) } catch { files = null }
  if (!files) {
    try { files = readLocal(u8, dv) } catch { files = null }
    if (!files || !files.size) throw new Error(dv.getUint32(0, true) === 0x04034b50 ? 'This file is damaged and could not be repaired.' : 'This file is not a valid Office Open XML (ZIP) package.')
    files.repaired = true
  }
  return files
}

/* ---------- writer ---------- */
function dosDateTime(d: Date): { time: number; date: number } {
  const time = (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2)
  const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()
  return { time, date }
}

export interface ZipWriteFile {
  name: string
  data: Uint8Array | string
  store?: boolean
}

export async function writeZip(files: ZipWriteFile[], mime?: string): Promise<Blob> {
  const parts: Uint8Array[] = []
  const central: Uint8Array[] = []
  let offset = 0
  const { time, date } = dosDateTime(new Date())
  for (const f of files) {
    const nameBytes = te.encode(f.name)
    const data = typeof f.data === 'string' ? te.encode(f.data) : f.data
    const crc = crc32(data)
    let comp: Uint8Array | null = null
    if (!f.store && data.length > 64) comp = await deflateRaw(data)
    const method = comp && comp.length < data.length ? 8 : 0
    const body = method === 8 ? comp! : data
    const lh = new Uint8Array(30 + nameBytes.length)
    const ldv = new DataView(lh.buffer)
    ldv.setUint32(0, 0x04034b50, true)
    ldv.setUint16(4, 20, true)
    ldv.setUint16(6, 0x0800, true)
    ldv.setUint16(8, method, true)
    ldv.setUint16(10, time, true)
    ldv.setUint16(12, date, true)
    ldv.setUint32(14, crc, true)
    ldv.setUint32(18, body.length, true)
    ldv.setUint32(22, data.length, true)
    ldv.setUint16(26, nameBytes.length, true)
    ldv.setUint16(28, 0, true)
    lh.set(nameBytes, 30)
    parts.push(lh, body)
    const ch = new Uint8Array(46 + nameBytes.length)
    const cdv = new DataView(ch.buffer)
    cdv.setUint32(0, 0x02014b50, true)
    cdv.setUint16(4, 20, true)
    cdv.setUint16(6, 20, true)
    cdv.setUint16(8, 0x0800, true)
    cdv.setUint16(10, method, true)
    cdv.setUint16(12, time, true)
    cdv.setUint16(14, date, true)
    cdv.setUint32(16, crc, true)
    cdv.setUint32(20, body.length, true)
    cdv.setUint32(24, data.length, true)
    cdv.setUint16(28, nameBytes.length, true)
    cdv.setUint32(42, offset, true)
    ch.set(nameBytes, 46)
    central.push(ch)
    offset += lh.length + body.length
  }
  const cdSize = central.reduce((a, c) => a + c.length, 0)
  const end = new Uint8Array(22)
  const edv = new DataView(end.buffer)
  edv.setUint32(0, 0x06054b50, true)
  edv.setUint16(8, files.length, true)
  edv.setUint16(10, files.length, true)
  edv.setUint32(12, cdSize, true)
  edv.setUint32(16, offset, true)
  return new Blob([...parts, ...central, end] as BlobPart[], { type: mime || 'application/zip' })
}

export const encode = (s: string): Uint8Array => te.encode(s)
export const decode = (b: Uint8Array): string => td.decode(b)