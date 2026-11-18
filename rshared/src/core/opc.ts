/* Open Packaging Conventions carry, TypeScript implementation for the React rewrite.
 * Preservation model: every part of the source package is kept as raw bytes; parts the
 * editor did not change are written back byte-for-byte. Edited parts are replaced, and
 * [Content_Types].xml / .rels are regenerated only when parts or relationships changed.
 */
import { readZip, writeZip, type ZipFiles } from './zip'
import { parseXml, serializeXml, decodeBytes, esc, type El } from './xml'

export const NS = Object.freeze({
  rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  strictRel: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
  pkg: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  s: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
})

export interface FormatInfo {
  ext: string
  macro: boolean
  template: boolean
  family: string
  contentType: string
  mime: string
}

export const formats: Record<string, FormatInfo> = Object.freeze(Object.fromEntries([
  ['docx', 'wordprocessingml.document', 'ms-word.document', false, false],
  ['docm', 'wordprocessingml.document', 'ms-word.document', true, false],
  ['dotx', 'wordprocessingml.template', 'ms-word.template', false, true],
  ['dotm', 'wordprocessingml.template', 'ms-word.template', true, true],
  ['xlsx', 'spreadsheetml.sheet', 'ms-excel.sheet', false, false],
  ['xlsm', 'spreadsheetml.sheet', 'ms-excel.sheet', true, false],
  ['xltx', 'spreadsheetml.template', 'ms-excel.template', false, true],
  ['xltm', 'spreadsheetml.template', 'ms-excel.template', true, true],
  ['pptx', 'presentationml.presentation', 'ms-powerpoint.presentation', false, false],
  ['pptm', 'presentationml.presentation', 'ms-powerpoint.presentation', true, false],
  ['ppsx', 'presentationml.slideshow', 'ms-powerpoint.slideshow', false, false],
  ['ppsm', 'presentationml.slideshow', 'ms-powerpoint.slideshow', true, false],
  ['potx', 'presentationml.template', 'ms-powerpoint.template', false, true],
  ['potm', 'presentationml.template', 'ms-powerpoint.template', true, true],
].map(([ext, standard, enabled, macro, template]) => {
  const stem = macro ? 'application/vnd.' + enabled : 'application/vnd.openxmlformats-officedocument.' + standard
  const e = ext as string, m = macro as boolean
  return [e, Object.freeze({
    ext: e, macro: m, template: template as boolean, family: (standard as string).split('.')[0],
    contentType: stem + (m ? (e === 'dotm' ? '.macroEnabledTemplate.main+xml' : '.macroEnabled.main+xml') : '.main+xml'),
    mime: stem + (m ? '.macroEnabled.12' : ''),
  })]
})) as Record<string, FormatInfo>)

export const uuid = (): string => {
  if (crypto.randomUUID) return crypto.randomUUID()
  const b = crypto.getRandomValues(new Uint8Array(16))
  b[6] = (b[6] & 0x0f) | 0x40
  b[8] = (b[8] & 0x3f) | 0x80
  const x = Array.from(b, (n) => n.toString(16).padStart(2, '0')).join('')
  return `${x.slice(0, 8)}-${x.slice(8, 12)}-${x.slice(12, 16)}-${x.slice(16, 20)}-${x.slice(20)}`
}

export interface Rel { id: string; type: string; target: string; mode?: string }

const CT_PART = '[Content_Types].xml'
const relsPath = (part: string) => {
  const i = part.lastIndexOf('/')
  return (i < 0 ? '' : part.slice(0, i + 1)) + '_rels/' + (i < 0 ? part : part.slice(i + 1)) + '.rels'
}
/** resolve a relationship target against the part that owns it */
export function resolveTarget(part: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const base = part.includes('/') ? part.slice(0, part.lastIndexOf('/') + 1) : ''
  const segs = (base + target).split('/')
  const out: string[] = []
  for (const s of segs) {
    if (s === '.' || s === '') continue
    if (s === '..') out.pop()
    else out.push(s)
  }
  return out.join('/')
}
const relType = (t: string) => t.replace(NS.strictRel + '/', NS.rel + '/')

export class Package {
  /** raw bytes of every part, keyed by part name (no leading slash) */
  parts = new Map<string, Uint8Array>()
  /** parts replaced by the editor */
  edited = new Set<string>()
  private textCache = new Map<string, string>()
  private xmlCache = new Map<string, El>()
  private relsCache = new Map<string, Rel[]>()
  defaults = new Map<string, string>()
  overrides = new Map<string, string>()
  private ctDirty = false
  private relsDirty = new Set<string>()
  main = ''
  repaired = false

  static async open(bytes: Uint8Array | ArrayBuffer): Promise<Package> {
    const files: ZipFiles = await readZip(bytes)
    const pkg = new Package()
    pkg.repaired = !!files.repaired
    for (const [name, entry] of files) pkg.parts.set(name, await entry.bytes())
    const ct = pkg.parts.get(CT_PART)
    if (!ct) throw new Error('This package has no [Content_Types].xml part.')
    const ctRoot = parseXml(decodeBytes(ct))
    for (const el of ctRoot.children) {
      if (el.localName === 'Default') pkg.defaults.set(el.getAttribute('Extension') || '', el.getAttribute('ContentType') || '')
      else if (el.localName === 'Override') pkg.overrides.set((el.getAttribute('PartName') || '').replace(/^\//, ''), el.getAttribute('ContentType') || '')
    }
    const rootRels = pkg.relsOf('')
    const office = rootRels.find((r) => relType(r.type) === NS.rel + '/officeDocument')
    if (office) pkg.main = resolveTarget('', office.target)
    return pkg
  }

  has(part: string): boolean { return this.parts.has(part) }

  bytes(part: string): Uint8Array | undefined { return this.parts.get(part) }

  text(part: string): string {
    let t = this.textCache.get(part)
    if (t === undefined) {
      const b = this.parts.get(part)
      if (!b) throw new Error('No such part: ' + part)
      t = decodeBytes(b)
      this.textCache.set(part, t)
    }
    return t
  }

  xml(part: string): El {
    let x = this.xmlCache.get(part)
    if (!x) { x = parseXml(this.text(part)); this.xmlCache.set(part, x) }
    return x
  }

  type(part: string): string {
    const ov = this.overrides.get(part)
    if (ov) return ov
    const i = part.lastIndexOf('.')
    return (i < 0 ? '' : this.defaults.get(part.slice(i + 1))) || ''
  }

  /** relationships of a part ('' = the package root) */
  relsOf(part: string): Rel[] {
    const key = part || ''
    let rels = this.relsCache.get(key)
    if (rels) return rels
    rels = []
    const b = this.parts.get(relsPath(part))
    if (b) {
      const root = parseXml(decodeBytes(b))
      for (const el of root.children) {
        if (el.localName !== 'Relationship') continue
        rels.push({
          id: el.getAttribute('Id') || '',
          type: relType(el.getAttribute('Type') || ''),
          target: el.getAttribute('Target') || '',
          mode: el.getAttribute('TargetMode') || undefined,
        })
      }
    }
    this.relsCache.set(key, rels)
    return rels
  }

  rel(part: string, id: string): Rel | undefined { return this.relsOf(part).find((r) => r.id === id) }

  /** parts related to `part` by relationship type (resolved names) */
  related(part: string, type: string): string[] {
    return this.relsOf(part).filter((r) => r.type === type && r.mode !== 'External').map((r) => resolveTarget(part, r.target))
  }

  setPart(part: string, data: string | Uint8Array): void {
    this.parts.set(part, typeof data === 'string' ? new TextEncoder().encode(data) : data)
    this.edited.add(part)
    this.textCache.delete(part)
    this.xmlCache.delete(part)
  }

  addPart(part: string, data: string | Uint8Array, contentType?: string): void {
    this.setPart(part, data)
    if (contentType) { this.overrides.set(part, contentType); this.ctDirty = true }
    else {
      const ext = part.includes('.') ? part.slice(part.lastIndexOf('.') + 1) : ''
      if (ext && !this.defaults.has(ext)) {
        this.defaults.set(ext, contentTypeForExt(ext))
        this.ctDirty = true
      }
    }
  }

  removePart(part: string): void {
    this.parts.delete(part)
    this.edited.delete(part)
    this.textCache.delete(part)
    this.xmlCache.delete(part)
    if (this.overrides.delete(part)) this.ctDirty = true
    const rp = relsPath(part)
    if (this.parts.delete(rp)) this.relsDirty.delete(rp)
  }

  /** add a relationship; returns its new Id (rIdN, as Office numbers them) */
  addRel(part: string, type: string, target: string, mode?: string): string {
    const rels = this.relsOf(part).slice()
    let max = 0
    for (const r of rels) { const m = /^rId(\d+)$/.exec(r.id); if (m) max = Math.max(max, +m[1]) }
    const id = 'rId' + (max + 1)
    rels.push({ id, type, target, mode })
    const key = part || ''
    this.relsCache.set(key, rels)
    this.relsDirty.add(key)
    return id
  }

  /** next free media part name, as Office numbers them (image1.png, image2.png, ...) */
  nextMediaName(ext: string, stem = 'image'): string {
    let n = 1
    while (this.parts.has(`word/media/${stem}${n}.${ext}`) || this.parts.has(`ppt/media/${stem}${n}.${ext}`) || this.parts.has(`xl/media/${stem}${n}.${ext}`)) n++
    return `${stem}${n}.${ext}`
  }

  variant(fallback: string): string {
    const t = this.type(this.main)
    return Object.values(formats).find((f) => f.contentType === t)?.ext || fallback
  }

  private writeContentTypes(): void {
    if (!this.ctDirty) return
    const lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="' + NS.ct + '">']
    for (const [ext, type] of [...this.defaults].sort()) lines.push(`<Default Extension="${esc(ext)}" ContentType="${esc(type)}"/>`)
    for (const [part, type] of [...this.overrides].sort()) lines.push(`<Override PartName="/${esc(part)}" ContentType="${esc(type)}"/>`)
    lines.push('</Types>')
    this.parts.set(CT_PART, new TextEncoder().encode(lines.join('')))
  }

  private writeRels(): void {
    for (const key of this.relsDirty) {
      const rels = this.relsCache.get(key) || []
      const part = key
      const lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="' + NS.pkg + '">']
      for (const r of rels) {
        lines.push(`<Relationship Id="${esc(r.id)}" Type="${esc(r.type)}" Target="${esc(r.target)}"${r.mode ? ` TargetMode="${esc(r.mode)}"` : ''}/>`)
      }
      lines.push('</Relationships>')
      const rp = relsPath(part)
      this.parts.set(rp, new TextEncoder().encode(lines.join('')))
      if (!this.overrides.has(rp)) { this.defaults.set('rels', 'application/vnd.openxmlformats-package.relationships+xml'); this.ctDirty = true }
    }
  }

  async save(mime?: string): Promise<Blob> {
    this.writeRels()
    this.writeContentTypes()
    const files = [...this.parts].map(([name, data]) => ({ name, data }))
    return writeZip(files, mime || 'application/zip')
  }
}

export function contentTypeForExt(ext: string): string {
  const map: Record<string, string> = {
    xml: 'application/xml',
    rels: 'application/vnd.openxmlformats-package.relationships+xml',
    png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp',
    tiff: 'image/tiff', svg: 'image/svg+xml', emf: 'image/x-emf', wmf: 'image/x-wmf',
    bin: 'application/vnd.openxmlformats-officedocument.oleObject',
  }
  return map[ext.toLowerCase()] || 'application/octet-stream'
}

/** a fresh package with a main part and the officeDocument relationship */
export function newPackage(mainPart: string, mainContent: string, format: FormatInfo, extraDefaults: Record<string, string> = {}): Package {
  const pkg = new Package()
  pkg.main = mainPart
  pkg.defaults.set('rels', 'application/vnd.openxmlformats-package.relationships+xml')
  pkg.defaults.set('xml', 'application/xml')
  for (const [ext, type] of Object.entries(extraDefaults)) pkg.defaults.set(ext, type)
  pkg.overrides.set(mainPart, format.contentType)
  pkg.setPart(mainPart, mainContent)
  pkg.addRel('', NS.rel + '/officeDocument', mainPart)
  return pkg
}

/** serialize a parsed tree into a full XML part with declaration */
export function xmlPart(root: El): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + serializeXml(root)
}
