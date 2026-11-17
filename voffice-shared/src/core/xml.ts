/* Portable XML tree (browser and Node), TypeScript rewrite of public/common/xml.js. */

const ENT: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }

export function unescape(s: string): string {
  if (s.indexOf('&') < 0) return s
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e: string) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)
      if (!(cp >= 0 && cp <= 0x10ffff)) return m
      try { return String.fromCodePoint(cp) } catch { return m }
    }
    return ENT[e] != null ? ENT[e] : m
  })
}

export const esc = (s: string): string => {
  s = String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  return /[\t\n\r]/.test(s) ? s.replace(/\t/g, '&#9;').replace(/\n/g, '&#10;').replace(/\r/g, '&#13;') : s
}

const BAD_CR = new RegExp('[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F\\uFFFE\\uFFFF]', 'g')
const BAD = new RegExp('[\\x00-\\x08\\x0B-\\x1F\\uFFFE\\uFFFF]', 'g')

export const escText = (s: string, keepCR?: boolean): string =>
  String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/_x([0-9A-Fa-f]{4})_/g, '_x005F_x$1_')
    .replace(keepCR ? BAD_CR : BAD, (c) => '_x' + c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0') + '_')

export const unx = (s: string): string =>
  s.indexOf('_x') < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)))

export interface XmlAttr { name: string; prefix: string | null; localName: string; value: string }

export class Txt {
  data: string
  parentNode: El | null
  constructor(data: string, parent: El | null) { this.data = data; this.parentNode = parent }
  get nodeType(): number { return 3 }
  get textContent(): string { return this.data }
  get nodeValue(): string { return this.data }
  get localName(): null { return null }
  remove(): void {
    const p = this.parentNode
    if (!p) return
    const i = p.childNodes.indexOf(this)
    if (i >= 0) p.childNodes.splice(i, 1)
    this.parentNode = null
  }
}

export type XmlNode = El | Txt

export class El {
  nodeName: string
  prefix: string | null
  localName: string
  attrs: Record<string, string> | null = null
  children: El[] = []
  childNodes: XmlNode[] = []
  parentNode: El | null
  constructor(name: string, parent: El | null) {
    this.nodeName = name
    const i = name.indexOf(':')
    this.prefix = i < 0 ? null : name.slice(0, i)
    this.localName = i < 0 ? name : name.slice(i + 1)
    this.parentNode = parent
  }
  get nodeType(): number { return 1 }
  get tagName(): string { return this.nodeName }
  getAttribute(q: string): string | null { const a = this.attrs; if (!a) return null; const v = a[q]; return v === undefined ? null : v }
  hasAttribute(q: string): boolean { return this.getAttribute(q) != null }
  setAttribute(q: string, v: string | number): void { const a = this.attrs ?? (this.attrs = Object.create(null)); a[q] = String(v) }
  removeAttribute(q: string): void { if (this.attrs) delete this.attrs[q] }
  get attributes(): XmlAttr[] {
    const out: XmlAttr[] = []
    if (this.attrs) for (const k in this.attrs) {
      const i = k.indexOf(':')
      out.push({ name: k, prefix: i < 0 ? null : k.slice(0, i), localName: i < 0 ? k : k.slice(i + 1), value: this.attrs[k] })
    }
    return out
  }
  lookupNamespaceURI(prefix: string | null): string | null {
    const key = prefix ? 'xmlns:' + prefix : 'xmlns'
    for (let e: El | null = this; e && e.nodeType === 1; e = e.parentNode) {
      const a = e.attrs
      if (a && a[key] !== undefined) return a[key]
    }
    if (prefix === 'xml') return 'http://www.w3.org/XML/1998/namespace'
    return null
  }
  get namespaceURI(): string | null { return this.lookupNamespaceURI(this.prefix) }
  getAttributeNS(ns: string | null, local: string): string | null {
    const a = this.attrs
    if (!a) return null
    if (!ns) { const v = a[local]; return v === undefined ? null : v }
    for (const k in a) {
      const i = k.indexOf(':')
      if (i < 0 || k.slice(i + 1) !== local) continue
      const p = k.slice(0, i)
      if (p === 'xmlns') continue
      if (this.lookupNamespaceURI(p) === ns) return a[k]
    }
    return null
  }
  get textContent(): string {
    if (this.childNodes.length === 1 && this.childNodes[0].nodeType === 3) return (this.childNodes[0] as Txt).data
    let s = ''
    const walk = (e: El) => { for (const n of e.childNodes) s += n.nodeType === 3 ? (n as Txt).data : (walk(n as El), '') }
    walk(this)
    return s
  }
  set textContent(v: string) { this.children = []; this.childNodes = [new Txt(String(v), this)] }
  get firstElementChild(): El | null { return this.children[0] || null }
  get firstChild(): XmlNode | null { return this.childNodes[0] || null }
  get parentElement(): El | null { return this.parentNode && this.parentNode.nodeType === 1 ? this.parentNode : null }
  getElementsByTagName(name: string): El[] {
    const out: El[] = []
    const any = name === '*'
    const walk = (e: El) => { for (const c of e.children) { if (any || c.nodeName === name) out.push(c); if (c.children.length) walk(c) } }
    walk(this)
    return out
  }
  getElementsByTagNameNS(ns: string, local: string): El[] {
    const out: El[] = []
    const walk = (e: El) => {
      for (const c of e.children) {
        if ((local === '*' || c.localName === local) && (ns === '*' || c.namespaceURI === ns)) out.push(c)
        if (c.children.length) walk(c)
      }
    }
    walk(this)
    return out
  }
  appendChild(n: XmlNode): XmlNode {
    if (n.parentNode) n.remove()
    n.parentNode = this
    this.childNodes.push(n)
    if (n.nodeType === 1) this.children.push(n as El)
    return n
  }
  remove(): void {
    const p = this.parentNode
    if (!p) return
    let i = p.childNodes.indexOf(this)
    if (i >= 0) p.childNodes.splice(i, 1)
    i = p.children.indexOf(this)
    if (i >= 0) p.children.splice(i, 1)
    this.parentNode = null
  }
  replaceWith(...nodes: XmlNode[]): void {
    const p = this.parentNode
    if (!p) return
    for (const n of nodes) if (n.parentNode) n.remove()
    const i = p.childNodes.indexOf(this)
    p.childNodes.splice(i, 1, ...nodes)
    for (const n of nodes) n.parentNode = p
    p.children = p.childNodes.filter((n): n is El => n.nodeType === 1)
    this.parentNode = null
  }
  cloneNode(deep?: boolean): El {
    const e = new El(this.nodeName, null)
    if (this.attrs) e.attrs = { ...this.attrs }
    if (deep) for (const n of this.childNodes) e.appendChild(n.nodeType === 3 ? new Txt((n as Txt).data, e) : (n as El).cloneNode(true))
    return e
  }
}

/* ---------- parser ---------- */
const reAttr = /([^\s=\/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

function parseAttrs(s: string, el: El): void {
  reAttr.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = reAttr.exec(s))) {
    const a = el.attrs ?? (el.attrs = Object.create(null))
    const v = m[2] !== undefined ? m[2] : m[3]
    a[m[1]] = v.indexOf('&') < 0 ? v : unescape(v)
  }
}

export function attrs(s: string): Record<string, string> {
  const o: Record<string, string> = Object.create(null)
  reAttr.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = reAttr.exec(s))) { const v = m[2] !== undefined ? m[2] : m[3]; o[m[1]] = v.indexOf('&') < 0 ? v : unescape(v) }
  return o
}

export interface ParseOpts { keepWhitespace?: boolean; strict?: boolean }

const WS_KEEP = new Set(['t', 'v', 'f', 'formula', 'formula1', 'formula2', 'oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter', 'text', 'Data', 'Font', 'B', 'I', 'U', 'S', 'Sup', 'Sub'])

/** parse a document; returns the document element */
export function parseXml(text: string, opts?: ParseOpts): El {
  if (text && text.charCodeAt(0) === 0xfeff) text = text.slice(1)
  const keepWs = !!opts?.keepWhitespace
  const strict = !!opts?.strict
  const doc = new El('#document', null)
  let cur: El = doc
  let i = 0
  const n = text.length
  const addText = (s: string) => {
    if (cur === doc) return
    if (!keepWs && /^\s*$/.test(s) && !WS_KEEP.has(cur.localName)) return
    cur.childNodes.push(new Txt(unescape(s), cur))
  }
  while (i < n) {
    const lt = text.indexOf('<', i)
    if (lt < 0) { addText(text.slice(i)); break }
    if (lt > i) addText(text.slice(i, lt))
    const c = text.charCodeAt(lt + 1)
    if (c === 47 /* / */) {
      const gt = text.indexOf('>', lt)
      if (gt < 0) throw new Error('XML parse error: unterminated end tag')
      const name = text.slice(lt + 2, gt).trim()
      let e: El | null = cur
      if (strict && (cur === doc || cur.nodeName !== name)) throw new Error('XML parse error: mismatched end tag ' + name)
      while (e && e !== doc && e.nodeName !== name) e = e.parentNode
      if (e && e !== doc) cur = e.parentNode!
      i = gt + 1
    } else if (c === 33 /* ! */) {
      if (text.startsWith('<!--', lt)) { const e = text.indexOf('-->', lt + 4); i = e < 0 ? n : e + 3 }
      else if (text.startsWith('<![CDATA[', lt)) {
        const e = text.indexOf(']]>', lt + 9)
        const data = text.slice(lt + 9, e < 0 ? n : e)
        if (cur !== doc) cur.childNodes.push(new Txt(data, cur))
        i = e < 0 ? n : e + 3
      } else {
        let depth = 0, j = lt + 2
        for (; j < n; j++) { const ch = text[j]; if (ch === '[') depth++; else if (ch === ']') depth--; else if (ch === '>' && depth <= 0) break }
        i = j + 1
      }
    } else if (c === 63 /* ? */) {
      const e = text.indexOf('?>', lt + 2)
      i = e < 0 ? n : e + 2
    } else {
      let j = lt + 1, q = 0
      for (; j < n; j++) {
        const ch = text.charCodeAt(j)
        if (q) { if (ch === q) q = 0 }
        else if (ch === 34 || ch === 39) q = ch
        else if (ch === 62) break
      }
      if (j >= n) throw new Error('XML parse error: unterminated start tag')
      let k = lt + 1
      while (k < j) { const ch = text.charCodeAt(k); if (ch === 32 || ch === 9 || ch === 10 || ch === 13 || ch === 47 || ch === 62) break; k++ }
      const name = text.slice(lt + 1, k)
      const self = text.charCodeAt(j - 1) === 47
      const el = new El(name, cur)
      const attrEnd = self ? j - 1 : j
      if (k < attrEnd) parseAttrs(text.slice(k, attrEnd), el)
      cur.childNodes.push(el)
      cur.children.push(el)
      if (!self) cur = el
      i = j + 1
    }
  }
  const de = doc.children[0]
  if (!de) throw new Error('XML parse error: no root element')
  if (strict && (cur !== doc || doc.children.length !== 1)) throw new Error('XML parse error: incomplete document')
  return de
}

/** character data between two tags as text: entities resolved, CDATA taken literally, line ends normalised */
export function chars(s: string): string {
  if (s.indexOf('\r') >= 0) s = s.replace(/\r\n?/g, '\n')
  return s.indexOf('<![CDATA[') < 0 ? unescape(s)
    : s.split(/(<!\[CDATA\[[\s\S]*?\]\]>)/).map((p) => (p.startsWith('<![CDATA[') ? p.slice(9, -3) : unescape(p))).join('')
}

/** Scan tags between [from, to): fn(local, kind, attrText, textBefore); kind 1 open, 2 close, 3 self-closing. */
export function scanXml(text: string, from: number, to: number, fn: (local: string, kind: number, attrText: string, textBefore: string) => void): void {
  const re = /<(\/?)(?:[A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)([^>]*?)(\/?)>/g
  re.lastIndex = from
  let last = from, m: RegExpExecArray | null
  while ((m = re.exec(text)) && m.index < to) {
    const txt = m.index > last ? text.slice(last, m.index) : ''
    fn(m[2], m[1] ? 2 : m[4] ? 3 : 1, m[3], txt)
    last = re.lastIndex
  }
}

/** bytes → string honouring a UTF-16 byte-order mark */
export function decodeBytes(u8: Uint8Array): string {
  if (u8.length >= 2 && u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder('utf-16le').decode(u8.subarray(2))
  if (u8.length >= 2 && u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder('utf-16be').decode(u8.subarray(2))
  if (u8.length >= 4 && u8[0] === 0 && u8[1] === 0x3c && u8[2] === 0 && u8[3] === 0x3f) return new TextDecoder('utf-16be').decode(u8)
  if (u8.length >= 4 && u8[0] === 0x3c && u8[1] === 0 && u8[2] === 0x3f && u8[3] === 0) return new TextDecoder('utf-16le').decode(u8)
  return new TextDecoder('utf-8').decode(u8)
}

/** serialize an element tree back to XML, carrying the in-scope namespace bindings */
export function serializeXml(el: El | null): string {
  if (!el) return ''
  const bindings: Record<string, string> = {}
  for (let e: El | null = el; e && e.nodeType === 1; e = e.parentNode) {
    if (e.attrs) for (const k in e.attrs) if (k === 'xmlns' || k.startsWith('xmlns:')) bindings[k] = e.attrs[k]
  }
  const write = (e: XmlNode, first: boolean): string => {
    if (e.nodeType === 3) return (e as Txt).data.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    const el2 = e as El
    const values: Record<string, string> = el2.attrs ? { ...el2.attrs } : {}
    if (first) for (const [key, value] of Object.entries(bindings)) if (!(key in values)) values[key] = value
    const tag = '<' + el2.nodeName + Object.entries(values).map(([k, v]) => ' ' + k + '="' + esc(v) + '"').join('')
    return el2.childNodes.length
      ? tag + '>' + el2.childNodes.map((c) => write(c, false)).join('') + '</' + el2.nodeName + '>'
      : tag + '/>'
  }
  return write(el, true)
}

/** create a detached element with a text child */
export function mkEl(name: string, attrs?: Record<string, string | number>, text?: string): El {
  const e = new El(name, null)
  if (attrs) for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v)
  if (text != null) e.textContent = text
  return e
}
