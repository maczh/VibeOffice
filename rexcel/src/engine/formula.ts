/* Formula engine: tokenizer → Pratt parser → evaluator, with a function library and
 * dependency-ordered recalculation. Values: number | string | boolean | Err | array (range). */
import { parseAddr, colIndex } from './model'

export class Err {
  constructor(public code: string) {}
  toString(): string { return this.code }
}
export const DIV0 = () => new Err('#DIV/0!')
export const VALUE = () => new Err('#VALUE!')
export const REF = () => new Err('#REF!')
export const NAME = () => new Err('#NAME?')
export const NA = () => new Err('#N/A')
export const NUM = () => new Err('#NUM!')

export type Val = number | string | boolean | Err | Val[]
export interface RefNode { r: number; c: number; sheet?: string }
export interface RangeNode { r1: number; c1: number; r2: number; c2: number; sheet?: string }

export type Node =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  | { k: 'err'; v: string }
  | { k: 'ref'; ref: RefNode }
  | { k: 'range'; range: RangeNode }
  | { k: 'un'; op: string; a: Node }
  | { k: 'bin'; op: string; a: Node; b: Node }
  | { k: 'pct'; a: Node }
  | { k: 'call'; name: string; args: Node[] }

/* ---------- tokenizer ---------- */
interface Tok { t: string; v: string }
function tokenize(src: string): Tok[] {
  const out: Tok[] = []
  let i = 0
  const n = src.length
  while (i < n) {
    const ch = src[i]
    if (ch === ' ' || ch === '\t' || ch === '\n') { i++; continue }
    if (ch === '"') {
      let s = ''; i++
      while (i < n) { if (src[i] === '"') { if (src[i + 1] === '"') { s += '"'; i += 2 } else { i++; break } } else s += src[i++] }
      out.push({ t: 'str', v: s }); continue
    }
    if (/[0-9.]/.test(ch)) {
      let s = ''
      while (i < n && /[0-9.]/.test(src[i])) s += src[i++]
      if (src[i] === 'e' || src[i] === 'E') { s += src[i++]; if (src[i] === '+' || src[i] === '-') s += src[i++]; while (i < n && /[0-9]/.test(src[i])) s += src[i++] }
      out.push({ t: 'num', v: s }); continue
    }
    if (/[A-Za-z_$]/.test(ch)) {
      let s = ''
      while (i < n && /[A-Za-z0-9_.$!]/.test(src[i])) s += src[i++]
      // sheet-qualified or range
      out.push({ t: 'ident', v: s }); continue
    }
    if (ch === '<' && src[i + 1] === '>') { out.push({ t: 'op', v: '<>' }); i += 2; continue }
    if (ch === '<' && src[i + 1] === '=') { out.push({ t: 'op', v: '<=' }); i += 2; continue }
    if (ch === '>' && src[i + 1] === '=') { out.push({ t: 'op', v: '>=' }); i += 2; continue }
    if ('+-*/^&=<>%,():'.includes(ch)) { out.push({ t: ch === '(' ? 'lp' : ch === ')' ? 'rp' : ch === ',' ? 'comma' : ch === ':' ? 'colon' : 'op', v: ch }); i++; continue }
    throw VALUE()
  }
  return out
}

function parseIdent(v: string): Node | null {
  const up = v.toUpperCase()
  if (up === 'TRUE') return { k: 'bool', v: true }
  if (up === 'FALSE') return { k: 'bool', v: false }
  if (/^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!)$/.test(up)) return { k: 'err', v: up }
  // sheet prefix
  let sheet: string | undefined
  let body = v
  const bang = v.lastIndexOf('!')
  if (bang >= 0) { sheet = v.slice(0, bang).replace(/^'(.*)'$/, '$1'); body = v.slice(bang + 1) }
  const clean = body.replace(/\$/g, '')
  const m = /^([A-Za-z]+)(\d+)$/.exec(clean)
  if (m) return { k: 'ref', ref: { r: +m[2] - 1, c: colIndex(m[1]), sheet } }
  return null
}

/* ---------- Pratt parser ---------- */
const PREC: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 }

export function parseFormula(src: string): Node {
  const toks = tokenize(src)
  let p = 0
  const peek = () => toks[p]
  const next = () => toks[p++]
  function parseExpr(minPrec: number): Node {
    let left = parseUnary()
    for (;;) {
      const t = peek()
      if (!t || t.t !== 'op' || !(t.v in PREC) || PREC[t.v] < minPrec) break
      const op = next().v
      const right = parseExpr(PREC[op] + (op === '^' ? 0 : 1))
      left = { k: 'bin', op, a: left, b: right }
    }
    return left
  }
  function parseUnary(): Node {
    const t = peek()
    if (t && t.t === 'op' && (t.v === '-' || t.v === '+')) { next(); return { k: 'un', op: t.v, a: parseUnary() } }
    return parsePostfix()
  }
  function parsePostfix(): Node {
    let a = parseAtom()
    while (peek() && peek().t === 'op' && peek().v === '%') { next(); a = { k: 'pct', a } }
    return a
  }
  function parseAtom(): Node {
    const t = next()
    if (!t) throw VALUE()
    if (t.t === 'num') return { k: 'num', v: parseFloat(t.v) }
    if (t.t === 'str') return { k: 'str', v: t.v }
    if (t.t === 'lp') { const e = parseExpr(1); if (next()?.t !== 'rp') throw VALUE(); return e }
    if (t.t === 'ident') {
      // function call?
      if (peek() && peek().t === 'lp') {
        next()
        const args: Node[] = []
        if (peek() && peek().t !== 'rp') {
          for (;;) { args.push(parseExpr(1)); if (peek()?.t === 'comma') { next(); continue } break }
        }
        if (next()?.t !== 'rp') throw VALUE()
        return { k: 'call', name: t.v.toUpperCase(), args }
      }
      // range A1:B2 ?
      if (peek() && peek().t === 'colon') {
        next()
        const b = next()
        const a = parseIdent(t.v); const bn = b && b.t === 'ident' ? parseIdent(b.v) : null
        if (a && a.k === 'ref' && bn && bn.k === 'ref') {
          return { k: 'range', range: { r1: Math.min(a.ref.r, bn.ref.r), c1: Math.min(a.ref.c, bn.ref.c), r2: Math.max(a.ref.r, bn.ref.r), c2: Math.max(a.ref.c, bn.ref.c), sheet: a.ref.sheet } }
        }
        throw REF()
      }
      const id = parseIdent(t.v)
      if (id) return id
      throw NAME()
    }
    throw VALUE()
  }
  const e = parseExpr(1)
  if (p < toks.length) throw VALUE()
  return e
}

/* ---------- coercion ---------- */
export function toNum(v: Val): number | Err {
  if (v instanceof Err) return v
  if (Array.isArray(v)) return v.length ? toNum(v[0]) : VALUE()
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  if (v === '' ) return 0
  const n = Number(v)
  return isNaN(n) ? VALUE() : n
}
export function toStr(v: Val): string | Err {
  if (v instanceof Err) return v
  if (Array.isArray(v)) return v.length ? toStr(v[0]) : VALUE()
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return String(v)
}
export function toBool(v: Val): boolean | Err {
  if (v instanceof Err) return v
  if (Array.isArray(v)) return v.length ? toBool(v[0]) : VALUE()
  if (typeof v === 'boolean') return v
  if (typeof v === 'number') return v !== 0
  if (v === 'TRUE') return true
  if (v === 'FALSE') return false
  const n = Number(v)
  return isNaN(n) ? VALUE() : n !== 0
}
const isErr = (v: Val): v is Err => v instanceof Err
function flat(v: Val): Val[] { return Array.isArray(v) ? v.flatMap(flat) : [v] }
function nums(v: Val): number[] {
  const out: number[] = []
  for (const x of flat(v)) {
    if (typeof x === 'number') out.push(x)
    else if (typeof x === 'string' && x !== '' && !isNaN(Number(x))) out.push(Number(x))
    else if (typeof x === 'boolean') out.push(x ? 1 : 0)
  }
  return out
}

/* ---------- evaluator ---------- */
export interface EvalCtx {
  get(ref: RefNode): Val
  range(range: RangeNode): Val[]
  current?: { r: number; c: number }
}
export type RangeVal = Val[] & { __dims?: [number, number] }

export function evalNode(n: Node, ctx: EvalCtx): Val {
  switch (n.k) {
    case 'num': return n.v
    case 'str': return n.v
    case 'bool': return n.v
    case 'err': return new Err(n.v)
    case 'ref': return ctx.get(n.ref)
    case 'range': return ctx.range(n.range)
    case 'pct': { const a = toNum(evalNode(n.a, ctx)); return isErr(a) ? a : a / 100 }
    case 'un': {
      const a = toNum(evalNode(n.a, ctx)); if (isErr(a)) return a
      return n.op === '-' ? -a : a
    }
    case 'bin': {
      const a = evalNode(n.a, ctx); const b = evalNode(n.b, ctx)
      if (isErr(a)) return a
      if (isErr(b)) return b
      if (n.op === '&') { const x = toStr(a); const y = toStr(b); return isErr(x) ? x : isErr(y) ? y : x + y }
      if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) return compare(a, b, n.op)
      const x = toNum(a); const y = toNum(b)
      if (isErr(x)) return x
      if (isErr(y)) return y
      switch (n.op) {
        case '+': return x + y
        case '-': return x - y
        case '*': return x * y
        case '/': return y === 0 ? DIV0() : x / y
        case '^': { const r = Math.pow(x, y); return isNaN(r) ? NUM() : r }
      }
      return VALUE()
    }
    case 'call': return callFn(n.name, n.args.map((a) => evalNode(a, ctx)), ctx)
  }
}

function compare(a: Val, b: Val, op: string): boolean | Err {
  // type order: number < string < boolean (Excel); errors propagate earlier
  const an = typeof a === 'number', bn = typeof b === 'number'
  const as = typeof a === 'string', bs = typeof b === 'string'
  const ab = typeof a === 'boolean', bb = typeof b === 'boolean'
  let cmp: number
  if (an && bn) cmp = (a as number) - (b as number)
  else if (as && bs) cmp = (a as string).localeCompare(b as string)
  else if (ab && bb) cmp = (a === b ? 0 : a ? 1 : -1)
  else {
    const rank = (v: Val) => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2
    cmp = rank(a) - rank(b)
  }
  switch (op) {
    case '=': return cmp === 0
    case '<>': return cmp !== 0
    case '<': return cmp < 0
    case '>': return cmp > 0
    case '<=': return cmp <= 0
    case '>=': return cmp >= 0
  }
  return VALUE()
}

/* ---------- function library ---------- */
type Fn = (args: Val[], ctx: EvalCtx) => Val
const errFirst = (args: Val[]): Err | null => { for (const a of flat(args)) if (isErr(a)) return a; return null }

const FNS: Record<string, Fn> = {
  SUM: (a) => { const e = errFirst(a); if (e) return e; return nums(a).reduce((x, y) => x + y, 0) },
  AVERAGE: (a) => { const e = errFirst(a); if (e) return e; const n = nums(a); return n.length ? n.reduce((x, y) => x + y, 0) / n.length : DIV0() },
  MIN: (a) => { const n = nums(a); return n.length ? Math.min(...n) : 0 },
  MAX: (a) => { const n = nums(a); return n.length ? Math.max(...n) : 0 },
  COUNT: (a) => nums(a).length,
  COUNTA: (a) => flat(a).filter((x) => x !== '' && x != null).length,
  COUNTBLANK: (a) => flat(a).filter((x) => x === '' || x == null).length,
  PRODUCT: (a) => { const n = nums(a); return n.length ? n.reduce((x, y) => x * y, 1) : 0 },
  SUMPRODUCT: (a) => {
    const arrays = a.map((x) => flat(x))
    const len = Math.max(...arrays.map((x) => x.length), 0)
    let sum = 0
    for (let i = 0; i < len; i++) {
      let prod = 1
      for (const arr of arrays) { const v = toNum(arr[i] ?? 0); if (isErr(v)) return v; prod *= v }
      sum += prod
    }
    return sum
  },
  ROUND: (a) => { const x = toNum(a[0]); const d = toNum(a[1] ?? 0); if (isErr(x) || isErr(d)) return VALUE(); const p = Math.pow(10, d); return Math.round(x * p) / p },
  ROUNDUP: (a) => { const x = toNum(a[0]); const d = toNum(a[1] ?? 0); if (isErr(x) || isErr(d)) return VALUE(); const p = Math.pow(10, d); return (x < 0 ? -1 : 1) * Math.ceil(Math.abs(x) * p) / p },
  ROUNDDOWN: (a) => { const x = toNum(a[0]); const d = toNum(a[1] ?? 0); if (isErr(x) || isErr(d)) return VALUE(); const p = Math.pow(10, d); return (x < 0 ? -1 : 1) * Math.floor(Math.abs(x) * p) / p },
  INT: (a) => { const x = toNum(a[0]); return isErr(x) ? x : Math.floor(x) },
  MOD: (a) => { const x = toNum(a[0]); const y = toNum(a[1]); if (isErr(x) || isErr(y)) return VALUE(); return y === 0 ? DIV0() : x - y * Math.floor(x / y) },
  ABS: (a) => { const x = toNum(a[0]); return isErr(x) ? x : Math.abs(x) },
  SIGN: (a) => { const x = toNum(a[0]); return isErr(x) ? x : Math.sign(x) },
  POWER: (a) => { const x = toNum(a[0]); const y = toNum(a[1]); if (isErr(x) || isErr(y)) return VALUE(); const r = Math.pow(x, y); return isNaN(r) ? NUM() : r },
  SQRT: (a) => { const x = toNum(a[0]); if (isErr(x)) return x; return x < 0 ? NUM() : Math.sqrt(x) },
  EXP: (a) => { const x = toNum(a[0]); return isErr(x) ? x : Math.exp(x) },
  LN: (a) => { const x = toNum(a[0]); if (isErr(x)) return x; return x <= 0 ? NUM() : Math.log(x) },
  LOG: (a) => { const x = toNum(a[0]); const b = toNum(a[1] ?? 10); if (isErr(x) || isErr(b)) return VALUE(); return x <= 0 ? NUM() : Math.log(x) / Math.log(b) },
  LOG10: (a) => { const x = toNum(a[0]); if (isErr(x)) return x; return x <= 0 ? NUM() : Math.log10(x) },
  SIN: (a) => { const x = toNum(a[0]); return isErr(x) ? x : Math.sin(x) },
  COS: (a) => { const x = toNum(a[0]); return isErr(x) ? x : Math.cos(x) },
  TAN: (a) => { const x = toNum(a[0]); return isErr(x) ? x : Math.tan(x) },
  PI: () => Math.PI,
  RAND: () => Math.random(),
  RANDBETWEEN: (a) => { const lo = toNum(a[0]); const hi = toNum(a[1]); if (isErr(lo) || isErr(hi)) return VALUE(); return Math.floor(lo + Math.random() * (hi - lo + 1)) },
  TODAY: () => excelDate(new Date()),
  NOW: () => excelDate(new Date()) + timeFraction(new Date()),
  DATE: (a) => { const y = toNum(a[0]); const m = toNum(a[1]); const d = toNum(a[2]); if (isErr(y) || isErr(m) || isErr(d)) return VALUE(); return excelDate(new Date(y, m - 1, d)) },
  YEAR: (a) => { const x = toNum(a[0]); return isErr(x) ? x : fromExcel(x).getFullYear() },
  MONTH: (a) => { const x = toNum(a[0]); return isErr(x) ? x : fromExcel(x).getMonth() + 1 },
  DAY: (a) => { const x = toNum(a[0]); return isErr(x) ? x : fromExcel(x).getDate() },
  HOUR: (a) => { const x = toNum(a[0]); return isErr(x) ? x : Math.floor((x % 1) * 24) },
  MINUTE: (a) => { const x = toNum(a[0]); return isErr(x) ? x : Math.floor(((x % 1) * 24 % 1) * 60) },
  SECOND: (a) => { const x = toNum(a[0]); return isErr(x) ? x : Math.round((((x % 1) * 24 % 1) * 60 % 1) * 60) },
  WEEKDAY: (a) => { const x = toNum(a[0]); return isErr(x) ? x : fromExcel(x).getDay() + 1 },
  LEFT: (a) => { const s = toStr(a[0]); const n = toNum(a[1] ?? 1); if (isErr(s) || isErr(n)) return VALUE(); return s.slice(0, n) },
  RIGHT: (a) => { const s = toStr(a[0]); const n = toNum(a[1] ?? 1); if (isErr(s) || isErr(n)) return VALUE(); return n === 0 ? '' : s.slice(-n) },
  MID: (a) => { const s = toStr(a[0]); const st = toNum(a[1]); const n = toNum(a[2]); if (isErr(s) || isErr(st) || isErr(n)) return VALUE(); return s.substr(st - 1, n) },
  LEN: (a) => { const s = toStr(a[0]); return isErr(s) ? s : s.length },
  TRIM: (a) => { const s = toStr(a[0]); return isErr(s) ? s : s.replace(/\s+/g, ' ').trim() },
  UPPER: (a) => { const s = toStr(a[0]); return isErr(s) ? s : s.toUpperCase() },
  LOWER: (a) => { const s = toStr(a[0]); return isErr(s) ? s : s.toLowerCase() },
  PROPER: (a) => { const s = toStr(a[0]); return isErr(s) ? s : s.replace(/\w\S*/g, (t) => t[0].toUpperCase() + t.slice(1).toLowerCase()) },
  SUBSTITUTE: (a) => { const s = toStr(a[0]); const o = toStr(a[1]); const nw = toStr(a[2]); if (isErr(s) || isErr(o) || isErr(nw)) return VALUE(); return s.split(o).join(nw) },
  REPLACE: (a) => { const s = toStr(a[0]); const st = toNum(a[1]); const n = toNum(a[2]); const nw = toStr(a[3]); if (isErr(s) || isErr(st) || isErr(n) || isErr(nw)) return VALUE(); return s.slice(0, st - 1) + nw + s.slice(st - 1 + n) },
  FIND: (a) => { const f = toStr(a[0]); const s = toStr(a[1]); const st = toNum(a[2] ?? 1); if (isErr(f) || isErr(s) || isErr(st)) return VALUE(); const i = s.indexOf(f, st - 1); return i < 0 ? VALUE() : i + 1 },
  SEARCH: (a) => { const f = toStr(a[0]); const s = toStr(a[1]); const st = toNum(a[2] ?? 1); if (isErr(f) || isErr(s) || isErr(st)) return VALUE(); const i = s.toLowerCase().indexOf(f.toLowerCase(), st - 1); return i < 0 ? VALUE() : i + 1 },
  VALUE: (a) => { const s = toStr(a[0]); if (isErr(s)) return s; const n = Number(s); return isNaN(n) ? VALUE() : n },
  CONCATENATE: (a) => { const e = errFirst(a); if (e) return e; return flat(a).map((x) => toStr(x)).filter((x) => !isErr(x)).join('') },
  CONCAT: (a) => FNS.CONCATENATE(a, null as unknown as EvalCtx),
  EXACT: (a) => { const x = toStr(a[0]); const y = toStr(a[1]); if (isErr(x) || isErr(y)) return VALUE(); return x === y },
  REPT: (a) => { const s = toStr(a[0]); const n = toNum(a[1]); if (isErr(s) || isErr(n)) return VALUE(); return s.repeat(Math.max(0, Math.floor(n))) },
  AND: (a) => { const e = errFirst(a); if (e) return e; return flat(a).every((x) => toBool(x) === true) },
  OR: (a) => { const e = errFirst(a); if (e) return e; return flat(a).some((x) => toBool(x) === true) },
  NOT: (a) => { const b = toBool(a[0]); return isErr(b) ? b : !b },
  XOR: (a) => { const e = errFirst(a); if (e) return e; return flat(a).filter((x) => toBool(x) === true).length % 2 === 1 },
  IF: (a) => { const c = toBool(a[0]); if (isErr(c)) return c; return c ? (a[1] ?? true) : (a[2] ?? false) },
  IFERROR: (a) => isErr(a[0]) ? (a[1] ?? '') : a[0],
  ISBLANK: (a) => a[0] === '' || a[0] == null,
  ISNUMBER: (a) => typeof a[0] === 'number',
  ISTEXT: (a) => typeof a[0] === 'string',
  ISERROR: (a) => isErr(a[0]),
  ISNA: (a) => isErr(a[0]) && a[0].code === '#N/A',
  ISLOGICAL: (a) => typeof a[0] === 'boolean',
  CHOOSE: (a) => { const i = toNum(a[0]); if (isErr(i)) return i; return a[i] ?? VALUE() },
  VLOOKUP: (a) => lookup(a, ctx0, true),
  HLOOKUP: (a) => lookup(a, ctx0, false),
  INDEX: (a) => {
    const arr = flat(a[0]); const r = toNum(a[1] ?? 1); const c = toNum(a[2] ?? 1)
    if (isErr(r) || isErr(c)) return VALUE()
    const dims = (a[0] as RangeVal).__dims
    const cols = dims ? dims[1] : 1
    return arr[(r - 1) * cols + (c - 1)] ?? REF()
  },
  MATCH: (a) => {
    const target = a[0]; const arr = flat(a[1]); const type = toNum(a[2] ?? 1)
    for (let i = 0; i < arr.length; i++) {
      const cmp = compare(arr[i], target, '=')
      if (cmp === true) return i + 1
    }
    return NA()
  },
  SUMIF: (a) => sumIf(a, false),
  COUNTIF: (a) => countIf(a),
  AVERAGEIF: (a) => { const s = sumIf(a, true); return s },
  SUMIFS: (a) => sumIfs(a),
  COUNTIFS: (a) => countIfs(a),
  ROW: (a) => a[0] != null ? 0 : (ctx0.current ? ctx0.current.r + 1 : 0),
  COLUMN: (a) => a[0] != null ? 0 : (ctx0.current ? ctx0.current.c + 1 : 0),
  ROWS: (a) => { const d = (a[0] as RangeVal).__dims; return d ? d[0] : (Array.isArray(a[0]) ? a[0].length : 1) },
  COLUMNS: (a) => { const d = (a[0] as RangeVal).__dims; return d ? d[1] : 1 },
  OFFSET: (a) => REF(),
  INDIRECT: (a) => REF(),
  TRANSPOSE: (a) => flat(a[0]),
  TEXT: (a) => { const v = a[0]; const f = toStr(a[1]); if (isErr(f)) return f; return String(v) },
}
let ctx0: EvalCtx = { get: () => REF(), range: () => [] }

function criteriaMatch(v: Val, crit: Val): boolean {
  if (typeof crit === 'string') {
    const m = /^(<=|>=|<>|<|>|=)(.*)$/.exec(crit)
    if (m) {
      const rhs = isNaN(Number(m[2])) ? m[2] : Number(m[2])
      const r = compare(v, rhs as Val, m[1] === '=' ? '=' : m[1])
      return r === true
    }
    if (crit.includes('*') || crit.includes('?')) {
      const re = new RegExp('^' + crit.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i')
      return re.test(String(v))
    }
    return compare(v, isNaN(Number(crit)) ? crit : Number(crit), '=') === true
  }
  return compare(v, crit, '=') === true
}
function countIf(a: Val[]): Val {
  const range = flat(a[0]); const crit = a[1]
  let n = 0
  for (const v of range) if (criteriaMatch(v, crit)) n++
  return n
}
function sumIf(a: Val[], avg: boolean): Val {
  const range = flat(a[0]); const crit = a[1]; const sumRange = a[2] ? flat(a[2]) : range
  let sum = 0, n = 0
  for (let i = 0; i < range.length; i++) if (criteriaMatch(range[i], crit)) { const x = toNum(sumRange[i] ?? 0); if (!isErr(x)) { sum += x; n++ } }
  return avg ? (n ? sum / n : DIV0()) : sum
}
function sumIfs(a: Val[]): Val {
  const sumRange = flat(a[0]); let sum = 0
  const pairs: [Val[], Val][] = []
  for (let i = 1; i + 1 < a.length + 1; i += 2) pairs.push([flat(a[i]), a[i + 1]])
  for (let i = 0; i < sumRange.length; i++) {
    if (pairs.every(([r, c]) => criteriaMatch(r[i], c))) { const x = toNum(sumRange[i]); if (!isErr(x)) sum += x }
  }
  return sum
}
function countIfs(a: Val[]): Val {
  const pairs: [Val[], Val][] = []
  for (let i = 0; i + 1 < a.length + 1; i += 2) pairs.push([flat(a[i]), a[i + 1]])
  const len = pairs[0]?.[0].length || 0
  let n = 0
  for (let i = 0; i < len; i++) if (pairs.every(([r, c]) => criteriaMatch(r[i], c))) n++
  return n
}
function lookup(a: Val[], ctx: EvalCtx, vertical: boolean): Val {
  const target = a[0]; const table = a[1]; const idx = toNum(a[2]); const approx = a[3] != null ? toBool(a[3]) : true
  if (isErr(idx)) return VALUE()
  // table is a flat row-major array; we need dimensions — reconstruct via range not available; assume caller passes 2D via range flat with known width from ctx
  const arr = flat(table)
  // Without width info, treat VLOOKUP table as columns=idx max; approximate using square? Use ctx range dims stored.
  const dims = (table as unknown as { __dims?: [number, number] }).__dims
  const rows = dims ? dims[0] : arr.length
  const cols = dims ? dims[1] : 1
  if (vertical) {
    let best = -1
    for (let r = 0; r < rows; r++) {
      const v = arr[r * cols]
      if (compare(v, target, '=') === true) { best = r; break }
      if (approx === true && compare(v, target, '<') === true) best = r
    }
    if (best < 0) return NA()
    return arr[best * cols + (idx - 1)] ?? REF()
  } else {
    let best = -1
    for (let c = 0; c < cols; c++) {
      const v = arr[c]
      if (compare(v, target, '=') === true) { best = c; break }
      if (approx === true && compare(v, target, '<') === true) best = c
    }
    if (best < 0) return NA()
    return arr[(idx - 1) * cols + best] ?? REF()
  }
}

export function excelDate(d: Date): number {
  const epoch = Date.UTC(1899, 11, 30)
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())
  let days = Math.round((utc - epoch) / 86400000)
  if (days >= 60) days++ // Excel's 1900 leap-year bug
  return days
}
export function fromExcel(n: number): Date {
  let days = Math.floor(n)
  if (days >= 60) days--
  return new Date(Date.UTC(1899, 11, 30) + days * 86400000)
}
function timeFraction(d: Date): number {
  return (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400
}

export function callFn(name: string, args: Val[], ctx: EvalCtx): Val {
  ctx0 = ctx
  const fn = FNS[name]
  if (!fn) return NAME()
  try { return fn(args, ctx) } catch (e) { return e instanceof Err ? e : VALUE() }
}

/* ---------- dependency-ordered recalc ---------- */
export function collectRefs(n: Node, out: { refs: RefNode[]; ranges: RangeNode[] }): void {
  switch (n.k) {
    case 'ref': out.refs.push(n.ref); break
    case 'range': out.ranges.push(n.range); break
    case 'un': case 'pct': collectRefs(n.a, out); break
    case 'bin': collectRefs(n.a, out); collectRefs(n.b, out); break
    case 'call': n.args.forEach((a) => collectRefs(a, out)); break
  }
}
