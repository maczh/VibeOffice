/* Number-format subset: General, fixed/decimal/grouped numeric, percent, currency, text, dates. */
import { fromExcel } from './formula'

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
const MONTHS_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']

export function formatValue(v: string | number | boolean | null | undefined, fmt?: string): string {
  if (v == null) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'string') return v
  const f = fmt || 'General'
  if (f === 'General') return general(v)
  if (f === '@') return String(v)
  if (isDateFmt(f)) return formatDate(v, f)
  return formatNumeric(v, f)
}

function general(n: number): string {
  if (!isFinite(n)) return String(n)
  if (Number.isInteger(n) && Math.abs(n) < 1e11) return String(n)
  const s = n.toPrecision(11)
  return String(parseFloat(s))
}

function isDateFmt(f: string): boolean {
  if (/[#0]/.test(f)) return false
  return /[ymdhs]/i.test(f)
}

function pad(n: number, w: number): string { return String(n).padStart(w, '0') }

function formatDate(serial: number, f: string): string {
  const d = fromExcel(serial)
  const frac = serial - Math.floor(serial)
  const secs = Math.round(frac * 86400)
  const H = Math.floor(secs / 3600), M = Math.floor((secs % 3600) / 60), S = secs % 60
  const y = d.getUTCFullYear(), mo = d.getUTCMonth(), day = d.getUTCDate(), dow = d.getUTCDay()
  let out = ''
  let i = 0
  const lower = f
  while (i < lower.length) {
    const ch = lower[i]
    const run = (c: string) => { let n = 0; while (lower[i + n] === c) n++; return n }
    if (ch === 'y' || ch === 'Y') { const n = run(ch); i += n; out += n >= 3 ? String(y) : pad(y % 100, 2) }
    else if (ch === 'm' || ch === 'M') {
      const n = run(ch); i += n
      // m after h or before s means minutes
      const prev = out.slice(-1)
      const isMin = /h$/i.test(lower.slice(0, i - n)) || /^s/i.test(lower.slice(i))
      if (isMin) out += n >= 2 ? pad(M, 2) : String(M)
      else out += n >= 4 ? MONTHS_FULL[mo] : n === 3 ? MONTHS[mo] : n === 2 ? pad(mo + 1, 2) : String(mo + 1)
    }
    else if (ch === 'd' || ch === 'D') { const n = run(ch); i += n; out += n >= 4 ? DAYS[dow] : n === 3 ? DAYS[dow] : n === 2 ? pad(day, 2) : String(day) }
    else if (ch === 'h' || ch === 'H') { const n = run(ch); i += n; const h12 = /AM\/PM/i.test(f) ? (H % 12 || 12) : H; out += n >= 2 ? pad(h12, 2) : String(h12) }
    else if (ch === 's' || ch === 'S') { const n = run(ch); i += n; out += n >= 2 ? pad(S, 2) : String(S) }
    else if (ch === 'A' && /^AM\/PM/i.test(lower.slice(i))) { out += H < 12 ? 'AM' : 'PM'; i += 5 }
    else if (ch === 'a' && /^am\/pm/i.test(lower.slice(i))) { out += H < 12 ? 'am' : 'pm'; i += 5 }
    else { out += ch; i++ }
  }
  return out
}

function formatNumeric(n: number, f: string): string {
  let pattern = f
  let percent = false
  if (pattern.includes('%')) { percent = true; n *= 100; pattern = pattern.replace(/%/g, '') }
  // strip color/condition sections like [Red]
  pattern = pattern.replace(/\[[^\]]*\]/g, '')
  // take first section (positive)
  pattern = pattern.split(';')[0]
  const dot = pattern.indexOf('.')
  const decPart = dot >= 0 ? pattern.slice(dot + 1) : ''
  const decimals = (decPart.match(/[0#]/g) || []).length
  const grouping = /,/.test(dot >= 0 ? pattern.slice(0, dot) : pattern)
  // literal prefix/suffix (currency symbols, spaces, text in quotes)
  const intPart = dot >= 0 ? pattern.slice(0, dot) : pattern
  const prefix = (intPart.match(/^[^0#,]*/) || [''])[0].replace(/"/g, '')
  const suffix = (decPart.match(/[^0#]*$/) || [''])[0].replace(/"/g, '') + (percent ? '%' : '')
  let body = Math.abs(n).toFixed(decimals)
  if (grouping) {
    const [ip, dp] = body.split('.')
    body = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (dp ? '.' + dp : '')
  }
  const sign = n < 0 ? '-' : ''
  return sign + prefix + body + suffix
}