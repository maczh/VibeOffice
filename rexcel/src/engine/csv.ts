/* CSV / TSV parse and serialize. */
import { type Workbook, type Sheet, newSheet, newWorkbook, addr } from './model'

export function parseCsv(text: string, delim = ','): Workbook {
  const wb = newWorkbook()
  const sheet = wb.sheets[0]
  sheet.name = 'Sheet1'
  const rows = splitRows(text, delim)
  rows.forEach((row, r) => {
    row.forEach((val, c) => {
      if (val === '') return
      const num = Number(val)
      const isNum = val.trim() !== '' && !isNaN(num) && /^[-+]?[\d.]+(e[-+]?\d+)?$/i.test(val.trim())
      sheet.cells.set(addr(r, c), isNum ? { v: num, t: 'n' } : { v: val, t: 's' })
    })
  })
  return wb
}

function splitRows(text: string, delim: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let cur = ''
  let inQ = false
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (inQ) {
      if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++ } else inQ = false }
      else cur += ch
    } else if (ch === '"') inQ = true
    else if (ch === delim) { row.push(cur); cur = '' }
    else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = '' }
    else if (ch === '\r') { /* skip */ }
    else cur += ch
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row) }
  return rows
}

export function writeCsv(wb: Workbook, sheet: Sheet, delim = ','): string {
  let maxR = 0, maxC = 0
  for (const key of sheet.cells.keys()) {
    const m = /^([A-Z]+)(\d+)$/.exec(key)!
    let c = 0
    for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64)
    maxC = Math.max(maxC, c - 1); maxR = Math.max(maxR, +m[2] - 1)
  }
  const lines: string[] = []
  for (let r = 0; r <= maxR; r++) {
    const cells: string[] = []
    for (let c = 0; c <= maxC; c++) {
      const cell = sheet.cells.get(addr(r, c))
      let v = cell ? String(cell.v ?? '') : ''
      if (v.includes(delim) || v.includes('"') || v.includes('\n')) v = '"' + v.replace(/"/g, '""') + '"'
      cells.push(v)
    }
    lines.push(cells.join(delim))
  }
  return lines.join('\r\n')
}
