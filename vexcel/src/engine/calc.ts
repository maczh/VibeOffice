/* Dependency-ordered recalculation of all formula cells in a workbook. */
import { type Workbook, type Sheet, addr, parseAddr, colName } from './model'
import { parseFormula, evalNode, collectRefs, Err, type Val, type EvalCtx, type RangeVal, type RefNode, type RangeNode } from './formula'

interface FCell { key: string; sheet: number; r: number; c: number; ast: ReturnType<typeof parseFormula>; deps: string[] }

function sheetIndex(wb: Workbook, name?: string, fallback = 0): number {
  if (!name) return fallback
  const i = wb.sheets.findIndex((s) => s.name === name)
  return i >= 0 ? i : fallback
}

export function recalc(wb: Workbook): void {
  const fcells: FCell[] = []
  const byKey = new Map<string, FCell>()
  wb.sheets.forEach((sheet, si) => {
    for (const [key, cell] of sheet.cells) {
      if (!cell.f) continue
      const rc = parseAddr(key)!
      try {
        const ast = parseFormula(cell.f)
        const refs: { refs: RefNode[]; ranges: RangeNode[] } = { refs: [], ranges: [] }
        collectRefs(ast, refs)
        const deps: string[] = []
        for (const r of refs.refs) deps.push(sheetIndex(wb, r.sheet, si) + '!' + addr(r.r, r.c))
        for (const rg of refs.ranges) {
          const s = sheetIndex(wb, rg.sheet, si)
          for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) deps.push(s + '!' + addr(r, c))
        }
        const fc: FCell = { key: si + '!' + key, sheet: si, r: rc.r, c: rc.c, ast, deps }
        fcells.push(fc); byKey.set(fc.key, fc)
      } catch {
        cell.v = '#VALUE!'; cell.t = 'e'
      }
    }
  })

  // Kahn topological order over formula cells
  const indeg = new Map<string, number>()
  const dependents = new Map<string, string[]>()
  for (const fc of fcells) {
    let d = 0
    for (const dep of fc.deps) {
      if (byKey.has(dep)) { d++; dependents.set(dep, [...(dependents.get(dep) || []), fc.key]) }
    }
    indeg.set(fc.key, d)
  }
  const queue = fcells.filter((f) => (indeg.get(f.key) || 0) === 0).map((f) => f.key)
  const order: string[] = []
  while (queue.length) {
    const k = queue.shift()!
    order.push(k)
    for (const d of dependents.get(k) || []) {
      const n = (indeg.get(d) || 0) - 1
      indeg.set(d, n)
      if (n === 0) queue.push(d)
    }
  }
  const cyclic = new Set(fcells.filter((f) => !order.includes(f.key)).map((f) => f.key))

  const valueOf = (si: number, r: number, c: number): Val => {
    const cell = wb.sheets[si]?.cells.get(addr(r, c))
    if (!cell) return ''
    return (cell.v ?? '') as Val
  }

  for (const key of order) {
    const fc = byKey.get(key)!
    const sheet = wb.sheets[fc.sheet]
    const cell = sheet.cells.get(addr(fc.r, fc.c))!
    const ctx: EvalCtx = {
      current: { r: fc.r, c: fc.c },
      get: (ref) => valueOf(sheetIndex(wb, ref.sheet, fc.sheet), ref.r, ref.c),
      range: (rg) => {
        const s = sheetIndex(wb, rg.sheet, fc.sheet)
        const out: Val[] = []
        for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) out.push(valueOf(s, r, c))
        ;(out as RangeVal).__dims = [rg.r2 - rg.r1 + 1, rg.c2 - rg.c1 + 1]
        return out
      },
    }
    let v: Val
    try { v = evalNode(fc.ast, ctx) } catch (e) { v = e instanceof Err ? e : new Err('#VALUE!') }
    if (Array.isArray(v)) v = v.length ? v[0] : ''
    if (v instanceof Err) { cell.v = v.code; cell.t = 'e' }
    else if (typeof v === 'number') { cell.v = v; cell.t = 'n' }
    else if (typeof v === 'boolean') { cell.v = v; cell.t = 'b' }
    else { cell.v = typeof v === 'string' ? v : String(v); cell.t = 's' }
  }
  for (const key of cyclic) {
    const fc = byKey.get(key)!
    const cell = wb.sheets[fc.sheet].cells.get(addr(fc.r, fc.c))!
    cell.v = '#REF!'; cell.t = 'e'
  }
}