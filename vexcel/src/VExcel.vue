<script setup lang="ts">
import { ref, computed, onMounted, onBeforeUnmount, watch, nextTick } from 'vue'
import {
  VoMenuBar, VoToolbar, VoStatusBar, VoDialog, VoToast, VoBusy, VoRibbon, VoOnlineDialog,
  useLook, t, pickFile, downloadBlob, fetchOnline, saveOnline,
  type MenuDef, type ToolbarItem, type RibbonTab, type StatusCell, type Look,
  type OnlineSource, type SaveTarget, type OfficeApi, type Package,
} from '@voffice/shared'
import { type Workbook, type Sheet, type Cell, newWorkbook, newSheet, addr, parseAddr, colName, getCell, setCell, styleIndex, cloneWorkbook, cellText } from './engine/model'
import { recalc } from './engine/calc'
import { formatValue } from './engine/numfmt'
import { computeLayout, drawGrid, pointToCell, cellRect, normSel, HDR_W, HDR_H, type Sel, type Layout } from './engine/grid'
import { readXlsx } from './engine/xlsx-read'
import { writeXlsx } from './engine/xlsx-write'
import { parseCsv, writeCsv } from './engine/csv'

const props = defineProps<{ src?: string | File | Blob | OnlineSource; fileName?: string; look?: Look; readonly?: boolean }>()
const emit = defineEmits<{ ready: []; change: [dirty: boolean]; save: [blob: Blob, name: string]; error: [err: Error]; openRequest: []; saveRequest: [] }>()
const { look, layout } = useLook(props.look)
watch(() => props.look, (l) => { if (l) look.value = l })

const root = ref<HTMLElement | null>(null)
const canvas = ref<HTMLCanvasElement | null>(null)
const wrap = ref<HTMLElement | null>(null)
const editor = ref<HTMLTextAreaElement | null>(null)

const wb = ref<Workbook>(newWorkbook())
const sel = ref<Sel>({ r: 0, c: 0, r2: 0, c2: 0 })
const scroll = ref({ x: 0, y: 0 })
const zoom = ref(100)
const showGrid = ref(true)
const editing = ref<{ r: number; c: number; value: string } | null>(null)
const fileName = ref(props.fileName || 'Book1')
const dirty = ref(false)
const busy = ref('')
const toast = ref('')
const version = ref(0)
const wrapSize = ref({ w: 800, h: 500 })
let pkg: Package | null = null
let sheetParts: string[] = []
let sourceFormat = 'xlsx'
let onlineSource: OnlineSource | null = null
let onlineWritable = false
const undoStack: Workbook[] = []
const redoStack: Workbook[] = []
let clipboard: { cells: (Cell | undefined)[][]; baseR: number; baseC: number } | null = null
let dragging = false

const sheet = computed(() => wb.value.sheets[wb.value.active])
const layoutCache = computed<Layout>(() => { version.value; return computeLayout(sheet.value) })
const totalW = computed(() => layoutCache.value.xs[layoutCache.value.nw])
const totalH = computed(() => layoutCache.value.ys[layoutCache.value.nh])
const activeCell = computed(() => getCell(sheet.value, sel.value.r, sel.value.c))
const formulaText = computed(() => activeCell.value?.f != null ? '=' + activeCell.value.f : (activeCell.value?.v != null ? String(activeCell.value.v) : ''))
const selStats = computed(() => {
  version.value
  const s = normSel(sel.value)
  let sum = 0, count = 0, n = 0
  for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) {
    const cell = getCell(sheet.value, r, c)
    if (cell && typeof cell.v === 'number') { sum += cell.v; n++ }
    if (cell && cell.v != null && cell.v !== '') count++
  }
  return { sum, count, avg: n ? sum / n : 0, n }
})

const editorStyle = computed(() => {
  if (!editing.value) return {}
  const z = zoom.value / 100
  const rect = cellRect(layoutCache.value, sheet.value, editing.value.r, editing.value.c)
  return {
    left: (rect.x - scroll.value.x) * z + 'px',
    top: (rect.y - scroll.value.y) * z + 'px',
    width: Math.max(60, rect.w * z) + 'px',
    height: Math.max(20, rect.h * z) + 'px',
    fontSize: 11 * z + 'px',
  }
})

function notify(m: string) { toast.value = m; setTimeout(() => { if (toast.value === m) toast.value = '' }, 5000) }
function setDirty(d: boolean) { if (dirty.value !== d) { dirty.value = d; emit('change', d) } }
function redraw() {
  if (!canvas.value) return
  drawGrid(canvas.value, { sheet: sheet.value, wb: wb.value, layout: layoutCache.value, sel: sel.value, scrollX: scroll.value.x, scrollY: scroll.value.y, zoom: zoom.value / 100, showGrid: showGrid.value })
}
function refresh() { version.value++; nextTick(redraw) }
function snapshot() { undoStack.push(cloneWorkbook(wb.value)); if (undoStack.length > 50) undoStack.shift(); redoStack.length = 0 }
function undo() { const s = undoStack.pop(); if (!s) return; redoStack.push(cloneWorkbook(wb.value)); wb.value = s; recalc(wb.value); refresh(); setDirty(true) }
function redo() { const s = redoStack.pop(); if (!s) return; undoStack.push(cloneWorkbook(wb.value)); wb.value = s; recalc(wb.value); refresh(); setDirty(true) }

/* ---------- editing ---------- */
function commitEdit(move: 'down' | 'right' | 'none' = 'down') {
  if (!editing.value) return
  const { r, c, value } = editing.value
  editing.value = null
  snapshot()
  applyValue(r, c, value)
  if (move === 'down') sel.value = { r: r + 1, c, r2: r + 1, c2: c }
  else if (move === 'right') sel.value = { r, c: c + 1, r2: r, c2: c + 1 }
  refresh()
  nextTick(() => root.value?.focus())
}
function applyValue(r: number, c: number, text: string) {
  const prev = getCell(sheet.value, r, c)
  let cell: Cell | undefined
  if (text === '') cell = prev?.s != null ? { s: prev.s } : undefined
  else if (text.startsWith('=')) cell = { f: text.slice(1), s: prev?.s }
  else {
    const num = Number(text)
    const isNum = text.trim() !== '' && !isNaN(num) && /^[-+]?[\d.,]+(e[-+]?\d+)?%?$/i.test(text.trim())
    cell = isNum ? { v: num, t: 'n', s: prev?.s } : { v: text, t: 's', s: prev?.s }
  }
  setCell(sheet.value, r, c, cell)
  recalc(wb.value)
  setDirty(true)
}
function startEdit(initial?: string) {
  if (props.readonly) return
  const { r, c } = sel.value
  editing.value = { r, c, value: initial != null ? initial : (activeCell.value?.f != null ? '=' + activeCell.value.f : String(activeCell.value?.v ?? '')) }
  nextTick(() => { editor.value?.focus(); if (initial == null) editor.value?.select() })
}
function clearCells() {
  snapshot()
  const s = normSel(sel.value)
  for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) {
    const prev = getCell(sheet.value, r, c)
    setCell(sheet.value, r, c, prev?.s != null ? { s: prev.s } : undefined)
  }
  recalc(wb.value); setDirty(true); refresh()
}
function fmtCell(mut: (s: import('./engine/model').CellStyle) => void) {
  snapshot()
  const s = normSel(sel.value)
  for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) {
    const prev = getCell(sheet.value, r, c) || {}
    const style = { ...(prev.s != null ? wb.value.styles[prev.s] : {}) }
    mut(style)
    setCell(sheet.value, r, c, { ...prev, s: styleIndex(wb.value, style) })
  }
  setDirty(true); refresh()
}

/* ---------- mouse ---------- */
function evPoint(e: MouseEvent): { x: number; y: number } {
  const rect = wrap.value!.getBoundingClientRect()
  return { x: e.clientX - rect.left + scroll.value.x, y: e.clientY - rect.top + scroll.value.y }
}
function onDown(e: MouseEvent) {
  if (editing.value) commitEdit('none')
  const { x, y } = evPoint(e)
  if (x < HDR_W && y < HDR_H) { sel.value = { r: 0, c: 0, r2: layoutCache.value.nh - 1, c2: layoutCache.value.nw - 1 }; refresh(); return }
  if (y < HDR_H) { const c = pointToCell(layoutCache.value, x, HDR_H + 1)!.c; sel.value = { r: 0, c, r2: layoutCache.value.nh - 1, c2: c }; refresh(); return }
  if (x < HDR_W) { const r = pointToCell(layoutCache.value, HDR_W + 1, y)!.r; sel.value = { r, c: 0, r2: r, c2: layoutCache.value.nw - 1 }; refresh(); return }
  const rc = pointToCell(layoutCache.value, x, y)
  if (!rc) return
  if (e.shiftKey) sel.value = { ...sel.value, r2: rc.r, c2: rc.c }
  else sel.value = { r: rc.r, c: rc.c, r2: rc.r, c2: rc.c }
  dragging = true
  refresh()
}
function onMove(e: MouseEvent) {
  if (!dragging) return
  const { x, y } = evPoint(e)
  const rc = pointToCell(layoutCache.value, Math.max(HDR_W, x), Math.max(HDR_H, y))
  if (rc) { sel.value = { ...sel.value, r2: rc.r, c2: rc.c }; refresh() }
}
function onUp() { dragging = false }
function onDblClick(e: MouseEvent) { const { x, y } = evPoint(e); if (x > HDR_W && y > HDR_H) startEdit() }
function onScroll() { scroll.value = { x: wrap.value!.scrollLeft, y: wrap.value!.scrollTop }; redraw() }

/* ---------- clipboard ---------- */
function copySel(cut = false) {
  const s = normSel(sel.value)
  const cells: (Cell | undefined)[][] = []
  for (let r = s.r1; r <= s.r2; r++) { const row: (Cell | undefined)[] = []; for (let c = s.c1; c <= s.c2; c++) row.push(getCell(sheet.value, r, c)); cells.push(row) }
  clipboard = { cells, baseR: s.r1, baseC: s.c1 }
  const tsv = cells.map((row) => row.map((cell) => cell ? String(cell.v ?? '') : '').join('\t')).join('\n')
  navigator.clipboard?.writeText(tsv).catch(() => {})
  if (cut) clearCells()
}
function paste() {
  if (!clipboard) return
  snapshot()
  const dr = sel.value.r - clipboard.baseR, dc = sel.value.c - clipboard.baseC
  clipboard.cells.forEach((row, i) => row.forEach((cell, j) => {
    if (!cell) return
    const r = sel.value.r + i, c = sel.value.c + j
    let nc: Cell = { ...cell }
    if (nc.f) nc.f = translateRefs(nc.f, dr, dc)
    setCell(sheet.value, r, c, nc)
  }))
  recalc(wb.value); setDirty(true); refresh()
}
function translateRefs(f: string, dr: number, dc: number): string {
  return f.replace(/(\$?)([A-Za-z]+)(\$?)(\d+)/g, (m, d1, col, d2, row) => {
    const c = d1 ? col : colName(parseAddr(col + '1')!.c + dc)
    const r = d2 ? row : String(+row + dr)
    return d1 + c + d2 + r
  })
}
async function onPasteEvent(e: ClipboardEvent) {
  const text = e.clipboardData?.getData('text/plain')
  if (!text || clipboard) return
  e.preventDefault()
  snapshot()
  const rows = text.replace(/\r/g, '').split('\n').filter((r, i, a) => !(r === '' && i === a.length - 1))
  rows.forEach((row, i) => row.split('\t').forEach((val, j) => { if (val !== '') applyValue(sel.value.r + i, sel.value.c + j, val) }))
  refresh()
}

/* ---------- sheets ---------- */
function addSheet() { snapshot(); wb.value.sheets.push(newSheet('Sheet' + (wb.value.sheets.length + 1))); wb.value.active = wb.value.sheets.length - 1; refresh(); setDirty(true) }
function selectSheet(i: number) { wb.value.active = i; sel.value = { r: 0, c: 0, r2: 0, c2: 0 }; scroll.value = { x: 0, y: 0 }; refresh() }
function renamePrompt(i: number) { renameSheet(i, window.prompt('Sheet name', wb.value.sheets[i].name) || wb.value.sheets[i].name) }
function renameSheet(i: number, name: string) { if (name) { wb.value.sheets[i].name = name; setDirty(true); refresh() } }
function deleteSheet(i: number) { if (wb.value.sheets.length <= 1) return; snapshot(); wb.value.sheets.splice(i, 1); wb.value.active = Math.min(wb.value.active, wb.value.sheets.length - 1); refresh(); setDirty(true) }

/* ---------- open / save ---------- */
async function openBytes(bytes: Uint8Array, name: string) {
  const isCsv = /\.(csv|txt)$/i.test(name)
  if (isCsv) { wb.value = parseCsv(new TextDecoder().decode(bytes)); pkg = null; sheetParts = []; sourceFormat = 'csv' }
  else { const res = await readXlsx(bytes); wb.value = res.wb; pkg = res.pkg; sheetParts = res.sheetParts; sourceFormat = pkg.variant('xlsx') }
  recalc(wb.value)
  fileName.value = name.replace(/\.(xlsx|xlsm|xltx|xltm|csv|txt)$/i, '')
  sel.value = { r: 0, c: 0, r2: 0, c2: 0 }
  undoStack.length = 0; redoStack.length = 0
  setDirty(false); refresh(); emit('ready')
}
async function open(input: File | Blob | string | OnlineSource, name?: string) {
  busy.value = 'Opening…'
  try {
    onlineSource = null; onlineWritable = false
    if (typeof input === 'string') { const f = await fetchOnline({ kind: 'http', url: input }); await openBytes(new Uint8Array(await f.blob.arrayBuffer()), name || f.name) }
    else if (input instanceof Blob) { await openBytes(new Uint8Array(await input.arrayBuffer()), name || (input instanceof File ? input.name : 'Book1.xlsx')) }
    else { const f = await fetchOnline(input); onlineSource = input; onlineWritable = input.kind !== 'http' && input.kind !== 'presigned' || (input.kind === 'presigned' && !!input.putUrl); await openBytes(new Uint8Array(await f.blob.arrayBuffer()), name || f.name) }
  } catch (e) { emit('error', e as Error); notify((e as Error).message) }
  finally { busy.value = '' }
}
function newDocument() { wb.value = newWorkbook(); pkg = null; sheetParts = []; sourceFormat = 'xlsx'; fileName.value = 'Book1'; onlineSource = null; undoStack.length = 0; setDirty(false); refresh(); emit('ready') }
async function getBlob(): Promise<Blob> {
  if (sourceFormat === 'csv') return new Blob([writeCsv(wb.value, sheet.value)], { type: 'text/csv' })
  return writeXlsx(wb.value, { pkg: pkg || undefined, sheetParts, format: sourceFormat })
}
async function save() {
  busy.value = 'Saving…'
  try {
    const blob = await getBlob()
    const name = fileName.value + '.' + sourceFormat
    if (onlineSource && onlineWritable) { await saveOnline(onlineSource as SaveTarget, blob, name); setDirty(false); notify('Saved.') }
    else { emit('save', blob, name); setDirty(false) }
  } catch (e) { emit('error', e as Error); notify((e as Error).message) }
  finally { busy.value = '' }
}
async function saveAs(name?: string): Promise<Blob> { const blob = await getBlob(); downloadBlob(blob, (name || fileName.value) + '.' + sourceFormat); setDirty(false); return blob }
async function saveTo(target: SaveTarget, name?: string) { busy.value = 'Saving…'; try { await saveOnline(target, await getBlob(), (name || fileName.value) + '.' + sourceFormat); setDirty(false); notify('Saved.') } catch (e) { emit('error', e as Error); notify((e as Error).message) } finally { busy.value = '' } }

/* ---------- commands / menus ---------- */
const dlg = ref<'' | 'online-open' | 'online-save' | 'goto' | 'about'>('')
function autoSum() {
  const { r, c } = sel.value
  let end = r - 1
  while (end >= 0 && typeof getCell(sheet.value, end, c)?.v === 'number') end--
  const start = end + 1
  if (start <= r - 1) { applyValue(r, c, `=SUM(${addr(start, c)}:${addr(r - 1, c)})`); refresh() }
  else { applyValue(r, c, '=SUM(' + addr(r, c) + ')'); refresh() }
}
function insertRows() { snapshot(); const s = normSel(sel.value); const n = s.r2 - s.r1 + 1; shiftCells(s.r1, n, 0); setDirty(true); refresh() }
function insertCols() { snapshot(); const s = normSel(sel.value); const n = s.c2 - s.c1 + 1; shiftCells(s.c1, 0, n); setDirty(true); refresh() }
function shiftCells(from: number, dr: number, dc: number) {
  const entries = [...sheet.value.cells.entries()]
  sheet.value.cells.clear()
  for (const [key, cell] of entries) {
    const rc = parseAddr(key)!
    const r = rc.r >= from && dr ? rc.r + dr : rc.r
    const c = rc.c >= from && dc ? rc.c + dc : rc.c
    sheet.value.cells.set(addr(r, c), cell)
  }
  recalc(wb.value)
}
const fileMenu: MenuDef = { label: t('File'), items: [
  { label: t('New'), icon: 'new', key: 'Ctrl+N', run: newDocument },
  { label: t('Open…'), icon: 'open', key: 'Ctrl+O', run: async () => { const f = await pickFile('.xlsx,.xlsm,.xltx,.xltm,.csv,.txt'); if (f[0]) open(f[0]) } },
  { label: t('Open from Online…'), icon: 'open', run: () => { dlg.value = 'online-open' } },
  { sep: true },
  { label: t('Save'), icon: 'save', key: 'Ctrl+S', run: save },
  { label: t('Save As…'), icon: 'save', run: () => saveAs() },
  { label: t('Save to Online…'), icon: 'save', run: () => { dlg.value = 'online-save' } },
  { sep: true },
  { label: t('Print…'), icon: 'print', key: 'Ctrl+P', run: () => window.print() },
] }
const editMenu: MenuDef = { label: t('Edit'), items: [
  { label: t('Undo'), icon: 'undo', key: 'Ctrl+Z', run: undo }, { label: t('Redo'), icon: 'redo', key: 'Ctrl+Y', run: redo },
  { sep: true },
  { label: t('Cut'), icon: 'cut', key: 'Ctrl+X', run: () => copySel(true) }, { label: t('Copy'), icon: 'copy', key: 'Ctrl+C', run: () => copySel() }, { label: t('Paste'), icon: 'paste', key: 'Ctrl+V', run: paste },
  { sep: true },
  { label: t('Clear Contents'), icon: 'close', key: 'Del', run: clearCells },
] }
const viewMenu: MenuDef = { label: t('View'), items: [
  { label: t('Gridlines'), icon: 'grid', checked: showGrid.value, run: () => { showGrid.value = !showGrid.value; redraw() } },
  { label: t('Zoom In'), icon: 'zoom', run: () => { zoom.value = Math.min(200, zoom.value + 10); redraw() } },
  { label: t('Zoom Out'), icon: 'zoom', run: () => { zoom.value = Math.max(50, zoom.value - 10); redraw() } },
] }
const insertMenu: MenuDef = { label: t('Insert'), items: [
  { label: t('Rows'), icon: 'table', run: insertRows }, { label: t('Columns'), icon: 'table', run: insertCols },
  { label: t('Sheet'), icon: 'new', run: addSheet },
] }
const formatMenu: MenuDef = { label: t('Format'), items: [
  { label: t('Bold'), icon: 'bold', run: () => fmtCell((s) => { s.font = { ...s.font, bold: !s.font?.bold } }) },
  { label: t('Italic'), icon: 'italic', run: () => fmtCell((s) => { s.font = { ...s.font, italic: !s.font?.italic } }) },
  { label: t('Underline'), icon: 'underline', run: () => fmtCell((s) => { s.font = { ...s.font, underline: !s.font?.underline } }) },
  { sep: true },
  { label: t('Align Left'), icon: 'alignLeft', run: () => fmtCell((s) => { s.align = { ...s.align, h: 'left' } }) },
  { label: t('Center'), icon: 'alignCenter', run: () => fmtCell((s) => { s.align = { ...s.align, h: 'center' } }) },
  { label: t('Align Right'), icon: 'alignRight', run: () => fmtCell((s) => { s.align = { ...s.align, h: 'right' } }) },
  { sep: true },
  { label: t('Number'), icon: 'properties', run: () => fmtCell((s) => { s.numFmt = '0.00' }) },
  { label: t('Percent'), icon: 'properties', run: () => fmtCell((s) => { s.numFmt = '0%' }) },
  { label: t('Currency'), icon: 'properties', run: () => fmtCell((s) => { s.numFmt = '$#,##0.00' }) },
] }
const dataMenu: MenuDef = { label: t('Data'), items: [
  { label: t('Sort Ascending'), icon: 'numbering', run: () => sortRange(1) },
  { label: t('Sort Descending'), icon: 'numbering', run: () => sortRange(-1) },
] }
function sortRange(dir: number) {
  snapshot()
  const s = normSel(sel.value)
  const rows: (Cell | undefined)[][] = []
  for (let r = s.r1; r <= s.r2; r++) { const row: (Cell | undefined)[] = []; for (let c = s.c1; c <= s.c2; c++) row.push(getCell(sheet.value, r, c)); rows.push(row) }
  rows.sort((a, b) => { const av = a[0]?.v; const bv = b[0]?.v; if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir; return String(av ?? '').localeCompare(String(bv ?? '')) * dir })
  rows.forEach((row, i) => row.forEach((cell, j) => setCell(sheet.value, s.r1 + i, s.c1 + j, cell)))
  recalc(wb.value); setDirty(true); refresh()
}
const toolsMenu: MenuDef = { label: t('Tools'), items: [{ label: t('About VExcel'), icon: 'help', run: () => { dlg.value = 'about' } }] }
const helpMenu: MenuDef = { label: t('Help'), items: [{ label: t('About VExcel'), icon: 'help', run: () => { dlg.value = 'about' } }] }
const menus = computed(() => [fileMenu, editMenu, viewMenu, insertMenu, formatMenu, dataMenu, toolsMenu, helpMenu])

const fonts = ['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Georgia', 'Verdana']
const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36]
const curFont = ref('Calibri'); const curSize = ref(11)
const stdTools = computed<ToolbarItem[]>(() => [
  { kind: 'button', id: 'new', icon: 'new', tip: t('New'), run: newDocument },
  { kind: 'button', id: 'open', icon: 'open', tip: t('Open'), run: async () => { const f = await pickFile('.xlsx,.xlsm,.csv'); if (f[0]) open(f[0]) } },
  { kind: 'button', id: 'save', icon: 'save', tip: t('Save'), run: save },
  { kind: 'button', id: 'print', icon: 'print', tip: t('Print'), run: () => window.print() },
  { kind: 'sep' },
  { kind: 'button', id: 'cut', icon: 'cut', run: () => copySel(true) }, { kind: 'button', id: 'copy', icon: 'copy', run: () => copySel() }, { kind: 'button', id: 'paste', icon: 'paste', run: paste },
  { kind: 'sep' },
  { kind: 'button', id: 'undo', icon: 'undo', run: undo }, { kind: 'button', id: 'redo', icon: 'redo', run: redo },
  { kind: 'sep' },
  { kind: 'button', id: 'sum', icon: 'numbering', tip: t('AutoSum'), run: autoSum },
  { kind: 'button', id: 'sort', icon: 'numbering', tip: t('Sort'), run: () => sortRange(1) },
])
const fmtTools = computed<ToolbarItem[]>(() => [
  { kind: 'combo', id: 'font', value: curFont.value, width: 120, editable: true, items: fonts.map((f) => ({ label: f })), onChange: (v) => { curFont.value = v; fmtCell((s) => { s.font = { ...s.font, name: v } }) } },
  { kind: 'combo', id: 'size', value: String(curSize.value), width: 46, editable: true, items: sizes.map((s) => ({ label: String(s) })), onChange: (v) => { curSize.value = +v; fmtCell((s) => { s.font = { ...s.font, size: +v } }) } },
  { kind: 'sep' },
  { kind: 'button', id: 'bold', icon: 'bold', run: () => fmtCell((s) => { s.font = { ...s.font, bold: !s.font?.bold } }) },
  { kind: 'button', id: 'italic', icon: 'italic', run: () => fmtCell((s) => { s.font = { ...s.font, italic: !s.font?.italic } }) },
  { kind: 'button', id: 'underline', icon: 'underline', run: () => fmtCell((s) => { s.font = { ...s.font, underline: !s.font?.underline } }) },
  { kind: 'sep' },
  { kind: 'button', id: 'left', icon: 'alignLeft', run: () => fmtCell((s) => { s.align = { ...s.align, h: 'left' } }) },
  { kind: 'button', id: 'center', icon: 'alignCenter', run: () => fmtCell((s) => { s.align = { ...s.align, h: 'center' } }) },
  { kind: 'button', id: 'right', icon: 'alignRight', run: () => fmtCell((s) => { s.align = { ...s.align, h: 'right' } }) },
  { kind: 'sep' },
  { kind: 'button', id: 'fill', icon: 'fillColor', tip: t('Fill Color'), run: () => fmtCell((s) => { s.fill = s.fill ? undefined : '#ffff00' }) },
  { kind: 'button', id: 'border', icon: 'tablesBorders', tip: t('Borders'), run: () => fmtCell((s) => { s.border = s.border ? undefined : { top: '#000', bottom: '#000', left: '#000', right: '#000' } }) },
])
const ribbonTabs = computed<RibbonTab[]>(() => [
  { id: 'home', label: t('Home'), groups: [ { label: t('Clipboard'), items: stdTools.value.slice(5, 10) }, { label: t('Font'), items: fmtTools.value.slice(0, 7) }, { label: t('Alignment'), items: fmtTools.value.slice(8, 13) }, { label: t('Cells'), items: [ { kind: 'button', id: 'sum', icon: 'numbering', label: t('AutoSum'), run: autoSum } ] } ] },
  { id: 'insert', label: t('Insert'), groups: [ { label: t('Cells'), items: [ { kind: 'button', id: 'rows', icon: 'table', label: t('Rows'), run: insertRows }, { kind: 'button', id: 'cols', icon: 'table', label: t('Columns'), run: insertCols }, { kind: 'button', id: 'sheet', icon: 'new', label: t('Sheet'), run: addSheet } ] } ] },
  { id: 'data', label: t('Data'), groups: [ { label: t('Sort'), items: [ { kind: 'button', id: 'asc', icon: 'numbering', label: t('A→Z'), run: () => sortRange(1) }, { kind: 'button', id: 'desc', icon: 'numbering', label: t('Z→A'), run: () => sortRange(-1) } ] } ] },
])
const statusCells = computed<StatusCell[]>(() => {
  const s = selStats.value
  return [
    { text: t('Ready'), grow: true },
    ...(s.n ? [{ text: `Sum: ${formatValue(s.sum)}` }, { text: `Avg: ${formatValue(s.avg)}` }, { text: `Count: ${s.count}` }] : []),
    { text: zoom.value + '%', act: true, run: () => { zoom.value = zoom.value >= 200 ? 50 : zoom.value + 10; redraw() } },
  ]
})

/* ---------- keyboard ---------- */
function onKey(e: KeyboardEvent) {
  if (editing.value) {
    if (e.key === 'Enter') { e.preventDefault(); commitEdit(e.shiftKey ? 'none' : 'down') }
    else if (e.key === 'Tab') { e.preventDefault(); commitEdit('right') }
    else if (e.key === 'Escape') { editing.value = null; refresh(); nextTick(() => root.value?.focus()) }
    return
  }
  const mod = e.ctrlKey || e.metaKey
  const k = e.key.toLowerCase()
  if (mod) {
    const map: Record<string, () => void> = {
      s: save, o: async () => { const f = await pickFile('.xlsx,.xlsm,.csv'); if (f[0]) open(f[0]) }, n: newDocument, p: () => window.print(),
      z: () => (e.shiftKey ? redo() : undo()), y: redo, c: () => copySel(), x: () => copySel(true), v: paste,
      b: () => fmtCell((s) => { s.font = { ...s.font, bold: !s.font?.bold } }), i: () => fmtCell((s) => { s.font = { ...s.font, italic: !s.font?.italic } }),
      home: () => { sel.value = { r: 0, c: 0, r2: 0, c2: 0 }; scroll.value = { x: 0, y: 0 }; refresh() },
    }
    if (map[k]) { e.preventDefault(); map[k]() }
    return
  }
  const move = (dr: number, dc: number, extend = false) => {
    const r = Math.max(0, sel.value.r + dr), c = Math.max(0, sel.value.c + dc)
    sel.value = extend ? { ...sel.value, r2: r, c2: c } : { r, c, r2: r, c2: c }
    scrollIntoView(r, c); refresh()
  }
  switch (e.key) {
    case 'ArrowUp': e.preventDefault(); move(-1, 0, e.shiftKey); break
    case 'ArrowDown': e.preventDefault(); move(1, 0, e.shiftKey); break
    case 'ArrowLeft': e.preventDefault(); move(0, -1, e.shiftKey); break
    case 'ArrowRight': e.preventDefault(); move(0, 1, e.shiftKey); break
    case 'Tab': e.preventDefault(); move(0, e.shiftKey ? -1 : 1); break
    case 'Enter': e.preventDefault(); startEdit(); break
    case 'F2': e.preventDefault(); startEdit(); break
    case 'Delete': case 'Backspace': e.preventDefault(); clearCells(); break
    case 'PageDown': e.preventDefault(); move(20, 0, e.shiftKey); break
    case 'PageUp': e.preventDefault(); move(-20, 0, e.shiftKey); break
    default:
      if (e.key.length === 1 && !e.altKey) { e.preventDefault(); startEdit(e.key) }
  }
}
function scrollIntoView(r: number, c: number) {
  const rect = cellRect(layoutCache.value, sheet.value, r, c)
  const vw = wrapSize.value.w, vh = wrapSize.value.h
  let { x, y } = scroll.value
  if (rect.x - x < HDR_W) x = rect.x - HDR_W
  if (rect.x + rect.w - x > vw) x = rect.x + rect.w - vw
  if (rect.y - y < HDR_H) y = rect.y - HDR_H
  if (rect.y + rect.h - y > vh) y = rect.y + rect.h - vh
  scroll.value = { x: Math.max(0, x), y: Math.max(0, y) }
  if (wrap.value) { wrap.value.scrollLeft = scroll.value.x; wrap.value.scrollTop = scroll.value.y }
}
function onOnlineSubmit(source: OnlineSource | SaveTarget) { if (dlg.value === 'online-open') { dlg.value = ''; open(source as OnlineSource) } else { dlg.value = ''; saveTo(source as SaveTarget) } }

let ro: ResizeObserver | null = null
onMounted(async () => {
  ro = new ResizeObserver(() => { if (wrap.value) { wrapSize.value = { w: wrap.value.clientWidth, h: wrap.value.clientHeight }; redraw() } })
  if (wrap.value) ro.observe(wrap.value)
  if (props.src) await open(props.src, props.fileName); else newDocument()
  root.value?.focus()
  refresh()
})
onBeforeUnmount(() => ro?.disconnect())
watch(() => props.src, (s) => { if (s) open(s, props.fileName) })

defineExpose<OfficeApi>({ open, newDocument, getBlob, save, saveAs, saveTo, get name() { return fileName.value }, get dirty() { return dirty.value } })
</script>

<template>
  <div ref="root" class="vo-app" :data-look="look" :data-layout="layout" tabindex="0" @keydown="onKey" @paste="onPasteEvent">
    <div class="app">
      <div class="cmdbars">
        <div class="menurow"><VoMenuBar :menus="menus" /></div>
        <div class="toolbars">
          <div class="tb-row"><VoToolbar :items="stdTools" /></div>
          <div class="tb-row"><VoToolbar :items="fmtTools" /></div>
        </div>
        <VoRibbon :tabs="ribbonTabs" @file="dlg = 'online-open'" />
      </div>
      <div class="formulabar">
        <input class="namebox" :value="addr(sel.r, sel.c)" readonly @focus="($event.target as HTMLInputElement).select()">
        <span class="fx">fx</span>
        <input class="finput" :value="editing ? editing.value : formulaText" @input="editing ? editing.value = ($event.target as HTMLInputElement).value : (editing = { r: sel.r, c: sel.c, value: ($event.target as HTMLInputElement).value })" @keydown.enter="commitEdit('down')" @focus="!editing && (editing = { r: sel.r, c: sel.c, value: formulaText })">
      </div>
      <div class="main">
        <div class="center">
          <div class="gridarea">
            <canvas ref="canvas" class="gridcanvas"></canvas>
            <div ref="wrap" class="gridscroll" @scroll="onScroll" @mousedown="onDown" @mousemove="onMove" @mouseup="onUp" @mouseleave="onUp" @dblclick="onDblClick">
              <div class="gridsizer" :style="{ width: totalW + 'px', height: totalH + 'px' }"></div>
            </div>
            <textarea v-if="editing" ref="editor" class="celleditor" :value="editing.value"
              @input="editing.value = ($event.target as HTMLTextAreaElement).value"
              :style="editorStyle"></textarea>
          </div>
        </div>
      </div>
      <div class="stabrow">
        <button class="stab add" title="New sheet" @click="addSheet">＋</button>
        <button v-for="(s, i) in wb.sheets" :key="i" class="stab" :class="{ on: i === wb.active }"
          @click="selectSheet(i)" @dblclick="renamePrompt(i)"
          @contextmenu.prevent="deleteSheet(i)">{{ s.name }}</button>
      </div>
      <VoStatusBar :cells="statusCells" />
    </div>

    <VoDialog v-if="dlg === 'about'" :title="t('About VExcel')" :width="360" @close="dlg = ''">
      <div class="about"><div class="about-head"><span class="about-name">VExcel</span></div>
        <p>{{ t('A Vue 3 spreadsheet with a formula engine that opens and saves Excel workbooks.') }}</p></div>
      <template #foot><button class="btn primary" @click="dlg = ''">{{ t('OK') }}</button></template>
    </VoDialog>
    <VoOnlineDialog v-if="dlg === 'online-open'" mode="open" @close="dlg = ''" @submit="onOnlineSubmit" />
    <VoOnlineDialog v-if="dlg === 'online-save'" mode="save" :file-name="fileName + '.' + sourceFormat" @close="dlg = ''" @submit="onOnlineSubmit" />
    <VoBusy v-if="busy" :label="busy" />
    <VoToast v-if="toast" :message="toast" @done="toast = ''" />
  </div>
</template>

<script lang="ts">
export default { name: 'VExcel' }
</script>

<style scoped>
.vo-app { outline: none; display: flex; flex-direction: column; }
.app { display: flex; flex-direction: column; height: 100%; }
.formulabar { flex: none; display: flex; align-items: center; gap: 4px; padding: 2px 4px; background: var(--dock-b); border-bottom: 1px solid var(--chrome-line); }
.namebox { width: 70px; border: 1px solid var(--field-border); background: var(--surface); height: 20px; padding: 0 4px; text-align: center; }
.fx { color: var(--muted); font-style: italic; padding: 0 4px; }
.finput { flex: 1; border: 1px solid var(--field-border); background: var(--surface); height: 20px; padding: 0 4px; }
.main { flex: 1; min-height: 0; display: flex; }
.center { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.gridarea { position: relative; flex: 1; min-height: 0; overflow: hidden; background: var(--surface); }
.gridcanvas { position: absolute; inset: 0; width: 100%; height: 100%; }
.gridscroll { position: absolute; inset: 0; overflow: auto; }
.gridsizer { pointer-events: none; }
.celleditor { position: absolute; z-index: 5; border: 2px solid var(--sel-line, #1b3d7a); background: #fff; padding: 0 3px; font: 11px Calibri, Carlito, sans-serif; resize: none; overflow: hidden; outline: none; }
.stabrow { flex: none; display: flex; align-items: center; gap: 2px; padding: 2px 4px; background: var(--dock-b); border-top: 1px solid var(--chrome-line); overflow-x: auto; }
.stab { border: 1px solid var(--sheet-tab-edge); background: var(--face-2); padding: 0 12px; height: 20px; cursor: default; }
.stab.on { background: var(--surface); font-weight: bold; }
.stab.add { padding: 0 8px; }
</style>