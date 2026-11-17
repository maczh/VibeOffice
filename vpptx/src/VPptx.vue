<script setup lang="ts">
import { ref, computed, onMounted, onBeforeUnmount, watch, nextTick } from 'vue'
import {
  VoMenuBar, VoToolbar, VoStatusBar, VoDialog, VoToast, VoBusy, VoRibbon, VoOnlineDialog,
  useLook, t, pickFile, downloadBlob, fetchOnline, saveOnline,
  type MenuDef, type ToolbarItem, type RibbonTab, type StatusCell, type Look,
  type OnlineSource, type SaveTarget, type OfficeApi, type Package,
} from '@voffice/shared'
import { type Presentation, type Slide, type Shape, type RPr, newPresentation, newSlide, newTextShape, uid, px, emu, clonePresentation, EMU_PER_PX } from './engine/model'
import { presetSvg, PRESETS } from './engine/shapes'
import { slideHtml } from './engine/render'
import { readPptx } from './engine/pptx-read'
import { writePptx } from './engine/pptx-write'

const props = defineProps<{ src?: string | File | Blob | OnlineSource; fileName?: string; look?: Look; readonly?: boolean }>()
const emit = defineEmits<{ ready: []; change: [dirty: boolean]; save: [blob: Blob, name: string]; error: [err: Error]; openRequest: []; saveRequest: [] }>()
const { look, layout } = useLook(props.look)
watch(() => props.look, (l) => { if (l) look.value = l })

const root = ref<HTMLElement | null>(null)
const editorWrap = ref<HTMLElement | null>(null)
const pres = ref<Presentation>(newPresentation())
const cur = ref(0)
const selected = ref<string | null>(null)
const editingText = ref<string | null>(null)
const zoom = ref(60)
const tool = ref<'select' | 'textbox' | string>('select')
const fileName = ref(props.fileName || 'Presentation')
const dirty = ref(false)
const busy = ref('')
const toast = ref('')
const show = ref(false)
const showIndex = ref(0)
const dlg = ref<'' | 'online-open' | 'online-save' | 'bg' | 'about'>('')
const bgColor = ref('#ffffff')
let pkg: Package | null = null
let slideParts: string[] = []
let onlineSource: OnlineSource | null = null
let onlineWritable = false
const undoStack: Presentation[] = []
const redoStack: Presentation[] = []

const slide = computed(() => pres.value.slides[cur.value])
const slideWpx = computed(() => px(pres.value.slideW))
const slideHpx = computed(() => px(pres.value.slideH))
const scale = computed(() => zoom.value / 100)
const selShape = computed(() => slide.value.shapes.find((s) => s.id === selected.value) || null)

/** fit the slide to the editor area (as the original does on open) */
function fitZoom() {
  const el = editorWrap.value
  if (!el) return
  const w = el.clientWidth - 48, h = el.clientHeight - 48
  if (w <= 0 || h <= 0) return
  zoom.value = Math.max(20, Math.min(200, Math.floor(Math.min(w / slideWpx.value, h / slideHpx.value) * 100)))
}
let fitRo: ResizeObserver | null = null

function notify(m: string) { toast.value = m; setTimeout(() => { if (toast.value === m) toast.value = '' }, 5000) }
function setDirty(d: boolean) { if (dirty.value !== d) { dirty.value = d; emit('change', d) } }
function snapshot() { undoStack.push(clonePresentation(pres.value)); if (undoStack.length > 50) undoStack.shift(); redoStack.length = 0 }
function undo() { const s = undoStack.pop(); if (!s) return; redoStack.push(clonePresentation(pres.value)); pres.value = s; selected.value = null; setDirty(true) }
function redo() { const s = redoStack.pop(); if (!s) return; undoStack.push(clonePresentation(pres.value)); pres.value = s; selected.value = null; setDirty(true) }

/* ---------- shape geometry helpers ---------- */
function shapeStyle(s: Shape) {
  return {
    left: px(s.x) * scale.value + 'px', top: px(s.y) * scale.value + 'px',
    width: px(s.w) * scale.value + 'px', height: px(s.h) * scale.value + 'px',
    transform: s.rot ? `rotate(${s.rot}deg)` : undefined,
  }
}
function shapeSvg(s: Shape): string {
  return presetSvg(s.preset || 'rect', px(s.w), px(s.h), s.fill ? s.fill.color : null, s.line ? s.line.color : null, s.line?.width || 1)
}
function textHtml(s: Shape): string {
  return (s.text || []).map((p) => `<div style="text-align:${p.align || 'left'};min-height:1em">${p.bullet ? '• ' : ''}${p.runs.map((r) => `<span style="${r.rPr?.bold ? 'font-weight:bold;' : ''}${r.rPr?.italic ? 'font-style:italic;' : ''}${r.rPr?.size ? `font-size:${r.rPr.size}px;` : ''}${r.rPr?.color ? `color:${r.rPr.color};` : ''}">${r.text.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</span>`).join('')}</div>`).join('')
}

/* ---------- drag: move / resize / create ---------- */
interface Drag { mode: 'move' | 'resize' | 'create'; id: string; handle?: string; sx: number; sy: number; ox: number; oy: number; ow: number; oh: number; preset?: string }
let drag: Drag | null = null
function startMove(s: Shape, e: MouseEvent) {
  if (props.readonly) return
  if (tool.value !== 'select') return
  selected.value = s.id
  if (editingText.value && editingText.value !== s.id) commitText()
  snapshot()
  drag = { mode: 'move', id: s.id, sx: e.clientX, sy: e.clientY, ox: s.x, oy: s.y, ow: s.w, oh: s.h }
  e.stopPropagation()
}
function startResize(s: Shape, handle: string, e: MouseEvent) {
  snapshot()
  drag = { mode: 'resize', id: s.id, handle, sx: e.clientX, sy: e.clientY, ox: s.x, oy: s.y, ow: s.w, oh: s.h }
  e.stopPropagation(); e.preventDefault()
}
function canvasDown(e: MouseEvent) {
  if (props.readonly) return
  if (tool.value === 'select') { selected.value = null; commitText(); return }
  // create a shape by dragging
  const rect = (e.currentTarget as HTMLElement).getBoundingClientRect()
  const x = (e.clientX - rect.left) / scale.value, y = (e.clientY - rect.top) / scale.value
  snapshot()
  const preset = tool.value === 'textbox' ? 'rect' : tool.value
  const s: Shape = tool.value === 'textbox'
    ? newTextShape(emu(x), emu(y), emu(10), emu(10), '')
    : { id: uid('sh'), kind: 'sp', x: emu(x), y: emu(y), w: emu(10), h: emu(10), preset, fill: tool.value === 'line' ? null : { color: '#4472c4' }, line: { color: '#2f5597', width: 1 }, text: [] }
  slide.value.shapes.push(s)
  selected.value = s.id
  drag = { mode: 'create', id: s.id, sx: e.clientX, sy: e.clientY, ox: s.x, oy: s.y, ow: 0, oh: 0, preset }
  tool.value = 'select'
  setDirty(true)
}
function onMove(e: MouseEvent) {
  if (!drag) return
  const s = slide.value.shapes.find((x) => x.id === drag!.id)
  if (!s) return
  const dx = emu((e.clientX - drag.sx) / scale.value), dy = emu((e.clientY - drag.sy) / scale.value)
  if (drag.mode === 'move') { s.x = drag.ox + dx; s.y = drag.oy + dy }
  else if (drag.mode === 'create') { s.w = Math.max(emu(8), dx); s.h = Math.max(emu(8), dy) }
  else if (drag.mode === 'resize') {
    const h = drag.handle!
    let { ox, oy, ow, oh } = drag
    if (h.includes('e')) s.w = Math.max(emu(8), ow + dx)
    if (h.includes('s')) s.h = Math.max(emu(8), oh + dy)
    if (h.includes('w')) { s.x = ox + dx; s.w = Math.max(emu(8), ow - dx) }
    if (h.includes('n')) { s.y = oy + dy; s.h = Math.max(emu(8), oh - dy) }
  }
  setDirty(true)
}
function onUp() { if (drag) { drag = null } }
function dblShape(s: Shape, e: MouseEvent) {
  if (s.kind === 'pic' || s.raw != null) return
  selected.value = s.id
  editingText.value = s.id
  if (!s.text) s.text = [{ runs: [] }]
  nextTick(() => {
    const el = root.value?.querySelector<HTMLElement>(`[data-text="${s.id}"]`)
    el?.focus()
  })
  e.stopPropagation()
}
function commitText() {
  if (!editingText.value) return
  const s = slide.value.shapes.find((x) => x.id === editingText.value)
  const el = root.value?.querySelector<HTMLElement>(`[data-text="${editingText.value}"]`)
  if (s && el) {
    const lines = el.innerText.replace(/\n$/, '').split('\n')
    const baseRPr = s.text?.[0]?.runs[0]?.rPr
    s.text = lines.map((ln) => ({ align: s.text?.[0]?.align, runs: ln ? [{ text: ln, rPr: baseRPr }] : [] }))
    setDirty(true)
  }
  editingText.value = null
}

/* ---------- shape / slide ops ---------- */
function deleteShape() { if (!selected.value) return; snapshot(); slide.value.shapes = slide.value.shapes.filter((s) => s.id !== selected.value); selected.value = null; setDirty(true) }
function addSlide() { snapshot(); pres.value.slides.splice(cur.value + 1, 0, newSlide()); cur.value++; selected.value = null; setDirty(true) }
function dupSlide() { snapshot(); const c = clonePresentation(pres.value).slides[cur.value]; c.id = uid('sl'); pres.value.slides.splice(cur.value + 1, 0, c); cur.value++; setDirty(true) }
function deleteSlide() { if (pres.value.slides.length <= 1) return; snapshot(); pres.value.slides.splice(cur.value, 1); cur.value = Math.min(cur.value, pres.value.slides.length - 1); selected.value = null; setDirty(true) }
function moveSlide(i: number, dir: number) { const j = i + dir; if (j < 0 || j >= pres.value.slides.length) return; snapshot(); const [s] = pres.value.slides.splice(i, 1); pres.value.slides.splice(j, 0, s); cur.value = j; setDirty(true) }
function selectSlide(i: number) { commitText(); cur.value = i; selected.value = null }
async function insertPicture() {
  const files = await pickFile('image/*')
  if (!files[0]) return
  const url = await new Promise<string>((res) => { const fr = new FileReader(); fr.onload = () => res(fr.result as string); fr.readAsDataURL(files[0]) })
  const img = new Image(); await new Promise((res) => { img.onload = res; img.src = url })
  const w = Math.min(px(pres.value.slideW) * 0.5, img.width)
  const h = img.height * (w / img.width)
  snapshot()
  slide.value.shapes.push({ id: uid('sh'), kind: 'pic', x: emu(px(pres.value.slideW) / 2 - w / 2), y: emu(px(pres.value.slideH) / 2 - h / 2), w: emu(w), h: emu(h), image: { part: '', rId: '', url } })
  setDirty(true)
}
function setFill(color: string | null) { if (!selShape.value) return; snapshot(); selShape.value.fill = color ? { color } : null; setDirty(true) }
function setLine(color: string | null) { if (!selShape.value) return; snapshot(); selShape.value.line = color ? { color, width: 1 } : null; setDirty(true) }
function applyBg() { snapshot(); slide.value.background = bgColor.value; setDirty(true); dlg.value = '' }
function fmtText(mut: (r: RPr) => void) {
  if (!selShape.value?.text) return
  snapshot()
  for (const p of selShape.value.text) for (const r of p.runs) { const rp = r.rPr || (r.rPr = {}); mut(rp) }
  setDirty(true)
}

/* ---------- slideshow ---------- */
function startShow(from = 0) { showIndex.value = from; show.value = true }
function showNext() { if (showIndex.value < pres.value.slides.length - 1) showIndex.value++; else show.value = false }
function showPrev() { if (showIndex.value > 0) showIndex.value-- }
function onShowKey(e: KeyboardEvent) {
  if (e.key === 'Escape') show.value = false
  else if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter'].includes(e.key)) showNext()
  else if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(e.key)) showPrev()
}
const showHtml = computed(() => show.value ? slideHtml(pres.value.slides[showIndex.value], pres.value, Math.min(innerWidth, innerHeight * (pres.value.slideW / pres.value.slideH))) : '')

/* ---------- open / save ---------- */
async function openBytes(bytes: Uint8Array, name: string) {
  const res = await readPptx(bytes)
  pres.value = res.pres; pkg = res.pkg; slideParts = res.slideParts
  fileName.value = name.replace(/\.(pptx|pptm|ppsx|ppsm|potx|potm)$/i, '')
  cur.value = 0; selected.value = null; undoStack.length = 0
  setDirty(false); emit('ready')
}
async function open(input: File | Blob | string | OnlineSource, name?: string) {
  busy.value = 'Opening…'
  try {
    onlineSource = null; onlineWritable = false
    if (typeof input === 'string') { const f = await fetchOnline({ kind: 'http', url: input }); await openBytes(new Uint8Array(await f.blob.arrayBuffer()), name || f.name) }
    else if (input instanceof Blob) { await openBytes(new Uint8Array(await input.arrayBuffer()), name || (input instanceof File ? input.name : 'Presentation.pptx')) }
    else { const f = await fetchOnline(input); onlineSource = input; onlineWritable = input.kind === 'webdav' || input.kind === 's3' || (input.kind === 'presigned' && !!input.putUrl); await openBytes(new Uint8Array(await f.blob.arrayBuffer()), name || f.name) }
  } catch (e) { emit('error', e as Error); notify((e as Error).message) }
  finally { busy.value = '' }
}
function newDocument() { pres.value = newPresentation(); pkg = null; slideParts = []; fileName.value = 'Presentation'; cur.value = 0; selected.value = null; undoStack.length = 0; setDirty(false); emit('ready') }
async function getBlob(): Promise<Blob> { return writePptx(pres.value, { pkg: pkg || undefined, slideParts, format: pres.value.format }) }
async function save() {
  busy.value = 'Saving…'
  try {
    const blob = await getBlob(); const name = fileName.value + '.' + pres.value.format
    if (onlineSource && onlineWritable) { await saveOnline(onlineSource as SaveTarget, blob, name); setDirty(false); notify('Saved.') }
    else { emit('save', blob, name); setDirty(false) }
  } catch (e) { emit('error', e as Error); notify((e as Error).message) }
  finally { busy.value = '' }
}
async function saveAs(name?: string): Promise<Blob> { const b = await getBlob(); downloadBlob(b, (name || fileName.value) + '.' + pres.value.format); setDirty(false); return b }
async function saveTo(target: SaveTarget, name?: string) { busy.value = 'Saving…'; try { await saveOnline(target, await getBlob(), (name || fileName.value) + '.' + pres.value.format); setDirty(false); notify('Saved.') } catch (e) { emit('error', e as Error); notify((e as Error).message) } finally { busy.value = '' } }

/* ---------- menus / toolbars ---------- */
const fileMenu: MenuDef = { label: t('File'), items: [
  { label: t('New'), icon: 'new', key: 'Ctrl+N', run: newDocument },
  { label: t('Open…'), icon: 'open', key: 'Ctrl+O', run: async () => { const f = await pickFile('.pptx,.pptm,.ppsx,.ppsm,.potx,.potm'); if (f[0]) open(f[0]) } },
  { label: t('Open from Online…'), icon: 'open', run: () => { dlg.value = 'online-open' } },
  { sep: true },
  { label: t('Save'), icon: 'save', key: 'Ctrl+S', run: save },
  { label: t('Save As…'), icon: 'save', run: () => saveAs() },
  { label: t('Save to Online…'), icon: 'save', run: () => { dlg.value = 'online-save' } },
] }
const editMenu: MenuDef = { label: t('Edit'), items: [
  { label: t('Undo'), icon: 'undo', key: 'Ctrl+Z', run: undo }, { label: t('Redo'), icon: 'redo', key: 'Ctrl+Y', run: redo },
  { sep: true }, { label: t('Delete Shape'), icon: 'close', key: 'Del', run: deleteShape },
] }
const viewMenu: MenuDef = { label: t('View'), items: [
  { label: t('Fit to Window'), icon: 'zoom', run: fitZoom },
  { label: t('Zoom In'), icon: 'zoom', run: () => { zoom.value = Math.min(200, zoom.value + 10) } },
  { label: t('Zoom Out'), icon: 'zoom', run: () => { zoom.value = Math.max(20, zoom.value - 10) } },
  { label: t('Slide Show'), icon: 'slideshow', key: 'F5', run: () => startShow(cur.value) },
] }
const insertMenu: MenuDef = { label: t('Insert'), items: [
  { label: t('New Slide'), icon: 'newSlide', key: 'Ctrl+M', run: addSlide },
  { label: t('Duplicate Slide'), icon: 'copy', run: dupSlide },
  { label: t('Text Box'), icon: 'textbox', run: () => { tool.value = 'textbox' } },
  { label: t('Picture…'), icon: 'picture', run: insertPicture },
  { label: t('Shape'), icon: 'rect', sub: PRESETS.map((p) => ({ label: p, run: () => { tool.value = p } })) },
] }
const formatMenu: MenuDef = { label: t('Format'), items: [
  { label: t('Bold'), icon: 'bold', run: () => fmtText((r) => { r.bold = !r.bold }) },
  { label: t('Italic'), icon: 'italic', run: () => fmtText((r) => { r.italic = !r.italic }) },
  { label: t('Shape Fill'), icon: 'fillColor', run: () => setFill('#4472c4') },
  { label: t('No Fill'), icon: 'fillColor', run: () => setFill(null) },
  { label: t('Shape Outline'), icon: 'lineColor', run: () => setLine('#2f5597') },
  { label: t('Slide Background…'), icon: 'design', run: () => { bgColor.value = slide.value.background || '#ffffff'; dlg.value = 'bg' } },
] }
const showMenu: MenuDef = { label: t('Slide Show'), items: [
  { label: t('From Beginning'), icon: 'slideshow', key: 'F5', run: () => startShow(0) },
  { label: t('From Current Slide'), icon: 'slideshow', run: () => startShow(cur.value) },
] }
const toolsMenu: MenuDef = { label: t('Tools'), items: [{ label: t('About VPptx'), icon: 'help', run: () => { dlg.value = 'about' } }] }
const helpMenu: MenuDef = { label: t('Help'), items: [{ label: t('About VPptx'), icon: 'help', run: () => { dlg.value = 'about' } }] }
const menus = computed(() => [fileMenu, editMenu, viewMenu, insertMenu, formatMenu, showMenu, toolsMenu, helpMenu])

const stdTools = computed<ToolbarItem[]>(() => [
  { kind: 'button', id: 'new', icon: 'new', run: newDocument },
  { kind: 'button', id: 'open', icon: 'open', run: async () => { const f = await pickFile('.pptx,.ppsx'); if (f[0]) open(f[0]) } },
  { kind: 'button', id: 'save', icon: 'save', run: save },
  { kind: 'sep' },
  { kind: 'button', id: 'newslide', icon: 'newSlide', tip: t('New Slide'), run: addSlide },
  { kind: 'button', id: 'undo', icon: 'undo', run: undo }, { kind: 'button', id: 'redo', icon: 'redo', run: redo },
  { kind: 'sep' },
  { kind: 'button', id: 'show', icon: 'slideshow', tip: t('Slide Show'), run: () => startShow(cur.value) },
])
const drawTools = computed<ToolbarItem[]>(() => [
  { kind: 'button', id: 'select', icon: 'select', tip: t('Select'), on: tool.value === 'select', run: () => { tool.value = 'select' } },
  { kind: 'button', id: 'textbox', icon: 'textbox', tip: t('Text Box'), on: tool.value === 'textbox', run: () => { tool.value = 'textbox' } },
  { kind: 'button', id: 'line', icon: 'line', tip: t('Line'), on: tool.value === 'line', run: () => { tool.value = 'line' } },
  { kind: 'button', id: 'rect', icon: 'rect', tip: t('Rectangle'), on: tool.value === 'rect', run: () => { tool.value = 'rect' } },
  { kind: 'button', id: 'oval', icon: 'oval', tip: t('Oval'), on: tool.value === 'ellipse', run: () => { tool.value = 'ellipse' } },
  { kind: 'button', id: 'arrow', icon: 'arrowTool', tip: t('Arrow'), on: tool.value === 'rightArrow', run: () => { tool.value = 'rightArrow' } },
  { kind: 'sep' },
  { kind: 'button', id: 'fill', icon: 'fillColor', tip: t('Fill'), run: () => setFill('#4472c4') },
  { kind: 'button', id: 'linecolor', icon: 'lineColor', tip: t('Line'), run: () => setLine('#2f5597') },
  { kind: 'button', id: 'pic', icon: 'picture', tip: t('Picture'), run: insertPicture },
])
const ribbonTabs = computed<RibbonTab[]>(() => [
  { id: 'home', label: t('Home'), groups: [ { label: t('Slides'), items: [ { kind: 'button', id: 'ns', icon: 'newSlide', label: t('New Slide'), run: addSlide } ] }, { label: t('Clipboard'), items: stdTools.value.slice(4, 7) } ] },
  { id: 'insert', label: t('Insert'), groups: [ { label: t('Illustrations'), items: [ { kind: 'button', id: 'pic', icon: 'picture', label: t('Picture'), run: insertPicture }, { kind: 'button', id: 'shape', icon: 'rect', label: t('Shapes'), menu: PRESETS.map((p) => ({ label: p, run: () => { tool.value = p } })) }, { kind: 'button', id: 'tb', icon: 'textbox', label: t('Text Box'), run: () => { tool.value = 'textbox' } } ] } ] },
  { id: 'show', label: t('Slide Show'), groups: [ { label: t('Start'), items: [ { kind: 'button', id: 'fs', icon: 'slideshow', label: t('From Start'), run: () => startShow(0) }, { kind: 'button', id: 'fc', icon: 'slideshow', label: t('Current'), run: () => startShow(cur.value) } ] } ] },
])
const statusCells = computed<StatusCell[]>(() => [
  { text: `Slide ${cur.value + 1} of ${pres.value.slides.length}`, grow: true },
  { text: zoom.value + '%', act: true, run: () => { zoom.value = zoom.value >= 200 ? 20 : zoom.value + 10 } },
])
function thumbHtml(s: Slide): string { return slideHtml(s, pres.value, 150) }

/* ---------- keyboard ---------- */
function onKey(e: KeyboardEvent) {
  if (show.value) { onShowKey(e); return }
  if (editingText.value) { if (e.key === 'Escape') { commitText(); nextTick(() => root.value?.focus()) } return }
  const mod = e.ctrlKey || e.metaKey
  const k = e.key.toLowerCase()
  if (mod) {
    const map: Record<string, () => void> = { s: save, o: async () => { const f = await pickFile('.pptx'); if (f[0]) open(f[0]) }, n: newDocument, z: () => (e.shiftKey ? redo() : undo()), y: redo, m: addSlide, d: dupSlide }
    if (map[k]) { e.preventDefault(); map[k]() }
    return
  }
  if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteShape() }
  else if (e.key === 'F5') { e.preventDefault(); startShow(0) }
  else if (e.key === 'Escape') { selected.value = null }
  else if (selected.value && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
    e.preventDefault()
    const s = selShape.value!; const d = emu(e.shiftKey ? 10 : 1)
    if (e.key === 'ArrowUp') s.y -= d
    if (e.key === 'ArrowDown') s.y += d
    if (e.key === 'ArrowLeft') s.x -= d
    if (e.key === 'ArrowRight') s.x += d
    setDirty(true)
  }
}
function onOnlineSubmit(source: OnlineSource | SaveTarget) { if (dlg.value === 'online-open') { dlg.value = ''; open(source as OnlineSource) } else { dlg.value = ''; saveTo(source as SaveTarget) } }

onMounted(async () => {
  window.addEventListener('mousemove', onMove)
  window.addEventListener('mouseup', onUp)
  fitRo = new ResizeObserver(() => fitZoom())
  if (editorWrap.value) fitRo.observe(editorWrap.value)
  if (props.src) await open(props.src, props.fileName); else newDocument()
  nextTick(fitZoom)
  root.value?.focus()
})
onBeforeUnmount(() => { window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp); fitRo?.disconnect() })
watch(() => props.src, (s) => { if (s) open(s, props.fileName) })

defineExpose<OfficeApi>({ open, newDocument, getBlob, save, saveAs, saveTo, get name() { return fileName.value }, get dirty() { return dirty.value } })
</script>

<template>
  <div ref="root" class="vo-app" :data-look="look" :data-layout="layout" tabindex="0" @keydown="onKey">
    <div class="app">
      <div class="cmdbars">
        <div class="menurow"><VoMenuBar :menus="menus" /></div>
        <div class="toolbars"><div class="tb-row"><VoToolbar :items="stdTools" /></div></div>
        <VoRibbon :tabs="ribbonTabs" @file="dlg = 'online-open'" />
      </div>
      <div class="main">
        <div class="leftpane">
          <div class="thumbs">
            <div v-for="(s, i) in pres.slides" :key="s.id" class="thumb" :class="{ on: i === cur }" @click="selectSlide(i)">
              <div class="thumb-in" v-html="thumbHtml(s)"></div>
              <div class="thumb-n">{{ i + 1 }}</div>
            </div>
          </div>
        </div>
        <div class="center">
          <div ref="editorWrap" class="editor-wrap">
            <div class="slidecanvas" :style="{ width: slideWpx * scale + 'px', height: slideHpx * scale + 'px', background: slide.background || '#fff' }" @mousedown="canvasDown">
              <div v-for="s in slide.shapes" :key="s.id" class="shape" :class="{ sel: selected === s.id }" :style="shapeStyle(s)"
                @mousedown="startMove(s, $event)" @dblclick="dblShape(s, $event)">
                <img v-if="s.kind === 'pic' && s.image" :src="s.image.url" class="shape-img" draggable="false">
                <svg v-else-if="s.raw == null" :width="px(s.w)" :height="px(s.h)" :viewBox="`0 0 ${px(s.w)} ${px(s.h)}`" class="shape-svg" v-html="shapeSvg(s)"></svg>
                <div v-else class="shape-raw"></div>
                <div v-if="s.kind !== 'pic' && s.raw == null" class="shape-text" :data-text="s.id"
                  :contenteditable="editingText === s.id" @input="setDirty(true)" @blur="commitText" v-html="textHtml(s)"></div>
                <template v-if="selected === s.id && editingText !== s.id">
                  <span v-for="h in ['nw','n','ne','e','se','s','sw','w']" :key="h" class="oh" :class="'oh-' + h" @mousedown="startResize(s, h, $event)"></span>
                </template>
              </div>
            </div>
          </div>
          <div class="notes">
            <textarea :value="slide.notes || ''" placeholder="Click to add notes" @input="slide.notes = ($event.target as HTMLTextAreaElement).value; setDirty(true)"></textarea>
          </div>
        </div>
      </div>
      <div class="drawrow"><VoToolbar :items="drawTools" :grip="false" /></div>
      <VoStatusBar :cells="statusCells" />
    </div>

    <div v-if="show" class="slideshow" @click="showNext" @keydown="onShowKey" tabindex="0">
      <div class="show-slide" v-html="showHtml"></div>
    </div>

    <VoDialog v-if="dlg === 'bg'" :title="t('Slide Background')" :width="300" @close="dlg = ''">
      <div class="field"><label>{{ t('Color') }}</label><input v-model="bgColor" type="color"></div>
      <template #foot><button class="btn primary" @click="applyBg">{{ t('Apply') }}</button><button class="btn" @click="dlg = ''">{{ t('Cancel') }}</button></template>
    </VoDialog>
    <VoDialog v-if="dlg === 'about'" :title="t('About VPptx')" :width="360" @close="dlg = ''">
      <div class="about"><div class="about-head"><span class="about-name">VPptx</span></div><p>{{ t('A Vue 3 presentation editor that opens and saves PowerPoint files.') }}</p></div>
      <template #foot><button class="btn primary" @click="dlg = ''">{{ t('OK') }}</button></template>
    </VoDialog>
    <VoOnlineDialog v-if="dlg === 'online-open'" mode="open" @close="dlg = ''" @submit="onOnlineSubmit" />
    <VoOnlineDialog v-if="dlg === 'online-save'" mode="save" :file-name="fileName + '.pptx'" @close="dlg = ''" @submit="onOnlineSubmit" />
    <VoBusy v-if="busy" :label="busy" />
    <VoToast v-if="toast" :message="toast" @done="toast = ''" />
  </div>
</template>

<style scoped>
.vo-app { outline: none; }
.main { flex: 1; min-height: 0; display: flex; }
.leftpane { flex: none; width: 180px; background: var(--surface); border-right: 1px solid var(--chrome-line); overflow: auto; }
.thumbs { display: flex; flex-direction: column; align-items: center; gap: 10px; padding: 8px; }
.thumb { border: 1px solid var(--grid-line); background: #fff; cursor: pointer; position: relative; }
.thumb.on { outline: 3px solid var(--thumb-on); }
.thumb-in { width: 150px; height: 84px; overflow: hidden; pointer-events: none; }
.thumb-n { font-size: 10px; text-align: center; }
.center { flex: 1; min-width: 0; display: flex; flex-direction: column; }
.editor-wrap { flex: 1; min-height: 0; overflow: auto; background: var(--canvas); display: flex; align-items: center; justify-content: center; padding: 20px; }
.slidecanvas { position: relative; box-shadow: 0 0 0 1px #c9c9c9, 2px 2px 10px rgba(0,0,0,.3); flex: none; }
.shape { position: absolute; }
.shape.sel { outline: 1px solid rgba(0,0,0,.6); }
.shape-svg { position: absolute; inset: 0; overflow: visible; pointer-events: none; }
.shape-img { width: 100%; height: 100%; object-fit: fill; pointer-events: none; }
.shape-raw { position: absolute; inset: 0; border: 1px dashed #999; background: #eee; }
.shape-text { position: absolute; inset: 4px; overflow: hidden; font-size: 18px; color: #000; outline: none; }
.shape-text[contenteditable="true"] { cursor: text; background: rgba(255,255,255,.6); }
.oh { position: absolute; width: 7px; height: 7px; background: #fff; border: 1px solid #000; margin: -3.5px 0 0 -3.5px; }
.oh-nw { left: 0; top: 0; cursor: nwse-resize; } .oh-n { left: 50%; top: 0; cursor: ns-resize; } .oh-ne { left: 100%; top: 0; cursor: nesw-resize; }
.oh-e { left: 100%; top: 50%; cursor: ew-resize; } .oh-se { left: 100%; top: 100%; cursor: nwse-resize; } .oh-s { left: 50%; top: 100%; cursor: ns-resize; }
.oh-sw { left: 0; top: 100%; cursor: nesw-resize; } .oh-w { left: 0; top: 50%; cursor: ew-resize; }
.notes { flex: none; height: 70px; border-top: 1px solid var(--chrome-line); background: var(--surface); }
.notes textarea { width: 100%; height: 100%; border: 0; resize: none; padding: 4px 8px; font: 11px var(--ui-font); outline: none; }
.drawrow { flex: none; padding: 1px 2px; background: linear-gradient(90deg, var(--dock-a), var(--dock-b)); border-top: 1px solid var(--chrome-line); }
.slideshow { position: fixed; inset: 0; background: #000; z-index: 4000; display: flex; align-items: center; justify-content: center; outline: none; }
.show-slide { background: #fff; }
</style>