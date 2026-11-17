<script setup lang="ts">
import { ref, computed, onMounted, watch, nextTick } from 'vue'
import {
  VoMenuBar, VoToolbar, VoStatusBar, VoDialog, VoToast, VoBusy, VoRibbon, VoOnlineDialog,
  useLook, t, icon, pickFile, downloadBlob, fetchOnline, saveOnline,
  type MenuDef, type ToolbarItem, type RibbonTab, type StatusCell, type Look,
  type OnlineSource, type SaveTarget, type OfficeApi,
} from '@voffice/shared'
import { type Doc, type Block, type Para, type Table, newDoc, cloneDoc, emptyPara, uid, isPara, wordCount, paraText } from './engine/model'
import { readDocx } from './engine/docx-read'
import { writeDocx } from './engine/docx-write'
import { renderBlock, syncBlock } from './engine/render'
import { paginate, resolveMarkers, getCaret, restoreCaret, type Caret } from './engine/layout'

const props = defineProps<{ src?: string | File | Blob | OnlineSource; fileName?: string; look?: Look; readonly?: boolean }>()
const emit = defineEmits<{
  ready: []; change: [dirty: boolean]; save: [blob: Blob, name: string]; error: [err: Error]
  openRequest: []; saveRequest: []
}>()

const { look, layout } = useLook(props.look)
watch(() => props.look, (l) => { if (l) look.value = l })

const root = ref<HTMLElement | null>(null)
const scroller = ref<HTMLElement | null>(null)
const measure = ref<HTMLElement | null>(null)
const doc = ref<Doc>(newDoc())
const pages = ref<Block[][]>([[]])
const view = ref<'print' | 'web'>('print')
const zoom = ref(100)
const fileName = ref(props.fileName || 'Document')
const dirty = ref(false)
const busy = ref('')
const toast = ref('')
const version = ref(0)

let pkg: import('@voffice/shared').Package | null = null
let part = 'word/document.xml'
let onlineSource: OnlineSource | null = null
let onlineWritable = false
const undoStack: Doc[] = []
const redoStack: Doc[] = []
let lastSnap = 0

const PX = 96 / 1440
const contentW = computed(() => Math.round((doc.value.section.pageW - doc.value.section.margins.left - doc.value.section.margins.right) * PX))
const contentH = computed(() => Math.round((doc.value.section.pageH - doc.value.section.margins.top - doc.value.section.margins.bottom) * PX))
const pageW = computed(() => Math.round(doc.value.section.pageW * PX))
const pageH = computed(() => Math.round(doc.value.section.pageH * PX))
const words = computed(() => { version.value; return wordCount(doc.value) })

function notify(msg: string) { toast.value = msg; setTimeout(() => { if (toast.value === msg) toast.value = '' }, 5000) }
function setDirty(d: boolean) { if (dirty.value !== d) { dirty.value = d; emit('change', d) } }

/* ---------- rendering ---------- */
async function renderPages(keepCaret?: Caret | null) {
  if (view.value === 'print' && measure.value) {
    pages.value = paginate(doc.value, contentW.value, contentH.value, measure.value)
  } else {
    pages.value = [doc.value.blocks]
  }
  await nextTick()
  const bodies = root.value?.querySelectorAll<HTMLElement>('.pg-body, .flow-body') || []
  bodies.forEach((b) => resolveMarkers(b, doc.value))
  if (keepCaret) {
    const target = root.value?.querySelector<HTMLElement>('.pg-body, .flow-body')
    if (target) restoreCaret(target, keepCaret)
  }
}

function pageHtml(blocks: Block[]): string { return blocks.map((b) => renderBlock(b, doc.value)).join('') }

/* ---------- model sync from the editing DOM ---------- */
function syncFromDOM() {
  const bodies = Array.from(root.value?.querySelectorAll<HTMLElement>('.pg-body, .flow-body') || [])
  const blocks: Block[] = []
  const prevById = new Map(doc.value.blocks.map((b) => [b.id, b]))
  for (const body of bodies) {
    for (const el of Array.from(body.children)) {
      const hel = el as HTMLElement
      if (!hel.hasAttribute?.('data-id') && hel.tagName !== 'TABLE' && !hel.classList.contains('p')) continue
      const id = hel.getAttribute('data-id') || undefined
      blocks.push(syncBlock(hel, doc.value, id ? prevById.get(id) : undefined))
    }
  }
  if (blocks.length) doc.value.blocks = blocks
  setDirty(true)
  version.value++
}

function snapshot(force = false) {
  const now = Date.now()
  if (!force && now - lastSnap < 600) return
  lastSnap = now
  undoStack.push(cloneDoc(doc.value))
  if (undoStack.length > 60) undoStack.shift()
  redoStack.length = 0
}

function onInput() {
  snapshot()
  syncFromDOM()
  // re-flow only when a page overflows, preserving the caret
  const bodies = Array.from(root.value?.querySelectorAll<HTMLElement>('.pg-body') || [])
  const overflow = bodies.some((b) => b.scrollHeight > contentH.value + 2)
  if (overflow && view.value === 'print') {
    const caret = getCaret(bodies[0])
    renderPages(caret)
  }
}

function undo() {
  const s = undoStack.pop()
  if (!s) return
  redoStack.push(cloneDoc(doc.value))
  doc.value = s
  renderPages()
  setDirty(true)
}
function redo() {
  const s = redoStack.pop()
  if (!s) return
  undoStack.push(cloneDoc(doc.value))
  doc.value = s
  renderPages()
  setDirty(true)
}

/* ---------- editing commands ---------- */
function activeBody(): HTMLElement | null {
  const sel = window.getSelection()
  let n: Node | null = sel?.anchorNode || null
  while (n && n !== root.value) {
    if (n.nodeType === 1 && ((n as HTMLElement).classList.contains('pg-body') || (n as HTMLElement).classList.contains('flow-body'))) return n as HTMLElement
    n = n.parentNode
  }
  return root.value?.querySelector<HTMLElement>('.pg-body, .flow-body') || null
}
function exec(cmd: string, val?: string) {
  if (props.readonly) return
  const body = activeBody()
  body?.focus()
  snapshot(true)
  document.execCommand(cmd, false, val)
  syncFromDOM()
}
function insertBlock(block: Block) {
  snapshot(true)
  const caret = getCaret(activeBody() || root.value!)
  const idx = caret ? doc.value.blocks.findIndex((b) => b.id === caret.id) : doc.value.blocks.length - 1
  doc.value.blocks.splice(idx + 1, 0, block)
  setDirty(true)
  renderPages({ id: block.id, offset: 0 })
}
function insertText(text: string) { exec('insertText', text) }

/* ---------- open / save ---------- */
async function openBytes(bytes: Uint8Array, name: string) {
  const res = await readDocx(bytes)
  doc.value = res.doc; pkg = res.pkg; part = res.part
  fileName.value = name.replace(/\.(docx|docm|dotx|dotm)$/i, '')
  undoStack.length = 0; redoStack.length = 0
  setDirty(false)
  await renderPages()
  emit('ready')
}
async function open(input: File | Blob | string | OnlineSource, name?: string) {
  busy.value = 'Opening…'
  try {
    onlineSource = null; onlineWritable = false
    if (typeof input === 'string') {
      const f = await fetchOnline({ kind: 'http', url: input })
      await openBytes(new Uint8Array(await f.blob.arrayBuffer()), name || f.name)
    } else if (input instanceof Blob) {
      const n = name || (input instanceof File ? input.name : 'Document.docx')
      await openBytes(new Uint8Array(await input.arrayBuffer()), n)
    } else {
      const f = await fetchOnline(input)
      onlineSource = input
      onlineWritable = input.kind === 'webdav' || input.kind === 's3' || (input.kind === 'presigned' && !!input.putUrl)
      await openBytes(new Uint8Array(await f.blob.arrayBuffer()), name || f.name)
    }
  } catch (e) { emit('error', e as Error); notify((e as Error).message) }
  finally { busy.value = '' }
}
function newDocument() {
  doc.value = newDoc(); pkg = null; part = 'word/document.xml'
  onlineSource = null; onlineWritable = false
  fileName.value = 'Document'; undoStack.length = 0; redoStack.length = 0
  setDirty(false); renderPages(); emit('ready')
}
async function getBlob(): Promise<Blob> { return writeDocx(doc.value, { pkg: pkg || undefined, part, format: doc.value.format }) }
async function save() {
  busy.value = 'Saving…'
  try {
    const blob = await getBlob()
    const name = fileName.value + '.' + (doc.value.format || 'docx')
    if (onlineSource && onlineWritable) {
      await saveOnline(onlineSource as SaveTarget, blob, name)
      setDirty(false); notify('Saved to ' + (onlineSource.kind === 's3' ? onlineSource.bucket : 'the online location') + '.')
    } else {
      emit('save', blob, name)
      setDirty(false)
    }
  } catch (e) { emit('error', e as Error); notify((e as Error).message) }
  finally { busy.value = '' }
}
async function saveAs(name?: string): Promise<Blob> {
  const blob = await getBlob()
  const n = (name || fileName.value) + '.' + (doc.value.format || 'docx')
  downloadBlob(blob, n)
  setDirty(false)
  return blob
}
async function saveTo(target: SaveTarget, name?: string) {
  busy.value = 'Saving…'
  try {
    const blob = await getBlob()
    await saveOnline(target, blob, (name || fileName.value) + '.' + (doc.value.format || 'docx'))
    setDirty(false); notify('Saved.')
  } catch (e) { emit('error', e as Error); notify((e as Error).message) }
  finally { busy.value = '' }
}

/* ---------- dialogs ---------- */
const dlg = ref<'' | 'find' | 'table' | 'link' | 'symbol' | 'about' | 'online-open' | 'online-save'>('')
const findText = ref(''); const replaceText = ref('')
const tblRows = ref(3); const tblCols = ref(3)
const linkUrl = ref('https://'); const linkText = ref('')
const symbols = '©®™°±×÷€£¥§¶†‡•…‰′″‹›«»←↑→↓↔∞≈≠≤≥√∑∏∫αβγδεπΩ∆'.split('')

function findNext() {
  if (!findText.value) return
  // @ts-ignore window.find is non-standard but works in Chromium
  const found = window.find(findText.value, false, false, true, false, true, false)
  if (!found) notify('No more matches for “' + findText.value + '”.')
}
function replaceOne() {
  const sel = window.getSelection()
  if (sel && sel.toString() === findText.value) { exec('insertText', replaceText.value) }
  findNext()
}
function replaceAll() {
  if (!findText.value) return
  snapshot(true)
  let n = 0
  for (const b of doc.value.blocks) {
    if (!isPara(b)) continue
    for (const r of b.runs) {
      if (r.text && r.text.includes(findText.value)) { r.text = r.text.split(findText.value).join(replaceText.value); n++ }
    }
  }
  setDirty(true); renderPages(); notify(`Replaced ${n} occurrence(s).`)
}
function doInsertTable() {
  const widths = Array.from({ length: tblCols.value }, () => Math.round(contentW.value / tblCols.value))
  const rows = Array.from({ length: tblRows.value }, () =>
    Array.from({ length: tblCols.value }, () => ({ paras: [emptyPara()], width: widths[0] })))
  const tbl: Table = { kind: 'tbl', id: uid('t'), rows, widths }
  insertBlock(tbl); dlg.value = ''
}
function doInsertLink() {
  if (linkText.value) {
    insertBlock({ ...emptyPara(), runs: [{ text: linkText.value, hyperlink: { rId: '', url: linkUrl.value }, rPr: { color: '#0563c1', underline: true } }] })
  } else exec('createLink', linkUrl.value)
  dlg.value = ''
}
async function insertPicture() {
  const files = await pickFile('image/*')
  if (!files[0]) return
  const url = await new Promise<string>((res) => { const fr = new FileReader(); fr.onload = () => res(fr.result as string); fr.readAsDataURL(files[0]) })
  const img = new Image()
  await new Promise((res) => { img.onload = res; img.src = url })
  const w = Math.min(img.width, contentW.value)
  const h = Math.round(img.height * (w / img.width))
  insertBlock({ ...emptyPara(), runs: [{ image: { url, part: '', rId: '', w, h } }] })
}
function insertPageBreak() { insertBlock({ ...emptyPara(), pPr: { pageBreakBefore: true } }) }

/* ---------- menus / toolbars / ribbon ---------- */
const fileMenu: MenuDef = {
  label: t('File'), items: [
    { label: t('New'), icon: 'new', key: 'Ctrl+N', run: () => newDocument() },
    { label: t('Open…'), icon: 'open', key: 'Ctrl+O', run: async () => { const f = await pickFile('.docx,.docm,.dotx,.dotm'); if (f[0]) open(f[0]) } },
    { label: t('Open from Online…'), icon: 'open', run: () => { dlg.value = 'online-open' } },
    { sep: true },
    { label: t('Save'), icon: 'save', key: 'Ctrl+S', run: () => save() },
    { label: t('Save As…'), icon: 'save', run: () => saveAs() },
    { label: t('Save to Online…'), icon: 'save', run: () => { dlg.value = 'online-save' } },
    { sep: true },
    { label: t('Print…'), icon: 'print', key: 'Ctrl+P', run: () => window.print() },
  ],
}
const editMenu: MenuDef = {
  label: t('Edit'), items: [
    { label: t('Undo'), icon: 'undo', key: 'Ctrl+Z', run: undo },
    { label: t('Redo'), icon: 'redo', key: 'Ctrl+Y', run: redo },
    { sep: true },
    { label: t('Cut'), icon: 'cut', key: 'Ctrl+X', run: () => exec('cut') },
    { label: t('Copy'), icon: 'copy', key: 'Ctrl+C', run: () => exec('copy') },
    { label: t('Paste'), icon: 'paste', key: 'Ctrl+V', run: () => exec('paste') },
    { sep: true },
    { label: t('Find…'), icon: 'find', key: 'Ctrl+F', run: () => { dlg.value = 'find' } },
    { label: t('Replace…'), icon: 'replace', run: () => { dlg.value = 'find' } },
    { label: t('Select All'), key: 'Ctrl+A', run: () => exec('selectAll') },
  ],
}
const viewMenu: MenuDef = {
  label: t('View'), items: [
    { label: t('Print Layout'), icon: 'preview', checked: view.value === 'print', radio: true, run: () => { view.value = 'print'; renderPages() } },
    { label: t('Web Layout'), icon: 'normalView', checked: view.value === 'web', radio: true, run: () => { view.value = 'web'; renderPages() } },
    { sep: true },
    { label: t('Zoom In'), icon: 'zoom', run: () => { zoom.value = Math.min(200, zoom.value + 10) } },
    { label: t('Zoom Out'), icon: 'zoom', run: () => { zoom.value = Math.max(50, zoom.value - 10) } },
  ],
}
const insertMenu: MenuDef = {
  label: t('Insert'), items: [
    { label: t('Page Break'), icon: 'new', run: insertPageBreak },
    { label: t('Table…'), icon: 'table', run: () => { dlg.value = 'table' } },
    { label: t('Picture…'), icon: 'picture', run: insertPicture },
    { label: t('Hyperlink…'), icon: 'hyperlink', key: 'Ctrl+K', run: () => { dlg.value = 'link' } },
    { label: t('Symbol…'), icon: 'symbol', run: () => { dlg.value = 'symbol' } },
    { label: t('Date and Time'), icon: 'date', run: () => insertText(new Date().toLocaleDateString()) },
  ],
}
const formatMenu: MenuDef = {
  label: t('Format'), items: [
    { label: t('Bold'), icon: 'bold', key: 'Ctrl+B', run: () => exec('bold') },
    { label: t('Italic'), icon: 'italic', key: 'Ctrl+I', run: () => exec('italic') },
    { label: t('Underline'), icon: 'underline', key: 'Ctrl+U', run: () => exec('underline') },
    { sep: true },
    { label: t('Align Left'), icon: 'alignLeft', run: () => exec('justifyLeft') },
    { label: t('Center'), icon: 'alignCenter', run: () => exec('justifyCenter') },
    { label: t('Align Right'), icon: 'alignRight', run: () => exec('justifyRight') },
    { label: t('Justify'), icon: 'justify', run: () => exec('justifyFull') },
    { sep: true },
    { label: t('Bullets'), icon: 'bullets', run: () => exec('insertUnorderedList') },
    { label: t('Numbering'), icon: 'numbering', run: () => exec('insertOrderedList') },
    { label: t('Increase Indent'), icon: 'incIndent', run: () => exec('indent') },
    { label: t('Decrease Indent'), icon: 'decIndent', run: () => exec('outdent') },
  ],
}
const toolsMenu: MenuDef = {
  label: t('Tools'), items: [
    { label: t('Word Count'), icon: 'properties', run: () => notify(`${words.value} words`) },
    { label: t('About VWord'), icon: 'help', run: () => { dlg.value = 'about' } },
  ],
}
const helpMenu: MenuDef = { label: t('Help'), items: [{ label: t('About VWord'), icon: 'help', run: () => { dlg.value = 'about' } }] }
const menus = computed<MenuDef[]>(() => [fileMenu, editMenu, viewMenu, insertMenu, formatMenu, toolsMenu, helpMenu])

const fonts = ['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Georgia', 'Verdana', 'Tahoma', 'Cambria', 'Consolas']
const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]
const curFont = ref('Calibri'); const curSize = ref(11)
const stdTools = computed<ToolbarItem[]>(() => [
  { kind: 'button', id: 'new', icon: 'new', tip: t('New'), run: () => newDocument() },
  { kind: 'button', id: 'open', icon: 'open', tip: t('Open'), run: async () => { const f = await pickFile('.docx,.docm,.dotx,.dotm'); if (f[0]) open(f[0]) } },
  { kind: 'button', id: 'save', icon: 'save', tip: t('Save'), run: () => save() },
  { kind: 'button', id: 'print', icon: 'print', tip: t('Print'), run: () => window.print() },
  { kind: 'sep' },
  { kind: 'button', id: 'cut', icon: 'cut', tip: t('Cut'), run: () => exec('cut') },
  { kind: 'button', id: 'copy', icon: 'copy', tip: t('Copy'), run: () => exec('copy') },
  { kind: 'button', id: 'paste', icon: 'paste', tip: t('Paste'), run: () => exec('paste') },
  { kind: 'sep' },
  { kind: 'button', id: 'undo', icon: 'undo', tip: t('Undo'), run: undo },
  { kind: 'button', id: 'redo', icon: 'redo', tip: t('Redo'), run: redo },
  { kind: 'sep' },
  { kind: 'button', id: 'find', icon: 'find', tip: t('Find'), run: () => { dlg.value = 'find' } },
])
const fmtTools = computed<ToolbarItem[]>(() => [
  { kind: 'combo', id: 'font', value: curFont.value, width: 130, editable: true, fontList: true, items: fonts.map((f) => ({ label: f })), onChange: (v) => { curFont.value = v; exec('fontName', v) } },
  { kind: 'combo', id: 'size', value: String(curSize.value), width: 48, editable: true, items: sizes.map((s) => ({ label: String(s) })), onChange: (v) => { curSize.value = +v; exec('fontSize', '4'); applySize(+v) } },
  { kind: 'sep' },
  { kind: 'button', id: 'bold', icon: 'bold', tip: t('Bold'), run: () => exec('bold') },
  { kind: 'button', id: 'italic', icon: 'italic', tip: t('Italic'), run: () => exec('italic') },
  { kind: 'button', id: 'underline', icon: 'underline', tip: t('Underline'), run: () => exec('underline') },
  { kind: 'button', id: 'color', icon: 'fontColor', tip: t('Font Color'), run: () => exec('foreColor', '#c00000') },
  { kind: 'sep' },
  { kind: 'button', id: 'left', icon: 'alignLeft', tip: t('Align Left'), run: () => exec('justifyLeft') },
  { kind: 'button', id: 'center', icon: 'alignCenter', tip: t('Center'), run: () => exec('justifyCenter') },
  { kind: 'button', id: 'right', icon: 'alignRight', tip: t('Align Right'), run: () => exec('justifyRight') },
  { kind: 'button', id: 'justify', icon: 'justify', tip: t('Justify'), run: () => exec('justifyFull') },
  { kind: 'sep' },
  { kind: 'button', id: 'bullets', icon: 'bullets', tip: t('Bullets'), run: () => exec('insertUnorderedList') },
  { kind: 'button', id: 'numbering', icon: 'numbering', tip: t('Numbering'), run: () => exec('insertOrderedList') },
  { kind: 'button', id: 'decindent', icon: 'decIndent', tip: t('Decrease Indent'), run: () => exec('outdent') },
  { kind: 'button', id: 'incindent', icon: 'incIndent', tip: t('Increase Indent'), run: () => exec('indent') },
])
function applySize(pt: number) {
  // execCommand fontSize uses 1-7; set a precise size on the selection via a span
  const sel = window.getSelection()
  if (!sel || sel.isCollapsed) return
  const range = sel.getRangeAt(0)
  const span = document.createElement('span')
  span.style.fontSize = pt + 'pt'
  try { range.surroundContents(span) } catch { exec('fontSize', '4') }
  syncFromDOM()
}
const ribbonTabs = computed<RibbonTab[]>(() => [
  { id: 'home', label: t('Home'), groups: [
    { label: t('Clipboard'), items: stdTools.value.slice(5, 11) },
    { label: t('Font'), items: fmtTools.value.slice(0, 7) },
    { label: t('Paragraph'), items: fmtTools.value.slice(8) },
  ] },
  { id: 'insert', label: t('Insert'), groups: [
    { label: t('Insert'), items: [
      { kind: 'button', id: 'tbl', icon: 'table', label: t('Table'), run: () => { dlg.value = 'table' } },
      { kind: 'button', id: 'pic', icon: 'picture', label: t('Picture'), run: insertPicture },
      { kind: 'button', id: 'lnk', icon: 'hyperlink', label: t('Link'), run: () => { dlg.value = 'link' } },
      { kind: 'button', id: 'sym', icon: 'symbol', label: t('Symbol'), run: () => { dlg.value = 'symbol' } },
      { kind: 'button', id: 'brk', icon: 'new', label: t('Page Break'), run: insertPageBreak },
    ] },
  ] },
  { id: 'view', label: t('View'), groups: [
    { label: t('Views'), items: [
      { kind: 'button', id: 'print', icon: 'preview', label: t('Print Layout'), on: view.value === 'print', run: () => { view.value = 'print'; renderPages() } },
      { kind: 'button', id: 'web', icon: 'normalView', label: t('Web Layout'), on: view.value === 'web', run: () => { view.value = 'web'; renderPages() } },
    ] },
  ] },
])
const statusCells = computed<StatusCell[]>(() => [
  { text: `${words.value} words`, act: true, run: () => notify(`${words.value} words`) },
  { text: t(view.value === 'print' ? 'Print Layout' : 'Web Layout'), grow: true },
  { text: zoom.value + '%', act: true, run: () => { zoom.value = zoom.value >= 200 ? 50 : zoom.value + 10 } },
])

/* ---------- keyboard (scoped to the component root) ---------- */
function onKey(e: KeyboardEvent) {
  const mod = e.ctrlKey || e.metaKey
  if (!mod) return
  const k = e.key.toLowerCase()
  const map: Record<string, () => void> = {
    s: () => save(), o: async () => { const f = await pickFile('.docx,.docm,.dotx,.dotm'); if (f[0]) open(f[0]) },
    n: () => newDocument(), p: () => window.print(), z: () => (e.shiftKey ? redo() : undo()), y: () => redo(),
    b: () => exec('bold'), i: () => exec('italic'), u: () => exec('underline'),
    f: () => { dlg.value = 'find' }, k: () => { dlg.value = 'link' },
  }
  if (map[k]) { e.preventDefault(); map[k]() }
}

function onOnlineSubmit(source: OnlineSource | SaveTarget) {
  if (dlg.value === 'online-open') { dlg.value = ''; open(source as OnlineSource) }
  else { dlg.value = ''; saveTo(source as SaveTarget) }
}

onMounted(async () => {
  if (props.src) await open(props.src, props.fileName)
  else { newDocument() }
  root.value?.focus()
})
watch(() => props.src, (s) => { if (s) open(s, props.fileName) })

defineExpose<OfficeApi>({
  open, newDocument, getBlob, save, saveAs, saveTo,
  get name() { return fileName.value }, get dirty() { return dirty.value },
})
</script>

<template>
  <div ref="root" class="vo-app" :data-look="look" :data-layout="layout" tabindex="0" @keydown="onKey">
    <div class="app" :class="{ readonly }">
      <div class="cmdbars">
        <div class="menurow"><VoMenuBar :menus="menus" /></div>
        <div class="toolbars">
          <div class="tb-row"><VoToolbar :items="stdTools" /></div>
          <div class="tb-row"><VoToolbar :items="fmtTools" /></div>
        </div>
        <VoRibbon :tabs="ribbonTabs" @file="dlg = 'online-open'" />
      </div>
      <div class="main">
        <div class="center">
          <div class="work">
            <div ref="scroller" class="scroller" tabindex="-1">
              <div class="sizer" :style="{ width: (view === 'print' ? pageW * zoom / 100 + 48 : contentW + 48) + 'px' }">
                <div v-if="view === 'print'" class="pages" :style="{ transform: `scale(${zoom / 100})`, transformOrigin: 'top center' }">
                  <div v-for="(pg, i) in pages" :key="i" class="pg" :style="{ width: pageW + 'px', height: pageH + 'px', padding: `${doc.section.margins.top * PX}px ${doc.section.margins.right * PX}px ${doc.section.margins.bottom * PX}px ${doc.section.margins.left * PX}px` }">
                    <div class="pg-body" :contenteditable="!readonly" spellcheck="false" :style="{ width: contentW + 'px', minHeight: contentH + 'px' }" @input="onInput" v-html="pageHtml(pg)"></div>
                  </div>
                </div>
                <div v-else class="flow" :style="{ transform: `scale(${zoom / 100})`, transformOrigin: 'top center' }">
                  <div class="flow-body" :contenteditable="!readonly" spellcheck="false" :style="{ width: contentW + 'px' }" @input="onInput" v-html="pageHtml(doc.blocks)"></div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
      <VoStatusBar :cells="statusCells" />
    </div>

    <div ref="measure" class="measure" aria-hidden="true"></div>

    <VoDialog v-if="dlg === 'find'" :title="t('Find and Replace')" :width="420" @close="dlg = ''">
      <div class="field wide"><label>{{ t('Find what') }}</label><input v-model="findText" type="text" style="flex:1" @keydown.enter="findNext"></div>
      <div class="field wide"><label>{{ t('Replace with') }}</label><input v-model="replaceText" type="text" style="flex:1"></div>
      <template #foot>
        <button class="btn" @click="findNext">{{ t('Find Next') }}</button>
        <button class="btn" @click="replaceOne">{{ t('Replace') }}</button>
        <button class="btn" @click="replaceAll">{{ t('Replace All') }}</button>
        <button class="btn" @click="dlg = ''">{{ t('Close') }}</button>
      </template>
    </VoDialog>

    <VoDialog v-if="dlg === 'table'" :title="t('Insert Table')" :width="300" @close="dlg = ''">
      <div class="field"><label>{{ t('Columns') }}</label><input v-model.number="tblCols" type="number" min="1" max="12" style="width:70px"></div>
      <div class="field"><label>{{ t('Rows') }}</label><input v-model.number="tblRows" type="number" min="1" max="40" style="width:70px"></div>
      <template #foot>
        <button class="btn primary" @click="doInsertTable">{{ t('Insert') }}</button>
        <button class="btn" @click="dlg = ''">{{ t('Cancel') }}</button>
      </template>
    </VoDialog>

    <VoDialog v-if="dlg === 'link'" :title="t('Insert Hyperlink')" :width="420" @close="dlg = ''">
      <div class="field wide"><label>{{ t('Address') }}</label><input v-model="linkUrl" type="text" style="flex:1"></div>
      <div class="field wide"><label>{{ t('Text (optional)') }}</label><input v-model="linkText" type="text" style="flex:1"></div>
      <template #foot>
        <button class="btn primary" @click="doInsertLink">{{ t('Insert') }}</button>
        <button class="btn" @click="dlg = ''">{{ t('Cancel') }}</button>
      </template>
    </VoDialog>

    <VoDialog v-if="dlg === 'symbol'" :title="t('Symbol')" :width="460" @close="dlg = ''">
      <div class="sym-grid">
        <button v-for="s in symbols" :key="s" @click="insertText(s); dlg = ''">{{ s }}</button>
      </div>
      <template #foot><button class="btn" @click="dlg = ''">{{ t('Close') }}</button></template>
    </VoDialog>

    <VoDialog v-if="dlg === 'about'" :title="t('About VWord')" :width="360" @close="dlg = ''">
      <div class="about">
        <div class="about-head"><span class="about-name">VWord</span></div>
        <p>{{ t('A Vue 3 word processor that opens and saves Word documents.') }}</p>
        <p class="about-marks">{{ t('Part of the VibeOffice Vue component suite.') }}</p>
      </div>
      <template #foot><button class="btn primary" @click="dlg = ''">{{ t('OK') }}</button></template>
    </VoDialog>

    <VoOnlineDialog v-if="dlg === 'online-open'" mode="open" @close="dlg = ''" @submit="onOnlineSubmit" />
    <VoOnlineDialog v-if="dlg === 'online-save'" mode="save" :file-name="fileName + '.docx'" @close="dlg = ''" @submit="onOnlineSubmit" />

    <VoBusy v-if="busy" :label="busy" />
    <VoToast v-if="toast" :message="toast" @done="toast = ''" />
  </div>
</template>

<style scoped>
.measure { position: absolute; left: -99999px; top: 0; visibility: hidden; }
.scroller { overflow: auto; height: 100%; background: var(--canvas); padding: 24px 0; }
.sizer { margin: 0 auto; }
.pages { display: flex; flex-direction: column; align-items: center; gap: 24px; }
.pg { background: #fff; box-shadow: 0 0 0 1px #c9c9c9, 2px 2px 8px rgba(0,0,0,.25); overflow: hidden; }
.pg-body { outline: none; }
.pg-body :deep(.p) { position: relative; min-height: 1em; }
.pg-body :deep(.lm) { color: inherit; }
.flow { background: #fff; padding: 24px; box-shadow: 0 0 0 1px #c9c9c9; }
.flow-body { outline: none; }
.flow-body :deep(.p) { position: relative; min-height: 1em; }
.vo-app { outline: none; }
</style>