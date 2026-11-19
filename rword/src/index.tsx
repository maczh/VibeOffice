import { forwardRef, useCallback, useEffect, useImperativeHandle, useMemo, useRef, useState, type ReactNode } from 'react'
import { ErrorBoundary, downloadBlob, fetchOnline, saveOnline, useDirtyGuard, type OfficeApi, type OfficeProps, type OpenInput, type SaveTarget } from '@voffice/react-shared'
import type { Package } from '@voffice/react-shared'
import {
  docText, emptyPara, htmlToText, imageRun, insertTable, newDoc, readMarkdown, renderBlock, resolveMarkers,
  syncBlock, textToHtml, writeMarkdown, wordCount, type Doc,
} from './engine'
import { readDocx, writeDocx } from './engine'
import { replaceAll } from './find'
import { buildWelcomeDoc } from './welcome'
import { documentCommands, fallbackLocation, locate, type CommandEnv, type CommandFn } from './services/documentCommands'
import { recordInsertion, reviewCommands } from './services/reviewCommands'
import { auxiliaryCommands, loadMergeSource } from './services/fieldCommands'
import { getIcon } from './icons'
import { QDialog, QIcon, QMenuBar, QTaskPane, QToolbar, QToolButton, QViewBar, iconFor, labelFor, type QItem, type QMenu, type QTool } from './quireChrome'
import { t } from './i18n'

/* Classic chrome stylesheets — RWord renders the same DOM as public/quire. */
import '../../public/common/looks.css'
import '../../public/common/ui.css'
import '../../public/fonts/fonts.css'
import './quire.css'

export { quireScriptMigration } from './quire-manifest'
export { TaskPane, type TaskPaneProps } from './components/TaskPane'
export { EditorSurface, type EditorSurfaceProps } from './components/EditorSurface'
export { Ruler, type RulerProps } from './components/Ruler'
export { MailMergePanel, type MailMergePanelProps } from './components/MailMergePanel'
export { ReviewPanel, type ReviewPanelProps } from './components/ReviewPanel'
export { RWordWorkspace, type RWordWorkspaceProps, type WorkspacePanel } from './components/RWordWorkspace'
export { OutlinePanel, type OutlinePanelProps } from './components/OutlinePanel'
export { FormattingPanel, type FormattingPanelProps } from './components/FormattingPanel'
export { CommandSurface, type CommandSurfaceProps } from './components/CommandSurface'

export type { OfficeApi, OfficeProps, OpenInput, SaveTarget }
export interface RWordProps extends OfficeProps { onReady?: () => void; onChange?: (dirty: boolean) => void; onSave?: (blob: Blob, name: string) => void; onError?: (error: Error) => void }

type Pane = 'getting-started' | 'help' | 'new' | 'search' | 'styles' | 'reveal' | 'mailmerge' | 'menu'
type Dialog = 'open' | 'save' | 'about' | 'find' | 'table' | 'wordcount' | 'font' | 'paragraph' | 'pagesetup' | 'zoom'
  | 'changeCase' | 'insertCells' | 'insertSymbol' | 'hyperlink' | 'bookmark' | 'field' | 'mergeSource'
  | 'tableProps' | 'borders' | 'formula' | 'comment' | 'pvGrid' | null

/* The labels Quire's command registry gives each id, before L.t translates them. */
const COMMAND_LABELS: Record<string, string> = {
  new: '&New...', open: '&Open...', close: '&Close', save: '&Save', saveAs: 'Save &As...',
  saveWeb: 'Save as Web Pa&ge...', fileSearch: 'File Searc&h...', printPreview: 'Print Pre&view',
  print: '&Print...', exportPdf: 'Export as &PDF', properties: 'Propert&ies', exitApp: 'E&xit',
  undo: '&Undo', redo: '&Redo', cut: 'Cu&t', copy: '&Copy', paste: '&Paste',
  clearFormats: 'Clear &Formats', clearContents: 'Clear &Contents', selectAll: 'Select &All',
  find: '&Find...', replace: 'R&eplace...', goTo: '&Go To...',
  viewNormal: '&Normal', viewWeb: '&Web Layout', viewPrint: '&Print Layout',
  viewReading: 'R&eading Layout', viewOutline: '&Outline', taskPane: 'Task &Pane',
  ruler: '&Ruler', docMap: 'Document &Map', thumbnails: '&Thumbnails',
  fullScreen: '&Full Screen', zoomDlg: '&Zoom...',
  insertBreak: '&Break...', pageNumbers: 'Page &Numbers', dateTime: '&Date and Time',
  insertSymbol: '&Symbol...', insertComment: '&Comment', insertTable: '&Table...',
  insertClipArt: '&Clip Art', insertPicture: '&From File...',
  fontDlg: '&Font...', paragraphDlg: '&Paragraph...', bulletsDlg: '&Bullets and Numbering...',
  bordersDlg: 'Borders and &Shading...', columnsDlg: '&Columns...', tabsDlg: '&Tabs...',
  stylesPane: 'Styles and &Formatting', revealFormatting: '&Reveal Formatting',
  spelling: '&Spelling and Grammar...', wordCount: '&Word Count...', trackChanges: '&Track Changes',
  mailMerge: '&Mail Merge', help: 'Quire &Help', gettingStarted: '&Getting Started',
  keyboardHelp: '&Keyboard Shortcuts', about: '&About Quire',
  drawLine: '&Line', drawArrow: '&Arrow', drawRect: '&Rectangle', drawOval: '&Oval',
  fillColor: '&Fill Color', lineColor: '&Line Color', selectObjects: '&Select Objects',
  newWindow: '&New Window', arrangeAll: '&Arrange All', splitWindow: '&Split',
}

/* Commands whose Quire implementation asks for a value before touching the document. */
const COMMAND_DIALOGS: Record<string, Dialog> = {
  changeCase: 'changeCase', insertCells: 'insertCells', insertSymbol: 'insertSymbol',
  hyperlink: 'hyperlink', bookmark: 'bookmark', insertField: 'field',
  insertComment: 'comment', mmInsertField: 'mergeSource', mmOpenSource: 'mergeSource',
  tableProps: 'tableProps', bordersDlg: 'borders', formulaDlg: 'formula', sortDlg: 'formula',
  fillColor: 'borders', lineColor: 'borders',
}

function docFromText(value: string, format = 'docx'): Doc {
  const doc = newDoc()
  doc.format = format
  doc.blocks = value.split(/\r?\n/).map((line) => ({ ...emptyPara(), runs: line ? [{ text: line }] : [] }))
  return doc
}

function docFromMarkdown(value: string): Doc {
  const doc = newDoc(); doc.format = 'md'; doc.blocks = []
  for (const line of value.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n').split('\n')) {
    const heading = /^(#{1,6})\s+(.*)$/.exec(line)
    const bullet = /^\s*[-*+]\s+(.*)$/.exec(line)
    const ordered = /^\s*\d+[.)]\s+(.*)$/.exec(line)
    const source = heading?.[2] ?? bullet?.[1] ?? ordered?.[1] ?? line
    const runs = source.split(/(\*\*[^*]+\*\*|__[^_]+__|\*[^*]+\*|_[^_]+_)/g).filter(Boolean).map((part) => {
      const strong = /^(\*\*|__)(.+)\1$/.exec(part); const italic = /^(\*|_)(.+)\1$/.exec(part)
      return strong ? { text: strong[2], rPr: { bold: true } } : italic ? { text: italic[2], rPr: { italic: true } } : { text: part }
    })
    const para = { ...emptyPara(), runs, pPr: heading ? { style: `Heading${heading[1].length}` } : bullet || ordered ? { indent: 24 } : undefined }
    if (bullet) para.runs.unshift({ text: '• ' }); if (ordered) para.runs.unshift({ text: `${line.match(/^\s*(\d+)/)?.[1]}. ` })
    doc.blocks.push(para)
  }
  if (!doc.blocks.length) doc.blocks.push(emptyPara()); return doc
}

function replaceInDoc(doc: Doc, find: string, replace: string): number {
  let count = 0
  const visit = (runs: { text?: string }[]) => {
    for (const run of runs) if (run.text) {
      const parts = run.text.split(find)
      count += parts.length - 1
      run.text = parts.join(replace)
    }
  }
  for (const block of doc.blocks) {
    if (block.kind === 'p') visit(block.runs)
    else for (const row of block.rows) for (const cell of row) for (const para of cell.paras) visit(para.runs)
  }
  return count
}

export const RWord = forwardRef<OfficeApi, RWordProps>(function RWord(
  { src, fileName = 'Document.docx', readonly, onReady, onChange, onSave, onError }, ref,
) {
  const [name, setName] = useState(fileName)
  const [dirty, setDirty] = useState(false)
  const [busy, setBusy] = useState('')
  const [status, setStatus] = useState('')
  const [surfaceVersion, setSurfaceVersion] = useState(0)
  const [view, setView] = useState<'normal' | 'web' | 'print' | 'outline' | 'reading' | 'preview'>('print')
  const [context, setContext] = useState<{ x: number; y: number } | null>(null)
  const [dialog, setDialog] = useState<Dialog>(null)
  const [findText, setFindText] = useState('')
  const [replaceText, setReplaceText] = useState('')
  const [findAllFlag, setFindAllFlag] = useState(false)
  const [matchCase, setMatchCase] = useState(false)
  const [tableRows, setTableRows] = useState(3)
  const [tableCols, setTableCols] = useState(3)
  const [taskPane, setTaskPane] = useState(() => typeof window === 'undefined' || window.innerWidth >= 820)
  const [pane, setPane] = useState<Pane>('getting-started')
  const [rulers, setRulers] = useState(true)
  const [vruler, setVruler] = useState(true)
  const [zoom, setZoom] = useState(1)
  const [docMapOpen, setDocMapOpen] = useState(false)
  const [thumbsOpen, setThumbsOpen] = useState(false)
  const [pvMag, setPvMag] = useState(true)
  const [pvZoomed, setPvZoomed] = useState(false)
  const [pvGrid, setPvGrid] = useState({ rows: 1, cols: 1 })
  const [spellOn, setSpellOn] = useState(true)
  const [marks, setMarks] = useState(false)
  const [trk, setTrk] = useState(false)
  const [ovr, setOvr] = useState(false)
  const [ext, setExt] = useState(false)
  const [font, setFont] = useState('Arial')
  const [size, setSize] = useState('16')
  const [style, setStyle] = useState('Normal')
  const [version, setVersion] = useState(0)
  const [commandArg, setCommandArg] = useState('')
  const [pendingCommand, setPendingCommand] = useState('')
  const [toolbars, setToolbars] = useState<Record<string, boolean>>({
    standard: true, formatting: true, drawing: true,
    tables: false, reviewing: false, outlining: false, headerFooter: false, picture: false, mailMerge: false, wordart: false,
  })

  const docRef = useRef<Doc>(newDoc())
  const pkgRef = useRef<Package | null>(null)
  const partRef = useRef('word/document.xml')
  const sourceRef = useRef<OpenInput | null>(null)
  const surfaceRef = useRef<HTMLDivElement>(null)
  const fileRef = useRef<HTMLInputElement>(null)
  const imageRef = useRef<HTMLInputElement>(null)
  const csvRef = useRef<HTMLInputElement>(null)
  const scrollerRef = useRef<HTMLDivElement>(null)
  const pageRef = useRef<HTMLDivElement>(null)
  const thumbBoxRef = useRef<HTMLDivElement>(null)
  const prevViewRef = useRef<typeof view>('print')
  const pvZoomRef = useRef(1)
  const history = useRef<string[]>([])
  const future = useRef<string[]>([])
  const caret = useRef<{ node: Node; offset: number } | null>(null)

  const mark = useCallback((value: boolean) => { setDirty(value); onChange?.(value) }, [onChange])
  const flash = useCallback((message: string) => { setStatus(message); window.setTimeout(() => setStatus((s) => (s === message ? '' : s)), 4000) }, [])

  const togglePreview = useCallback(() => {
    if (view === 'preview') setView(prevViewRef.current || 'print')
    else { prevViewRef.current = view; setPvZoomed(false); setView('preview') }
  }, [view])
  useDirtyGuard(dirty)

  /* Keep the document model authoritative; the editing surface is rendered from it. */
  const commit = useCallback((options?: { history?: boolean }) => {
    if (options?.history !== false) { history.current.push(JSON.stringify(docRef.current)); future.current = [] }
    setVersion((v) => v + 1)
    setSurfaceVersion((v) => v + 1)
    mark(true)
  }, [mark])

  const open = useCallback(async (input: OpenInput, requestedName?: string): Promise<void> => {
    try {
      setBusy('Opening…')
      const file = input instanceof Blob
        ? { blob: input, name: requestedName ?? (input instanceof File ? input.name : 'Document.docx') }
        : await fetchOnline(typeof input === 'string' ? { kind: 'http', url: input } : input)
      const ext = file.name.split('.').pop()?.toLowerCase()
      const bytes = new Uint8Array(await file.blob.arrayBuffer())
      if (ext === 'md' || ext === 'markdown') {
        docRef.current = docFromMarkdown(readMarkdown(new TextDecoder().decode(bytes)).text)
        pkgRef.current = null
      } else if (ext === 'html' || ext === 'htm') {
        docRef.current = docFromText(htmlToText(new TextDecoder().decode(bytes)), 'html')
        pkgRef.current = null
      } else if (ext === 'txt') {
        docRef.current = docFromText(new TextDecoder().decode(bytes), 'txt')
        pkgRef.current = null
      } else {
        const result = await readDocx(bytes)
        docRef.current = result.doc
        pkgRef.current = result.pkg
        partRef.current = result.part
      }
      sourceRef.current = input
      setName(file.name)
      history.current = []; future.current = []
      commit({ history: false })
      mark(false)
      onReady?.()
    } catch (error) {
      onError?.(error instanceof Error ? error : new Error(String(error)))
    } finally { setBusy('') }
  }, [commit, mark, onReady, onError])

  const getBlob = useCallback(async (): Promise<Blob> => {
    const format = docRef.current.format
    if (format === 'md') return new Blob([writeMarkdown(docText(docRef.current))], { type: 'text/markdown;charset=utf-8' })
    if (format === 'html') return new Blob([textToHtml(docText(docRef.current), name)], { type: 'text/html;charset=utf-8' })
    if (format === 'txt') return new Blob([docText(docRef.current)], { type: 'text/plain;charset=utf-8' })
    return writeDocx(docRef.current, { pkg: pkgRef.current ?? undefined, part: partRef.current, format })
  }, [name])

  const saveAs = useCallback(async (requestedName = name): Promise<Blob> => {
    const blob = await getBlob()
    const output = /\.[a-z0-9]+$/i.test(requestedName) ? requestedName : `${requestedName}.docx`
    downloadBlob(blob, output)
    mark(false)
    return blob
  }, [getBlob, name, mark])

  const save = useCallback(async (): Promise<void> => {
    const blob = await getBlob()
    const source = sourceRef.current
    if (source && typeof source !== 'string' && 'kind' in source && source.kind !== 'http') await saveOnline(source as SaveTarget, blob, name)
    else onSave?.(blob, name)
    mark(false)
  }, [getBlob, name, onSave, mark])

  const newDocument = useCallback((welcome = false) => {
    docRef.current = welcome ? buildWelcomeDoc() : newDoc()
    pkgRef.current = null
    sourceRef.current = null
    history.current = []; future.current = []
    setName('Document.docx')
    commit({ history: false })
    mark(false)
  }, [commit, mark])

  const undo = useCallback(() => {
    const previous = history.current.pop()
    if (previous === undefined) return
    future.current.push(JSON.stringify(docRef.current))
    docRef.current = JSON.parse(previous) as Doc
    commit({ history: false })
  }, [commit])

  const redo = useCallback(() => {
    const next = future.current.pop()
    if (next === undefined) return
    history.current.push(JSON.stringify(docRef.current))
    docRef.current = JSON.parse(next) as Doc
    commit({ history: false })
  }, [commit])

  const syncSurface = useCallback(() => {
    const surface = surfaceRef.current
    if (!surface) return
    const blocks = Array.from(surface.children).filter((element) => !(element as HTMLElement).classList.contains('react-page-break')).map((element, i) => syncBlock(element as HTMLElement, docRef.current, docRef.current.blocks[i]))
    docRef.current = { ...docRef.current, blocks }
    commit()
  }, [commit])

  const insertDocTable = useCallback((rows: number, cols: number) => {
    insertTable(docRef.current, Math.max(1, Math.min(20, rows)), Math.max(1, Math.min(20, cols)))
    commit()
    setDialog(null)
    flash('Table inserted.')
  }, [commit, flash])

  const insertImage = useCallback(async (file: File) => {
    const url = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = () => resolve(String(reader.result))
      reader.onerror = () => reject(reader.error)
      reader.readAsDataURL(file)
    })
    const image = new Image()
    image.src = url
    await image.decode().catch(() => undefined)
    const naturalW = image.naturalWidth || 320
    const naturalH = image.naturalHeight || 180
    const scale = Math.min(1, 576 / naturalW)
    const run = imageRun({ rId: '', part: '', url, w: Math.round(naturalW * scale), h: Math.round(naturalH * scale) }, file.name)
    const blocks = docRef.current.blocks
    const last = blocks[blocks.length - 1]
    if (last && last.kind === 'p') last.runs.push(run)
    else blocks.push({ ...emptyPara(), runs: [run] })
    commit()
    flash(t('Picture inserted.'))
  }, [commit, flash])

  const replaceDocumentText = useCallback((all: boolean) => {
    if (!findText) return
    const options = { text: findText, replace: replaceText, matchCase }
    if (all) {
      const result = replaceAll(docText(docRef.current), options)
      if (result.count) { docRef.current = docFromText(result.value, docRef.current.format); commit(); flash(`Replaced ${result.count} occurrence(s).`) }
    } else {
      const count = replaceInDoc(docRef.current, findText, replaceText)
      if (count) { commit(); flash('Replaced 1 occurrence.') }
    }
  }, [findText, replaceText, matchCase, commit, flash])

  const exec = useCallback((command: string, value?: string) => {
    surfaceRef.current?.focus()
    document.execCommand(command, false, value)
    if (trk && value && /^(insertText|insertHTML|paste)$/.test(command)) {
      const sel = window.getSelection()
      let node: Node | null = sel?.anchorNode ?? null
      if (node && node.nodeType !== Node.ELEMENT_NODE) node = node.parentElement
      const element = node instanceof HTMLElement ? node.closest<HTMLElement>('[data-id]') : null
      recordInsertion(docRef.current, element?.dataset.id, value.replace(/<[^>]*>/g, ''))
    }
    syncSurface()
  }, [syncSurface, trk])

  const doFindNext = useCallback(() => {
    if (!findText) return
    const found = (window as unknown as { find: (text: string, matchCase: boolean) => boolean }).find(findText, matchCase)
    if (!found) flash('Word not found. The search has reached the end of the document.')
  }, [findText, matchCase, flash])

  const runSearch = useCallback((query: string) => {
    const needle = query.trim().toLowerCase()
    if (!needle) return
    const walk = (items: QItem[]): QItem | undefined => {
      for (const item of items) {
        if (item.label && item.label.toLowerCase().includes(needle) && item.run) return item
        const nested = item.sub && walk(item.sub)
        if (nested) return nested
      }
      return undefined
    }
    const commands = menusRef.current.flatMap((m) => walk(m.items) ? [walk(m.items)!] : [])
    if (commands.length) commands[0].run?.()
    else flash(`No command found for “${query}”.`)
  }, [flash])

  const menusRef = useRef<QMenu[]>([])

  useEffect(() => { if (src) void open(src) }, [src]) // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    newDocument(true)
    /* Quire fits the whole page into a narrow window before the first paint. */
    if (window.innerWidth < 820) setZoom(Math.max(0.4, (window.innerWidth - 40) / 816))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      const mod = event.ctrlKey || event.metaKey
      const key = event.key.toLowerCase()
      if (mod) {
        if (key === 's') { event.preventDefault(); void save() }
        else if (key === 'z') { event.preventDefault(); undo() }
        else if (key === 'y') { event.preventDefault(); redo() }
        else if (key === 'n') { event.preventDefault(); newDocument() }
        else if (key === 'o') { event.preventDefault(); fileRef.current?.click() }
        else if (key === 'f' || key === 'h') { event.preventDefault(); setDialog('find') }
        else if (key === 'p') { event.preventDefault(); window.print() }
        else if (key === 'f2') { event.preventDefault(); togglePreview() }
        else if (event.altKey && key === 'i') { event.preventDefault(); togglePreview() }
        else if (key === 'f1') { event.preventDefault(); setTaskPane((shown) => !shown) }
        else if (key === 'b') { event.preventDefault(); exec('bold') }
        else if (key === 'i') { event.preventDefault(); exec('italic') }
        else if (key === 'u') { event.preventDefault(); exec('underline') }
        else if (key === 'l') { event.preventDefault(); exec('justifyLeft') }
        else if (event.shiftKey && key === 'e') { event.preventDefault(); setTrk((v) => !v) }
        else if (key === 'e') { event.preventDefault(); exec('justifyCenter') }
        else if (key === 'r') { event.preventDefault(); exec('justifyRight') }
        else if (key === 'j') { event.preventDefault(); exec('justifyFull') }
      } else if (event.key === 'F1') { event.preventDefault(); setPane('help'); setTaskPane(true) }
      else if (event.key === 'Escape') { setContext(null); setDialog(null) }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [save, undo, redo, newDocument, exec, togglePreview])

  useEffect(() => {
    const surface = surfaceRef.current
    if (!surface) return
    surface.innerHTML = docRef.current.blocks.map((block, index) => {
      const pageBreak = index > 0 && index % 32 === 0 ? '<div class="react-page-break" contenteditable="false" aria-label="Page break"></div>' : ''
      return pageBreak + renderBlock(block, docRef.current)
    }).join('')
    resolveMarkers(surface, docRef.current)
    if (caret.current) {
      try {
        const range = document.createRange()
        range.setStart(caret.current.node, Math.min(caret.current.offset, caret.current.node.textContent?.length || 0))
        range.collapse(true)
        const sel = window.getSelection(); sel?.removeAllRanges(); sel?.addRange(range)
      } catch { /* caret may point at a replaced node */ }
      caret.current = null
    }
  }, [surfaceVersion])

  const captureCaret = useCallback(() => {
    const sel = window.getSelection()
    if (sel && sel.rangeCount && sel.anchorNode && surfaceRef.current?.contains(sel.anchorNode)) {
      caret.current = { node: sel.anchorNode, offset: sel.anchorOffset }
    }
  }, [])

  const activeLocation = useCallback(() => {
    const sel = window.getSelection()
    let node: Node | null = sel?.anchorNode ?? null
    if (node && node.nodeType !== Node.ELEMENT_NODE) node = node.parentElement
    const element = node instanceof HTMLElement ? node.closest<HTMLElement>('[data-id]') : null
    return locate(docRef.current, element?.dataset.id) || fallbackLocation(docRef.current)
  }, [])

  const selectionText = useCallback(() => window.getSelection()?.toString() || '', [])

  const runCommand = useCallback((id: string, arg?: string): boolean => {
    const fn: CommandFn | undefined = documentCommands[id] || reviewCommands[id] || auxiliaryCommands[id]
    if (!fn) return false
    if (arg === undefined && COMMAND_DIALOGS[id]) {
      setCommandArg('')
      setPendingCommand(id)
      setDialog(COMMAND_DIALOGS[id]!)
      return true
    }
    const env: CommandEnv = { doc: docRef.current, loc: activeLocation(), selection: selectionText(), arg }
    const result = fn(env)
    if (typeof result === 'string' && result.startsWith('__LOCATE__')) {
      const target = result.slice('__LOCATE__'.length)
      surfaceRef.current?.querySelector<HTMLElement>(`[data-id="${target}"]`)?.scrollIntoView({ block: 'center' })
      return true
    }
    commit()
    if (result) flash(t(result))
    return true
  }, [activeLocation, selectionText, commit, flash])

  const acceptCommandDialog = useCallback(() => {
    const id = pendingCommand
    setDialog(null)
    if (!id) return
    if (id === 'mmOpenSource') { csvRef.current?.click(); return }
    const value = commandArg.trim()
    runCommand(id, value || undefined)
  }, [pendingCommand, commandArg, runCommand])

  const selectInSurface = useCallback((scope: 'para' | 'cell' | 'row' | 'col' | 'table') => {
    const surface = surfaceRef.current
    const id = activeLocation()?.para?.id
    if (!surface || !id) return
    const paraEl = surface.querySelector<HTMLElement>(`[data-id="${id}"]`)
    let target: HTMLElement | null = null
    if (scope === 'para' || scope === 'cell') target = paraEl
    else {
      const cell = paraEl?.closest('td')
      if (scope === 'row') target = cell?.closest('tr') ?? null
      else if (scope === 'table') target = cell?.closest('table') ?? null
      else if (scope === 'col') {
        const index = cell ? Array.from(cell.parentElement?.children || []).indexOf(cell) : -1
        target = index >= 0 ? (surface.querySelectorAll('table tr')[0]?.children[index] as HTMLElement | undefined) ?? null : null
      }
    }
    if (!target) { flash(t('Place the cursor in a table cell first.')); return }
    const range = document.createRange()
    range.selectNodeContents(target)
    const sel = window.getSelection(); sel?.removeAllRanges(); sel?.addRange(range)
    target.scrollIntoView({ block: 'center' })
  }, [activeLocation, flash])

  const browse = useCallback((direction: number) => {
    const surface = surfaceRef.current
    if (!surface) return
    const blocks = Array.from(surface.children) as HTMLElement[]
    const id = activeLocation()?.para?.id
    let index = blocks.findIndex((b) => b.dataset.id === id || b.querySelector(`[data-id="${id}"]`))
    if (index < 0) index = 0
    const target = blocks[Math.max(0, Math.min(blocks.length - 1, index + direction))]
    if (!target) return
    const range = document.createRange(); range.selectNodeContents(target)
    const sel = window.getSelection(); sel?.removeAllRanges(); sel?.addRange(range)
    target.scrollIntoView({ block: 'center' })
  }, [activeLocation])

  const focusCombo = useCallback((index: number) => {
    document.querySelectorAll<HTMLInputElement>('#tb-formatting .combo input')[index]?.focus()
  }, [])

  const words = useMemo(() => wordCount(docRef.current), [version])
  const section = docRef.current.section
  const pageW = section.pageW / 15
  const pageH = section.pageH / 15
  const margin = { top: section.margins.top / 15, right: section.margins.right / 15, bottom: section.margins.bottom / 15, left: section.margins.left / 15 }
  const contentW = Math.max(200, pageW - margin.left - margin.right)
  const contentH = Math.max(200, pageH - margin.top - margin.bottom)
  const columns = section.columns && section.columns > 1 ? section.columns : 1

  /** Whole-page fit: mirrors Quire's LY.zoomFor('wholePage'/'twoPages'). */
  const fitZoom = useCallback((multi = 1) => {
    const el = scrollerRef.current
    const availW = (el?.clientWidth || window.innerWidth) - 40
    const availH = (el?.clientHeight || window.innerHeight) - 30
    return Math.max(0.1, Math.min((availW - 16 * multi) / (pageW * multi), (availH - 14) / pageH))
  }, [pageW, pageH])

  /* Print Preview: fit the whole page and restore the previous zoom on exit. */
  useEffect(() => {
    if (view !== 'preview') return
    pvZoomRef.current = zoom
    setZoom(fitZoom())
    return () => { setZoom(pvZoomRef.current || 1); setPvZoomed(false) }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [view])

  /* Thumbnails: a scaled clone of the page, exactly like Quire's P.thumbs. */
  useEffect(() => {
    if (!thumbsOpen) return
    const box = thumbBoxRef.current
    const pg = pageRef.current
    if (!box || !pg) return
    box.innerHTML = ''
    const w = Math.max(60, box.clientWidth - 16)
    const k = w / pg.offsetWidth
    const frame = document.createElement('div')
    frame.className = 'thumb on'
    frame.style.width = `${w}px`
    frame.style.height = `${pg.offsetHeight * k}px`
    const clone = pg.cloneNode(true) as HTMLElement
    clone.removeAttribute('data-pg')
    for (const el of Array.from(clone.querySelectorAll('[contenteditable]'))) el.removeAttribute('contenteditable')
    clone.style.cssText += `;transform:scale(${k});transform-origin:0 0;margin:0;pointer-events:none;box-shadow:none`
    frame.addEventListener('click', () => pg.scrollIntoView({ block: 'start' }))
    frame.appendChild(clone)
    const label = document.createElement('div')
    label.className = 'thumb-n'
    label.textContent = '1'
    box.append(frame, label)
  }, [thumbsOpen, version, zoom])

  const actions: Record<string, () => void> = {
    new: () => newDocument(false),
    open: () => fileRef.current?.click(),
    close: () => newDocument(false),
    save: () => void save(),
    saveAs: () => void saveAs(),
    saveWeb: () => setDialog('save'),
    fileSearch: () => { setPane('search'); setTaskPane(true) },
    pageSetup: () => setDialog('pagesetup'),
    exportPdf: () => window.print(),
    print: () => window.print(),
    printPreview: () => togglePreview(),
    undo, redo,
    cut: () => exec('cut'), copy: () => exec('copy'), paste: () => exec('paste'),
    officeClipboard: () => { setPane('getting-started'); setTaskPane(true) },
    pasteSpecial: () => exec('paste'), pasteHyperlink: () => exec('paste'),
    selectAll: () => exec('selectAll'),
    find: () => setDialog('find'), replace: () => setDialog('find'),
    goTo: () => setDialog('find'),
    viewNormal: () => setView('normal'), viewWeb: () => setView('web'), viewPrint: () => setView('print'),
    viewOutline: () => setView('outline'), viewReading: () => setView('reading'),
    taskPane: () => setTaskPane((s) => !s),
    ruler: () => setRulers((s) => !s),
    docMap: () => { setDocMapOpen((s) => !s); setThumbsOpen(false) },
    thumbnails: () => { setThumbsOpen((s) => !s); setDocMapOpen(false) },
    fullScreen: () => { if (!document.fullscreenElement) void document.documentElement.requestFullscreen?.(); else void document.exitFullscreen?.() },
    zoomDlg: () => setDialog('zoom'),
    showMarks: () => setMarks((v) => !v),
    bold: () => exec('bold'), italic: () => exec('italic'), underline: () => exec('underline'),
    strike: () => exec('strikeThrough'), superscript: () => exec('superscript'), subscript: () => exec('subscript'),
    alignLeft: () => exec('justifyLeft'), alignCenter: () => exec('justifyCenter'),
    alignRight: () => exec('justifyRight'), justify: () => exec('justifyFull'),
    bullets: () => exec('insertUnorderedList'), numbering: () => exec('insertOrderedList'),
    decIndent: () => exec('outdent'), incIndent: () => exec('indent'),
    fontDlg: () => setDialog('font'), paragraphDlg: () => setDialog('paragraph'),
    bulletsDlg: () => exec('insertUnorderedList'),
    stylesPane: () => { setPane('styles'); setTaskPane(true) },
    revealFormatting: () => { setPane('reveal'); setTaskPane(true) },
    clearFormats: () => exec('removeFormat'), clearContents: () => exec('delete'),
    wordCount: () => setDialog('wordcount'),
    spelling: () => { setSpellOn((s) => !s); flash('Spelling check toggled.') },
    trackChanges: () => setTrk((s) => !s),
    insertTable: () => setDialog('table'), drawTable: () => setDialog('table'),
    insertPicture: () => imageRef.current?.click(),
    insertClipArt: () => { setPane('getting-started'); setTaskPane(true) },
    help: () => { setPane('help'); setTaskPane(true) },
    gettingStarted: () => { setPane('getting-started'); setTaskPane(true) },
    keyboardHelp: () => { setPane('help'); setTaskPane(true) },
    about: () => setDialog('about'),
    mailMerge: () => { setPane('mailmerge'); setTaskPane(true) },
    sample: () => { newDocument(true); flash('Welcome document opened.') },
    exitApp: () => { window.close(); flash(t('Close this window to exit.')) },
    sendMail: () => {
      const rows = docRef.current.merge?.rows || []
      if (!rows.length) { flash(t('Load a CSV data source before sending mail.')); return }
      const row = rows[docRef.current.merge?.index || 0]
      const merged = docText(docRef.current).replace(/«([^»]+)»/g, (_, key: string) => row[key] ?? '')
      const to = row['Email'] || row['email'] || row['邮件'] || ''
      window.location.href = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(name)}&body=${encodeURIComponent(merged.slice(0, 1800))}`
    },
    properties: () => setDialog('about'),
    selectTable: () => selectInSurface('table'),
    selectRow: () => selectInSurface('row'),
    selectCol: () => selectInSurface('col'),
    selectCell: () => selectInSurface('cell'),
    extendSel: () => selectInSurface('para'),
    browsePrev: () => browse(-1), browseNext: () => browse(1),
    focusFont: () => focusCombo(1), focusSize: () => focusCombo(2), focusStyle: () => focusCombo(0),
  }
  /* Every command with no component-specific implementation is delegated to the
     model command tables, so the full Quire command set stays reachable. */
  const allCommandIds = new Set([
    ...Object.keys(documentCommands), ...Object.keys(reviewCommands), ...Object.keys(auxiliaryCommands),
  ])
  for (const id of allCommandIds) {
    if (!(id in actions)) actions[id] = () => { runCommand(id) }
  }
  const A = actions
  const toggleToolbar = (id: string) => setToolbars((current) => ({ ...current, [id]: !current[id] }))
  const item = (id: string, key?: string, checked?: boolean): QItem => ({
    id, label: COMMAND_LABELS[id] || labelFor(id), icon: iconFor(id), key, checked,
    run: () => {
      const run = A[id]
      if (typeof run === 'function') run()
      else if (!runCommand(id)) flash(t('{0} is not available in this build.', COMMAND_LABELS[id] || labelFor(id)))
    },
  })
  const sep = (): QItem => ({ sep: true, label: '' })

  const menus: QMenu[] = [
    { label: '&File', items: [
      item('new', 'Ctrl+N'), item('open', 'Ctrl+O'), item('close'), sep(),
      item('save', 'Ctrl+S'), item('saveAs', 'F12'), item('saveWeb'), item('fileSearch'), sep(),
      item('pageSetup'), item('printPreview', 'Ctrl+F2'), item('print', 'Ctrl+P'), sep(),
      { label: 'Sen&d To', sub: [item('sendMail'), item('exportPdf')] },
      item('properties'), sep(), item('exitApp'),
    ] },
    { label: '&Edit', items: [
      item('undo', 'Ctrl+Z'), item('redo', 'Ctrl+Y'), sep(),
      item('cut', 'Ctrl+X'), item('copy', 'Ctrl+C'), item('officeClipboard'), item('paste', 'Ctrl+V'),
      item('pasteSpecial'), item('pasteHyperlink'), sep(),
      { label: 'Cle&ar', sub: [item('clearFormats'), item('clearContents')] },
      item('selectAll', 'Ctrl+A'), sep(), item('find', 'Ctrl+F'), item('replace', 'Ctrl+H'), item('goTo', 'Ctrl+G'), sep(),
      item('links'), item('editObject'),
    ] },
    { label: '&View', items: [
      item('viewNormal'), item('viewWeb'), item('viewPrint', 'Ctrl+Alt+P'), item('viewReading'), item('viewOutline'), sep(),
      item('taskPane', 'Ctrl+F1'),
      { label: '&Toolbars', sub: [
        { label: 'Standard', checked: toolbars.standard, run: () => toggleToolbar('standard') },
        { label: 'Formatting', checked: toolbars.formatting, run: () => toggleToolbar('formatting') },
        { label: 'Drawing', checked: toolbars.drawing, run: () => toggleToolbar('drawing') },
        { label: 'Tables and Borders', checked: toolbars.tables, run: () => toggleToolbar('tables') },
        { label: 'Reviewing', checked: toolbars.reviewing, run: () => toggleToolbar('reviewing') },
        { label: 'Picture', checked: toolbars.picture, run: () => toggleToolbar('picture') },
        { label: 'Outlining', checked: toolbars.outlining, run: () => toggleToolbar('outlining') },
        { label: 'Header and Footer', checked: toolbars.headerFooter, run: () => toggleToolbar('headerFooter') },
        { label: 'Mail Merge', checked: toolbars.mailMerge, run: () => toggleToolbar('mailMerge') },
        { label: 'LettersArt', checked: toolbars.wordart, run: () => toggleToolbar('wordart') },
        sep(), item('customize'),
      ] },
      item('ruler', undefined, rulers), sep(), item('docMap'), item('thumbnails'), sep(),
      item('headerFooter'), item('footnotes'), item('markup'), sep(), item('fullScreen'), item('zoomDlg'),
    ] },
    { label: '&Insert', items: [
      item('insertBreak'), item('pageNumbers'), item('dateTime'),
      { label: '&AutoText', sub: [item('autoText')] }, item('insertField'), item('insertSymbol'), item('insertComment'), sep(),
      { label: 'Refere&nce', sub: [item('insertFootnote'), item('insertCaption'), item('crossReference'), item('indexTables')] }, sep(),
      { label: '&Picture', icon: 'picture', sub: [
        item('insertClipArt'), item('insertPicture'), sep(), item('newDrawing'),
        { label: '&AutoShapes', icon: 'autoshapes', run: A.autoshapes }, item('insertWordArt'), item('insertChart'),
      ] },
      item('insertDiagram'), item('textBox'), item('insertFile'), item('insertObject'), item('bookmark'), item('hyperlink'),
    ] },
    { label: 'F&ormat', items: [
      item('fontDlg', 'Ctrl+D'), item('paragraphDlg'), item('bulletsDlg'), item('bordersDlg'), item('columnsDlg'), item('tabsDlg'),
      item('dropCap'), item('textDirection'), item('changeCase'), sep(),
      { label: '&Background', sub: [item('backgroundNone'), item('moreColors'), item('printedWatermark')] },
      item('themeDlg'), item('autoFormat'), sep(), item('stylesPane'), item('revealFormatting'), sep(), item('formatObject'),
    ] },
    { label: '&Tools', items: [
      item('spelling', 'F7'), item('research'),
      { label: '&Language', sub: [item('setLanguage'), item('thesaurus'), item('hyphenation')] },
      item('wordCount'), item('autoSummarize'), sep(),
      item('trackChanges', 'Ctrl+Shift+E'), item('compareDocs'), item('protectDoc'), sep(),
      { label: 'L&etters and Mailings', sub: [item('mailMerge'), item('tb_mailMerge'), item('envelopes'), item('letterWizard')] }, sep(),
      { label: '&Macro', sub: [item('macros')] }, item('templates'), item('autocorrectDlg'), item('customize'), item('optionsDlg'),
    ] },
    { label: 'T&able', items: [
      item('drawTable'),
      { label: '&Insert', sub: [item('insertTable'), sep(), item('colLeft'), item('colRight'), item('rowAbove'), item('rowBelow'), item('insertCells')] },
      { label: '&Delete', sub: [item('deleteTable'), item('deleteCols'), item('deleteRows'), item('deleteCells')] },
      { label: 'Sele&ct', sub: [item('selectTable'), item('selectCol'), item('selectRow'), item('selectCell')] },
      item('mergeCells'), item('splitCells'), item('splitTable'), sep(), item('tableAutoFormat'),
      { label: '&AutoFit', sub: [item('autofitContents'), item('autofitWindow'), item('fixedWidth'), sep(), item('distRows'), item('distCols')] },
      item('headingRows'), { label: 'Con&vert', sub: [item('textToTable'), item('tableToText')] },
      item('sortDlg'), item('formulaDlg'), item('gridlines'), sep(), item('tableProps'),
    ] },
    { label: '&Window', items: [item('newWindow'), item('arrangeAll'), item('splitWindow')] },
    { label: '&Help|menu bar', items: [item('help', 'F1'), item('showAssistant'), sep(), item('gettingStarted'), item('keyboardHelp'), sep(), item('about')] },
  ]
  menusRef.current = menus

  const standardTools: QTool[] = [
    { id: 'new', run: A.new }, { id: 'open', run: A.open }, { id: 'save', run: A.save },
    { id: 'permission', icon: 'permission', tip: 'Restrict Permission', run: A.permission },
    { id: 'sendMail', icon: 'envelope', tip: 'E-mail', run: A.sendMail },
    { id: 'print', run: A.print }, { id: 'printPreview', run: A.printPreview },
    { id: 'spelling', run: A.spelling }, { id: 'research', run: () => { setPane('help'); setTaskPane(true) } }, { sep: true },
    { id: 'cut', run: A.cut }, { id: 'copy', run: A.copy }, { id: 'paste', run: A.paste },
    { id: 'painter', run: A.clearFormats }, { sep: true },
    { id: 'undo', run: A.undo }, { id: 'redo', run: A.redo }, { sep: true },
    { id: 'hyperlink', run: A.hyperlink },
    { id: 'tb_tables', icon: 'tablesBorders', tip: 'Tables and Borders', run: A.insertTable },
    { id: 'table', icon: 'table', tip: 'Insert Table', run: A.insertTable },
    { id: 'insertExcel', icon: 'excel', tip: 'Insert Excel Sheet', run: A.insertExcel },
    { id: 'columns', icon: 'columns', tip: 'Columns', run: A.columnsDlg },
    { id: 'tb_drawing', icon: 'drawing', tip: 'Drawing', run: () => toggleToolbar('drawing') },
    { id: 'docMap', run: A.docMap }, { id: 'showMarks', run: A.showMarks },
    { combo: { value: `${Math.round(zoom * 100)}%`, options: ['500%', '200%', '150%', '100%', '75%', '50%', '25%', '10%', 'Page Width', 'Text Width', 'Whole Page', 'Two Pages'], width: 74,
      onChange: (v) => { const n = parseInt(v, 10); if (n > 0) setZoom(n / 100) } } },
    { id: 'help', run: A.help },
    { id: 'viewReading', icon: 'read', label: 'Read', tip: 'Read Mode', run: A.viewReading },
  ]

  const formattingTools: QTool[] = [
    { id: 'stylesPane', run: A.stylesPane },
    { combo: { value: style, options: ['Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Title', 'Subtitle'], width: 118,
      onChange: (v) => setStyle(v) } },
    { combo: { value: font, options: ['Arial', 'Calibri', 'Times New Roman', 'Courier New', 'Georgia', 'Tahoma', 'Verdana'], width: 136,
      onChange: (v) => { setFont(v); exec('fontName', v) } } },
    { combo: { value: size, options: ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '36', '48', '72'], width: 48,
      onChange: (v) => { setSize(v); exec('fontSize', v) } } },
    { id: 'bold', run: A.bold }, { id: 'italic', run: A.italic }, { id: 'underline', run: A.underline }, { sep: true },
    { id: 'alignLeft', run: A.alignLeft }, { id: 'alignCenter', run: A.alignCenter },
    { id: 'alignRight', run: A.alignRight }, { id: 'justify', run: A.justify }, { sep: true },
    { id: 'lineSpacing', icon: 'lineSpacing', tip: 'Line Spacing', run: A.lineSpacing }, { sep: true },
    { id: 'numbering', run: A.numbering }, { id: 'bullets', run: A.bullets },
    { id: 'decIndent', run: A.decIndent }, { id: 'incIndent', run: A.incIndent }, { sep: true },
    { id: 'borderBtn', icon: 'bordersAll', tip: 'Borders', run: A.bordersDlg },
    { id: 'fontColorBtn', run: () => exec('foreColor', '#C00000') },
    { id: 'highlightBtn', run: () => exec('hiliteColor', '#FFFF00') },
  ]

  const drawingTools: QTool[] = [
    { id: 'draw', label: 'Draw', run: A.draw },
    { id: 'selectObjects', run: A.selectObjects }, { sep: true },
    { id: 'autoshapes', label: 'AutoShapes', run: A.autoshapes },
    { id: 'drawLine', run: A.drawLine }, { id: 'drawArrow', run: A.drawArrow },
    { id: 'drawRect', run: A.drawRect }, { id: 'drawOval', run: A.drawOval },
    { id: 'textBox', run: A.textBox }, { id: 'insertWordArt', run: A.insertWordArt },
    { id: 'insertDiagram', run: A.insertDiagram }, { id: 'insertClipArt', run: A.insertClipArt },
    { id: 'insertPicture', run: A.insertPicture }, { sep: true },
    { id: 'fillColor', run: A.fillColor }, { id: 'lineColor', run: A.lineColor },
    { id: 'fontColorBtn', run: () => exec('foreColor', '#C00000') }, { sep: true },
    { id: 'lineStyle', run: A.lineStyle }, { id: 'dashStyle', run: A.dashStyle },
    { id: 'arrowStyle', run: A.arrowStyle }, { id: 'shadowStyle', run: A.shadowStyle },
  ]

  /* Quire keeps these toolbars in the DOM even while hidden, so their command state and
     dimensions stay in sync when View ▸ Toolbars changes. */
  const hiddenToolbars: { id: string; label: string; items: QTool[] }[] = [
    { id: 'tables', label: 'Tables and Borders', items: [
      { id: 'drawTable', run: A.drawTable }, { id: 'eraser', run: A.eraser }, { sep: true },
      { id: 'dashStyle', run: A.dashStyle }, { id: 'lineStyle', run: A.lineStyle }, { id: 'lineColor', run: A.lineColor },
      { id: 'borderBtn', run: A.bordersDlg }, { id: 'fillColor', run: A.fillColor }, { sep: true },
      { id: 'table', run: A.insertTable }, { id: 'mergeCells', run: A.mergeCells }, { id: 'splitCells', run: A.splitCells }, { sep: true },
      { id: 'alignTop', run: A.alignTop }, { id: 'distRows', run: A.distRows }, { id: 'distCols', run: A.distCols }, { sep: true },
      { id: 'tableAutoFormat', run: A.tableAutoFormat }, { id: 'textDirection', run: A.textDirection }, { sep: true },
      { id: 'sortAsc', run: A.sortAsc }, { id: 'sortDesc', run: A.sortDesc }, { id: 'autoSum', run: A.autoSum },
    ] },
    { id: 'reviewing', label: 'Reviewing', items: [
      { id: 'markupFinalMarkup', label: 'Final Showing Markup', run: A.markupFinalMarkup },
      { id: 'show', label: 'Show', run: A.showFormatting }, { sep: true },
      { id: 'prevChange', run: A.prevChange }, { id: 'nextChange', run: A.nextChange }, { sep: true },
      { id: 'acceptChange', run: A.acceptChange }, { id: 'rejectChange', run: A.rejectChange }, { sep: true },
      { id: 'insertComment', run: A.insertComment }, { id: 'trackChanges', run: A.trackChanges }, { id: 'reviewingPane', run: A.reviewingPane },
    ] },
    { id: 'outlining', label: 'Outlining', items: [
      { id: 'promoteH1', run: A.promoteH1 }, { id: 'promote', run: A.promote },
      { combo: { value: 'Body text', options: ['Level 1', 'Level 2', 'Body text'], width: 74 } },
      { id: 'demote', run: A.demote }, { id: 'demoteBody', run: A.demoteBody }, { sep: true },
      { id: 'moveUp', run: A.moveUp }, { id: 'moveDown', run: A.moveDown },
      { id: 'expand', run: A.expand }, { id: 'collapse', run: A.collapse },
      { combo: { value: 'Show All Levels', options: ['Show All Levels', 'Show Level 1', 'Show Level 2'], width: 96 } },
      { id: 'olFirstLine', run: A.olFirstLine }, { id: 'olFormatting', run: A.olFormatting }, { sep: true },
      { id: 'updateTOC', run: A.updateTOC }, { id: 'gotoTOC', run: A.gotoTOC },
    ] },
    { id: 'headerFooter', label: 'Header and Footer', items: [
      { id: 'hfAutoText', label: 'Insert AutoText', run: A.hfAutoText }, { sep: true },
      { id: 'hfPageNum', run: A.hfPageNum }, { id: 'hfNumPages', run: A.hfNumPages },
      { id: 'hfFormatPageNum', run: A.hfFormatPageNum }, { id: 'hfDate', run: A.hfDate }, { id: 'hfTime', run: A.hfTime }, { sep: true },
      { id: 'pageSetup', run: A.pageSetup }, { id: 'hfShowText', run: A.hfShowText }, { id: 'hfLinkPrev', run: A.hfLinkPrev }, { sep: true },
      { id: 'hfSwitch', run: A.hfSwitch }, { id: 'hfPrev', run: A.hfPrev }, { id: 'hfNext', run: A.hfNext }, { sep: true },
      { id: 'hfClose', label: 'Close', run: A.hfClose },
    ] },
    { id: 'picture', label: 'Picture', items: [
      { id: 'insertPicture', run: A.insertPicture }, { id: 'colorPic', label: 'Color', run: A.colorPic },
      { id: 'picMoreContrast', run: A.picMoreContrast }, { id: 'picLessContrast', run: A.picLessContrast },
      { id: 'picMoreBright', run: A.picMoreBright }, { id: 'picLessBright', run: A.picLessBright }, { sep: true },
      { id: 'picCrop', run: A.picCrop }, { id: 'picRotate', run: A.picRotate }, { id: 'lineStyle', run: A.lineStyle }, { sep: true },
      { id: 'textWrap', run: A.textWrap }, { id: 'formatObject', run: A.formatObject }, { id: 'picReset', run: A.picReset },
    ] },
    { id: 'mailMerge', label: 'Mail Merge', items: [
      { id: 'mmMainDoc', run: A.mailMerge }, { id: 'mmOpenSource', run: A.mmOpenSource }, { id: 'mmRecipients', run: A.mmRecipients }, { sep: true },
      { id: 'mmAddress', run: A.mmAddress }, { id: 'mmGreeting', run: A.mmGreeting }, { id: 'mmInsertField', run: A.mmInsertField }, { sep: true },
      { id: 'mmViewData', run: A.mmViewData }, { id: 'mmFirst', run: A.mmFirst }, { id: 'mmPrev', run: A.mmPrev },
      { combo: { value: `${(docRef.current.merge?.index || 0) + 1}`, options: (docRef.current.merge?.rows || []).map((_, i) => `${i + 1}`), width: 34 } },
      { id: 'mmNext', run: A.mmNext }, { id: 'mmLast', run: A.mmLast }, { sep: true },
      { id: 'mmMergeNew', run: A.mmMergeNew }, { id: 'mmMergePrint', run: A.mmMergePrint },
    ] },
    { id: 'wordart', label: 'LettersArt', items: [
      { id: 'insertWordArt', run: A.insertWordArt }, { id: 'wordartEdit', label: 'Edit Text...', run: A.wordartEdit },
      { id: 'wordartGallery', run: A.wordartGallery }, { id: 'formatObject', run: A.formatObject },
      { id: 'waShape', label: 'LettersArt Shape', run: A.waShape }, { id: 'textWrap', label: 'Text Wrapping', run: A.textWrap },
    ] },
  ]

  /* Print Preview toolbar — mirrors Quire's tb-preview. */
  const pvZoomOptions = ['500%', '200%', '150%', '100%', '75%', '50%', '25%', '10%', 'Page Width', 'Text Width', 'Whole Page', 'Two Pages']
  const pvZoomLabel = pvZoomOptions.includes(`${Math.round(zoom * 100)}%`) ? `${Math.round(zoom * 100)}%` : 'Whole Page'
  const applyPvZoom = (value: string) => {
    setPvZoomed(false)
    if (value.endsWith('%')) setZoom(parseInt(value, 10) / 100)
    else if (value === 'Two Pages') setZoom(fitZoom(2))
    else if (value === 'Whole Page') setZoom(fitZoom())
    else setZoom(Math.max(0.1, Math.min(1, ((scrollerRef.current?.clientWidth || window.innerWidth) - 40) / pageW)))
  }
  const previewTools: QTool[] = [
    { id: 'print', run: () => window.print() },
    { id: 'pvMag', tip: 'Magnifier', icon: 'zoom', checked: pvMag, run: () => setPvMag((v) => !v) },
    { id: 'onePage', tip: 'One Page', run: () => { setPvGrid({ rows: 1, cols: 1 }); setZoom(fitZoom()) } },
    { id: 'multiPage', tip: 'Multiple Pages', run: () => setDialog('pvGrid') },
    { combo: { value: pvZoomLabel, options: pvZoomOptions, width: 74, onChange: applyPvZoom } },
    { id: 'ruler', checked: rulers, run: () => setRulers((s) => !s) },
    { id: 'shrink', tip: 'Shrink to Fit', run: () => flash(t('This document is already one page long.')) },
    { id: 'fullScreen', run: A.fullScreen },
    { id: 'closePreview', label: t('Close Preview'), run: togglePreview },
    { id: 'help', run: A.help },
  ]

  /* Reading Layout toolbar — mirrors Quire's readbar. */
  const readingTools: QTool[] = [
    { id: 'docMap', label: t('Document Map'), run: A.docMap },
    { id: 'thumbnails', label: t('Thumbnails'), run: A.thumbnails },
    { id: 'find', label: t('Find'), run: A.find },
    { id: 'research', label: t('Research'), run: () => { setPane('search'); setTaskPane(true) } },
    { sep: true },
    { id: 'growFont', tip: 'Increase Text Size', run: () => setZoom((z) => Math.min(4, z * 1.1)) },
    { id: 'shrinkFont', tip: 'Decrease Text Size', run: () => setZoom((z) => Math.max(0.3, z / 1.1)) },
    { id: 'onePage', tip: 'Actual Page', run: () => { setPvGrid({ rows: 1, cols: 1 }); setZoom(fitZoom()) } },
    { id: 'multiPage', tip: 'Allow Multiple Pages', run: () => setZoom(fitZoom(2)) },
    { sep: true },
    { id: 'close', label: t('Close'), run: () => setView(prevViewRef.current !== 'reading' && prevViewRef.current !== 'preview' ? prevViewRef.current : 'print') },
  ]

  const contextItems: QItem[] = [
    item('cut', 'Ctrl+X'), item('copy', 'Ctrl+C'), item('paste', 'Ctrl+V'), sep(),
    { label: 'Font…', run: A.fontDlg }, { label: 'Paragraph…', run: A.paragraphDlg },
    { label: 'Bullets and Numbering…', run: A.bulletsDlg }, sep(),
    { label: 'Insert Table…', run: A.insertTable }, sep(),
    { label: 'Hyperlink…', run: A.hyperlink },
  ]

  const pageCount = Math.max(1, Math.ceil(docRef.current.blocks.length / 32))
  const hf = docRef.current.headerFooter || {}
  const resolveFields = (value: string) => value.replace(/«PAGE»/g, '1').replace(/«NUMPAGES»/g, String(pageCount))
  const headerText = hf.header ? resolveFields(hf.header) : 'Quire'
  const footerText = hf.footer ? resolveFields(hf.footer) : t('Page {0} of {1}', 1, 1)

  const pageStyle = {
    width: `${pageW}px`,
    height: `${pageH}px`,
  }
  const bodyStyle = {
    left: `${margin.left}px`, top: `${margin.top}px`,
    width: `${contentW}px`, height: `${contentH}px`,
    columnCount: columns > 1 ? columns : undefined,
    columnGap: columns > 1 ? '48px' : undefined,
  }

  const paneContent = (): ReactNode => {
    if (pane === 'getting-started') return <>
      <div className="tp-row"><QIcon name="app" size={32} /><b translate="no" style={{ color: 'var(--pane-ink)' }}>Quire Web Edition</b></div>
      <div className="tp-sec"><div className="tp-h">{t('Search for:')}</div><div className="tp-row">
        <input type="text" placeholder={t('Search help')} style={{ flex: 1, minWidth: 0 }} onKeyDown={(e) => { if (e.key === 'Enter') { setPane('help'); e.stopPropagation() } }} />
        <button className="btn small" onClick={() => setPane('help')}>{t('Go')}</button>
      </div></div>
      <div className="tp-sec"><div className="tp-h">{t('Open')}</div>
        <button className="tp-link" onClick={A.open}>{t('More...')}</button>
        <button className="tp-link" onClick={() => newDocument(false)}>{t('Create a new document...')}</button>
        <button className="tp-link" onClick={A.sample}>{t('Open the welcome document')}</button>
      </div>
      <div className="tp-sec"><div className="tp-h">{t('Learn')}</div>
        <button className="tp-link" onClick={A.keyboardHelp}>{t('Keyboard shortcuts and help')}</button>
        <button className="tp-link" onClick={A.stylesPane}>{t('Styles and Formatting')}</button>
        <button className="tp-link" onClick={A.mailMerge}>{t('Mail Merge wizard')}</button>
      </div>
    </>
    if (pane === 'help') return <div className="tp-sec">
      <div className="tp-h">{t('Keyboard shortcuts')}</div>
      <table className="kbd-table">
        <tbody>
          {[['Ctrl+N / O / S / P', 'New, Open, Save, Print'], ['Ctrl+Z / Ctrl+Y', 'Undo / Redo'], ['Ctrl+X / C / V', 'Cut, Copy, Paste'],
            ['Ctrl+A', 'Select all'], ['Ctrl+F / H', 'Find, Replace'], ['Ctrl+B / I / U', 'Bold, Italic, Underline'],
            ['Ctrl+L / E / R / J', 'Left, Center, Right, Justify'], ['Ctrl+F1', 'Task pane'], ['F1', 'Help'],
          ].map(([k, d]) => <tr key={k}><td translate="no">{k}</td><td>{t(d)}</td></tr>)}
        </tbody>
      </table>
      <button className="tp-link" onClick={() => setPane('getting-started')}>{t('Show all help topics')}</button>
    </div>
    if (pane === 'search') return <div className="tp-sec"><div className="tp-h">{t('Basic File Search')}</div>
      <input type="text" placeholder={t('Search this document')} style={{ width: '100%' }}
        onChange={(e) => { if (e.target.value) doFind() }} />
      <p className="tp-note">{t('Type to find text in the current document.')}</p>
    </div>
    if (pane === 'styles') return <div className="tp-sec"><div className="tp-h">{t('Styles and Formatting')}</div>
      <div className="sf-list">
        {['Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Title', 'Subtitle'].map((s) =>
          <button className={`sf-item${style === s ? ' on' : ''}`} key={s} onClick={() => { setStyle(s); exec('formatBlock', s) }}>
            <span>{t(s)}</span><span>¶</span>
          </button>)}
      </div>
    </div>
    if (pane === 'reveal') return <div className="tp-sec"><div className="tp-h">{t('Reveal Formatting')}</div>
      <p className="tp-note">{t('Font: {0}, {1} pt', font, size)}<br />{t('Style: {0}', style)}<br />{t('Language: {0}', t('English (U.S.)'))}</p>
    </div>
    if (pane === 'mailmerge') {
      const merge = docRef.current.merge || {}
      const rows = merge.rows || []
      const headers = Object.keys(rows[0] || {})
      return <div>
        <div className="tp-sec"><div className="tp-h">{t('Mail Merge')}</div>
          <div className="tp-row"><button className="btn small" onClick={() => csvRef.current?.click()}>{t('Open Data Source…')}</button>
            <span>{rows.length} {t('records')}</span></div>
        </div>
        <div className="tp-sec"><div className="tp-h">{t('Insert Merge Field')}</div>
          <div className="tp-row">{headers.map((key) =>
            <button className="btn small" key={key} onClick={() => runCommand('mmInsertField', key)}>«{key}»</button>)}</div>
        </div>
        <div className="tp-sec"><div className="tp-h">{t('View Record')}</div>
          <div className="tp-row">
            <button className="btn small" onClick={A.mmFirst}>{'|<'}</button>
            <button className="btn small" onClick={A.mmPrev}>{'<'}</button>
            <span>{rows.length ? `${(merge.index || 0) + 1} / ${rows.length}` : '0 / 0'}</span>
            <button className="btn small" onClick={A.mmNext}>{'>'}</button>
            <button className="btn small" onClick={A.mmLast}>{'|>'}</button>
          </div>
          <p className="tp-note">{rows[merge.index || 0] ? Object.entries(rows[merge.index || 0]!).map(([k, v]) => `${k}: ${v}`).join(', ') : t('No data source loaded.')}</p>
        </div>
        <div className="tp-btns"><button className="btn" onClick={A.mmMergeNew}>{t('Merge to New Document')}</button></div>
      </div>
    }
    if (pane === 'new') return <div className="tp-sec"><div className="tp-h">{t('New')}</div>
      <button className="tp-link" onClick={() => newDocument(false)}>{t('Blank document')}</button>
      <button className="tp-link" onClick={A.sample}>{t('Welcome document')}</button>
    </div>
    return <div className="tp-sec"><div className="tp-h">{t('Task Pane')}</div>
      {(['getting-started', 'help', 'search', 'styles', 'reveal', 'mailmerge'] as Pane[]).map((p) =>
        <button className="tp-link" key={p} onClick={() => setPane(p)}>{t(p.replace('-', ' ').replace(/\b\w/g, (c) => c.toUpperCase()))}</button>)}
    </div>
  }
  const doFind = () => { if (findText) doFindNext() }

  const sb = (cls: string, tip: string | undefined, value: ReactNode, onClick?: () => void, onDblClick?: () => void) =>
    <div className={`sb-cell ${cls}`} data-tip={tip ? t(tip) : undefined} onClick={onClick} onDoubleClick={onDblClick}>{typeof value === 'string' ? t(value) : value}</div>

  return <ErrorBoundary>
    <div className={`app${view === 'reading' ? ' reading' : ''}${view === 'preview' ? ' preview' : ''}`} aria-label={t('Quire word processor')}
      data-view={view} data-look="classic">
      <input ref={fileRef} id="rword-file" hidden type="file" accept=".docx,.docm,.dotx,.dotm,.md,.markdown,.html,.htm,.txt"
        onChange={(e) => { const file = e.target.files?.[0]; if (file) void open(file); e.target.value = '' }} />
      <input ref={imageRef} id="rword-image" hidden type="file" accept="image/*"
        onChange={(e) => { const file = e.target.files?.[0]; if (file) void insertImage(file); e.target.value = '' }} />
      <input ref={csvRef} id="rword-csv" hidden type="file" accept=".csv,.tsv,text/csv"
        onChange={(e) => {
          const file = e.target.files?.[0]
          if (file) void file.text().then((text) => {
            const count = loadMergeSource(docRef.current, text)
            commit(); setPane('mailmerge'); setTaskPane(true)
            flash(t('{0} records loaded.', count))
          })
          e.target.value = ''
        }} />

      <div className="cmdbars">
        <QMenuBar menus={menus} onSearch={runSearch} />
        <div className="toolbars" id="toolbars">
          <div className="tb-row" hidden={!toolbars.standard}><QToolbar id="standard" label={t('Standard')} items={standardTools} /></div>
          <div className="tb-row" hidden={!toolbars.formatting}><QToolbar id="formatting" label={t('Formatting')} items={formattingTools} /></div>
          <div className="tb-row" hidden={!hiddenToolbars.some((bar) => toolbars[bar.id])}>
            {hiddenToolbars.map((bar) => <QToolbar key={bar.id} id={bar.id} label={t(bar.label)} items={bar.items} hidden={!toolbars[bar.id]} />)}
          </div>
          {view === 'preview' && <div className="tb-row preview-bar"><QToolbar id="preview" label={t('Print Preview')} items={previewTools} /></div>}
          {view === 'reading' && <div className="readbar" role="toolbar" aria-label={t('Reading Layout')}>
            {readingTools.map((it, i) => <QToolButton key={`${it.id || it.label || 'r'}-${i}`} item={it} />)}
          </div>}
        </div>
      </div>

      <div className="main" id="main">
        {(docMapOpen || thumbsOpen) && <div className="docmap" id="docmap">
          {thumbsOpen ? <div className="thumbs" ref={thumbBoxRef} /> :
            docRef.current.blocks.filter((b) => b.kind === 'p').map((b) => {
              const p = b as Extract<typeof b, { kind: 'p' }>
              if (!p.pPr?.style?.startsWith('Heading')) return null
              return <button className="dm-item" key={p.id} onClick={() => surfaceRef.current?.querySelector<HTMLElement>(`[data-id="${p.id}"]`)?.scrollIntoView()}>
                {p.runs.map((r) => r.text || '').join('')}
              </button>
            })}
        </div>}
        <div className="center" id="center">
          <div className={`work${rulers ? ' rulers' : ''}${vruler ? ' vruler' : ''}`} id="work">
            <div className="r-corner" id="rcorner"><button className="tabsel" title={t('Toggle rulers')} onClick={() => setVruler((v) => !v)} /></div>
            <div className="ruler-h" id="hruler">
              <canvas ref={(el) => { if (el) drawRuler(el, 'horizontal', contentW, zoom) }} />
            </div>
            <div className="ruler-v" id="vruler">
              <canvas ref={(el) => { if (el) drawRuler(el, 'vertical', contentH, zoom) }} />
            </div>
            <div className={`scroller${view === 'web' || view === 'normal' ? ' flowview' : ''}${view === 'web' ? ' webview' : ''}${view === 'preview' && pvZoomed ? ' zoomed' : ''}`} id="scroller"
              ref={scrollerRef} tabIndex={-1}
              onClick={view === 'preview' && pvMag ? () => {
                if (zoom < 0.95) { setZoom(1); setPvZoomed(true) } else { setZoom(fitZoom()); setPvZoomed(false) }
              } : undefined}
              onContextMenu={(e) => { e.preventDefault(); captureCaret(); setContext({ x: e.clientX, y: e.clientY }) }}>
              <div className="sizer">
                <div className={`pages${columns > 1 ? ' multi' : ''}${marks ? ' marks' : ''}`} id="pages" style={{ width: '100%', transform: `scale(${zoom})` }}>
                  <div className="pg" ref={pageRef} style={pageStyle}>
                    <div className="pg-hdr" data-label={t('Header')} style={{ left: margin.left, top: margin.top / 2, right: margin.right, fontSize: 12 }}>
                      <div className="p"><div className="pc"><span>{headerText}</span><span style={{ float: 'right' }}>{t('Sample document')}</span></div></div>
                    </div>
                    <div className="pg-body" style={bodyStyle}>
                      <div className="pg-reg" style={{ width: '100%' }}>
                        <div className="pg-col" style={{ width: '100%', fontSize: '16px' }}
                          ref={surfaceRef} contentEditable={!readonly && view !== 'preview'} suppressContentEditableWarning
                          onInput={syncSurface} onMouseUp={captureCaret} onKeyUp={captureCaret}
                          aria-label="Document editor" spellCheck={spellOn} />
                      </div>
                    </div>
                    <div className="pg-fn" style={{ left: margin.left, right: margin.right, bottom: margin.bottom + 12, fontSize: 10 }}>
                      <div className="p"><div className="pc"><span>1</span> {t('Footnotes are numbered automatically and placed at the bottom of the page.')}</div></div>
                    </div>
                    <div className="pg-ftr" data-label={t('Footer')} style={{ left: margin.left, right: margin.right, bottom: margin.bottom / 2, fontSize: 12 }}>
                      <div className="p"><div className="pc"><span style={{ float: 'right' }}>{footerText}</span></div></div>
                    </div>
                    <div className="crop" />
                  </div>
                </div>
              </div>
            </div>
          </div>
          <QViewBar view={view} zoom={zoom} onView={(v) => setView(v as typeof view)}
            onBrowse={(dir) => flash(t(dir < 0 ? 'Previous page.' : 'Next page.'))} />
        </div>

        <QTaskPane hidden={!taskPane} title={taskPane ? (pane === 'getting-started' ? 'Quire Web Edition' : t(pane === 'help' ? 'Quire Help' : pane === 'styles' ? 'Styles and Formatting' : pane === 'reveal' ? 'Reveal Formatting' : pane === 'mailmerge' ? 'Mail Merge' : pane === 'new' ? 'New Document' : pane === 'search' ? 'Basic File Search' : 'Task Pane')) : ''}
          onNav={(action) => {
            if (action === 'home') setPane('getting-started')
            else if (action === 'menu') setPane('menu')
            else if (action === 'close') setTaskPane(false)
          }}>{taskPane ? paneContent() : null}</QTaskPane>
      </div>

      <div className="drawrow" id="drawrow" hidden={!toolbars.drawing}><QToolbar id="drawing" label={t('Drawing')} items={drawingTools} /></div>

      <div className="statusbar" id="statusbar">
        {sb('', 'Page', t('Page {0}', 1), undefined, A.goTo)}
        {sb('', 'Section', t('Sec {0}', 1))}
        {sb('', 'Page / pages', `1/${pageCount}`)}
        {sb('', 'Vertical position', t('At {0}', '2.96 cm'))}
        {sb('', 'Line|position', t('Ln {0}', 1))}
        {sb('', 'Column|position', t('Col {0}', 1))}
        {sb('flag', 'Record macro', 'REC')}
        {sb('flag act' + (trk ? ' on' : ''), 'Track Changes (double-click to toggle)', 'TRK', undefined, A.trackChanges)}
        {sb('flag act' + (ext ? ' on' : ''), 'Extend selection', 'EXT', undefined, () => setExt((v) => !v))}
        {sb('flag act' + (ovr ? ' on' : ''), 'Overtype (double-click to toggle)', 'OVR', undefined, () => setOvr((v) => !v))}
        {sb('act', 'Language', 'English (U.S.)')}
        {sb('act sb-spell' + (spellOn ? '' : ' off'), 'Spelling and grammar status',
          <svg width="18" height="14" viewBox="0 0 18 14"><rect x="1.5" y="1.5" width="11" height="11" fill="#fff" stroke="#556" /><path d="M3 4h7M3 6.5h7M3 9h5" stroke="#889" /><path d="M10 9l2.5 3 4.5-7" fill="none" stroke="#2f9a3a" strokeWidth="1.6" /></svg>,
          () => setSpellOn((s) => !s))}
        {sb('grow', undefined, status || busy)}
      </div>

      {context && <div className="menu" role="menu" style={{ left: context.x, top: context.y }} onMouseLeave={() => setContext(null)}>
        {contextItems.map((it, i) => it.sep ? <div className="m-sep" key={`s${i}`} /> :
          <div className="m-item" role="menuitem" key={`${it.label}-${i}`} onClick={() => { setContext(null); void it.run?.() }}>
            <span className="m-ico">{it.icon ? <QIcon name={iconFor(it.icon)} /> : null}</span>
            <span className="m-lbl">{t(it.label || '')}</span>{it.key && <span className="m-key">{it.key}</span>}
          </div>)}
      </div>}

      {dialog === 'save' && <QDialog title={t('Save to Online')} onClose={() => setDialog(null)}>
        <p>{name}</p><p className="tp-note">{t('Local and online saving via the office API. Use Save As for a local download.')}</p>
        <div className="tp-btns"><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button><button className="btn" onClick={() => { setDialog(null); void saveAs() }}>{t('Save locally')}</button></div>
      </QDialog>}
      {dialog === 'about' && <QDialog title={t('About Quire')} onClose={() => setDialog(null)}>
        <p>{t('Quire is the VibeOffice word processor — it opens and saves real .docx files.')}</p>
      </QDialog>}
      {dialog === 'find' && <QDialog title={t('Find and Replace')} onClose={() => setDialog(null)} wide>
        <div className="row"><label>{t('Find what')}</label><input type="text" value={findText} onChange={(e) => setFindText(e.target.value)} /></div>
        <div className="row"><label>{t('Replace with')}</label><input type="text" value={replaceText} onChange={(e) => setReplaceText(e.target.value)} /></div>
        <div className="row"><label className="ck"><input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} /> {t('Match case')}</label>
          <label className="ck"><input type="checkbox" checked={findAllFlag} onChange={(e) => setFindAllFlag(e.target.checked)} /> {t('Find all')}</label></div>
        <div className="tp-btns">
          <button className="btn" onClick={doFindNext}>{t('Find Next')}</button>
          <button className="btn" onClick={() => replaceDocumentText(false)}>{t('Replace')}</button>
          <button className="btn" onClick={() => replaceDocumentText(true)}>{t('Replace All')}</button>
          <button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button>
        </div>
      </QDialog>}
      {dialog === 'table' && <QDialog title={t('Insert Table')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Number of rows')}</label><input type="number" min="1" max="20" value={tableRows} onChange={(e) => setTableRows(Math.max(1, Math.min(20, Number(e.target.value))))} /></div>
        <div className="row"><label>{t('Number of columns')}</label><input type="number" min="1" max="20" value={tableCols} onChange={(e) => setTableCols(Math.max(1, Math.min(20, Number(e.target.value))))} /></div>
        <div className="tp-btns"><button className="btn" onClick={() => insertDocTable(tableRows, tableCols)}>{t('OK')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
      {dialog === 'wordcount' && <QDialog title={t('Word Count')} onClose={() => setDialog(null)}>
        <table className="wc-table"><tbody>
          <tr><td>{t('Pages')}</td><td>1</td></tr><tr><td>{t('Words')}</td><td>{words}</td></tr>
          <tr><td>{t('Characters (no spaces)')}</td><td>{docText(docRef.current).replace(/\s/g, '').length}</td></tr>
          <tr><td>{t('Characters (with spaces)')}</td><td>{docText(docRef.current).length}</td></tr>
          <tr><td>{t('Paragraphs')}</td><td>{docRef.current.blocks.length}</td></tr>
        </tbody></table>
        <div className="tp-btns"><button className="btn" onClick={() => setDialog(null)}>{t('Close')}</button></div>
      </QDialog>}
      {dialog === 'font' && <QDialog title={t('Font')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Font')}</label><select value={font} onChange={(e) => { setFont(e.target.value); exec('fontName', e.target.value) }}>
          {['Arial', 'Calibri', 'Times New Roman', 'Courier New', 'Georgia', 'Tahoma', 'Verdana'].map((f) => <option key={f}>{f}</option>)}
        </select></div>
        <div className="row"><label>{t('Size')}</label><select value={size} onChange={(e) => { setSize(e.target.value); exec('fontSize', e.target.value) }}>
          {['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '36', '48'].map((s) => <option key={s}>{s}</option>)}
        </select></div>
        <div className="row"><label>{t('Effects')}</label>
          <button className="btn" onClick={() => { setDialog(null); runCommand('hiddenText') }}>{t('Hidden')}</button>
          <button className="btn" onClick={() => { setDialog(null); runCommand('allCaps') }}>{t('All caps')}</button>
          <button className="btn" onClick={() => { setDialog(null); runCommand('smallCaps') }}>{t('Small caps')}</button>
          <button className="btn" onClick={() => { setDialog(null); runCommand('dblUnderline') }}>{t('Double underline')}</button>
        </div>
        <div className="tp-btns"><button className="btn" onClick={() => setDialog(null)}>{t('OK')}</button></div>
      </QDialog>}
      {dialog === 'paragraph' && <QDialog title={t('Paragraph')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Alignment')}</label>
          <button className="btn" onClick={A.alignLeft}>{t('Left')}</button><button className="btn" onClick={A.alignCenter}>{t('Center')}</button>
          <button className="btn" onClick={A.alignRight}>{t('Right')}</button><button className="btn" onClick={A.justify}>{t('Justify')}</button>
        </div>
        <div className="tp-btns"><button className="btn" onClick={() => setDialog(null)}>{t('OK')}</button></div>
      </QDialog>}
      {(dialog === 'pagesetup' || dialog === 'zoom') && <QDialog title={t(dialog === 'zoom' ? 'Zoom' : 'Page Setup')} onClose={() => setDialog(null)}>
        {dialog === 'zoom' ? <div className="row"><label>{t('Zoom')}</label>
          <select value={`${Math.round(zoom * 100)}%`} onChange={(e) => setZoom(parseInt(e.target.value, 10) / 100)}>
            {['50%', '75%', '100%', '125%', '150%', '200%'].map((v) => <option key={v}>{v}</option>)}
          </select></div>
          : <p className="tp-note">{t('Paper: Letter (8.5" × 11")')}<br />{t('Margins: 1" on all sides')}<br />{t('Orientation: Portrait')}</p>}
        <div className="tp-btns"><button className="btn" onClick={() => setDialog(null)}>{t('OK')}</button></div>
      </QDialog>}
      {dialog === 'pvGrid' && <QDialog title={t('Multiple Pages')} onClose={() => setDialog(null)}>
        <div className="tbl-grid-pick">
          <div className="tgp-cells" style={{ gridTemplateColumns: 'repeat(3,24px)' }}>
            {Array.from({ length: 6 }, (_, i) => {
              const rows = Math.floor(i / 3) + 1
              const cols = (i % 3) + 1
              const on = pvGrid.rows >= rows && pvGrid.cols >= cols
              return <span key={i} className={on ? 'on' : ''} style={{ width: 24, height: 30 }} role="button"
                aria-label={t('{0} x {1} Pages', rows, cols)}
                onMouseEnter={(e) => {
                  const cells = e.currentTarget.parentElement?.children || []
                  Array.from(cells).forEach((cell, k) => cell.classList.toggle('on', Math.floor(k / 3) <= Math.floor(i / 3) && (k % 3) <= i % 3))
                }}
                onClick={() => { setPvGrid({ rows, cols }); setZoom(fitZoom(cols)); setDialog(null) }} />
            })}
          </div>
          <div className="tgp-label">{t('{0} x {1} Pages', pvGrid.rows, pvGrid.cols)}</div>
        </div>
        <div className="tp-btns"><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}

      {dialog === 'changeCase' && <QDialog title={t('Change Case')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Case')}</label>
          <select value={commandArg || 'title'} onChange={(e) => setCommandArg(e.target.value)}>
            <option value="upper">UPPERCASE</option><option value="lower">lowercase</option>
            <option value="title">Title Case</option><option value="sentence">Sentence case</option>
          </select></div>
        <div className="tp-btns"><button className="btn" onClick={acceptCommandDialog}>{t('OK')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
      {dialog === 'insertCells' && <QDialog title={t('Insert Cells')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Number of cells')}</label>
          <input type="number" min="1" max="20" value={commandArg || '1'} onChange={(e) => setCommandArg(e.target.value)} /></div>
        <div className="tp-btns"><button className="btn" onClick={acceptCommandDialog}>{t('OK')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
      {dialog === 'insertSymbol' && <QDialog title={t('Symbol')} onClose={() => setDialog(null)}>
        <div className="tp-row">
          {['©', '®', '™', '…', '—', '–', '§', '¶', '•', '°', '±', '×', '÷', '€', '£', '¥'].map((c) =>
            <button className="btn small" key={c} onClick={() => { setCommandArg(c); acceptCommandDialog() }}>{c}</button>)}
        </div>
      </QDialog>}
      {dialog === 'hyperlink' && <QDialog title={t('Insert Hyperlink')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Address')}</label>
          <input type="text" value={commandArg} placeholder="https://" onChange={(e) => setCommandArg(e.target.value)} /></div>
        <div className="tp-btns"><button className="btn" onClick={acceptCommandDialog}>{t('OK')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
      {dialog === 'bookmark' && <QDialog title={t('Bookmark')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Bookmark name')}</label>
          <input type="text" value={commandArg} onChange={(e) => setCommandArg(e.target.value)} /></div>
        <div className="tp-btns"><button className="btn" onClick={acceptCommandDialog}>{t('OK')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
      {dialog === 'field' && <QDialog title={t('Field')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Field name')}</label>
          <select value={commandArg || 'DATE'} onChange={(e) => setCommandArg(e.target.value)}>
            {['DATE', 'TIME', 'PAGE', 'NUMPAGES', 'TITLE', 'AUTHOR', 'FILENAME'].map((f) => <option key={f}>{f}</option>)}
          </select></div>
        <div className="tp-btns"><button className="btn" onClick={acceptCommandDialog}>{t('OK')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
      {dialog === 'comment' && <QDialog title={t('Insert Comment')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Comment')}</label>
          <textarea value={commandArg} rows={4} onChange={(e) => setCommandArg(e.target.value)} /></div>
        <div className="tp-btns"><button className="btn" onClick={acceptCommandDialog}>{t('OK')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
      {dialog === 'mergeSource' && <QDialog title={t('Mail Merge Data Source')} onClose={() => setDialog(null)}>
        <p className="tp-note">{t('Choose a CSV file with a header row.')}</p>
        <div className="tp-btns"><button className="btn" onClick={acceptCommandDialog}>{t('Open…')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
      {dialog === 'tableProps' && <QDialog title={t('Table Properties')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Alignment')}</label>
          <select value={commandArg || 'left'} onChange={(e) => setCommandArg(e.target.value)}>
            <option value="left">{t('Left')}</option><option value="center">{t('Center')}</option><option value="right">{t('Right')}</option>
          </select></div>
        <div className="tp-btns"><button className="btn" onClick={acceptCommandDialog}>{t('OK')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
      {dialog === 'borders' && <QDialog title={t('Borders and Shading')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Colour')}</label>
          <input type="color" value={commandArg || '#1D4F9E'} onChange={(e) => setCommandArg(e.target.value)} /></div>
        <div className="tp-btns"><button className="btn" onClick={acceptCommandDialog}>{t('OK')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
      {dialog === 'formula' && <QDialog title={t('Formula and Sorting')} onClose={() => setDialog(null)}>
        <div className="row"><label>{t('Operation')}</label>
          <select value={commandArg || 'sum'} onChange={(e) => setCommandArg(e.target.value)}>
            <option value="sum">{t('Sum')}</option><option value="average">{t('Average')}</option>
            <option value="asc">{t('Sort ascending')}</option><option value="desc">{t('Sort descending')}</option>
          </select></div>
        <div className="tp-btns"><button className="btn" onClick={acceptCommandDialog}>{t('OK')}</button><button className="btn" onClick={() => setDialog(null)}>{t('Cancel')}</button></div>
      </QDialog>}
    </div>
  </ErrorBoundary>
})

/* Draw a classic measuring ruler onto a canvas. */
function drawRuler(canvas: HTMLCanvasElement, orientation: 'horizontal' | 'vertical', length: number, zoom: number): void {
  const ratio = window.devicePixelRatio || 1
  const wide = orientation === 'horizontal'
  const width = wide ? Math.max(600, length + 40) : 20
  const height = wide ? 20 : Math.max(600, length + 40)
  if (canvas.width !== Math.round(width * ratio) || canvas.height !== Math.round(height * ratio)) {
    canvas.width = Math.round(width * ratio)
    canvas.height = Math.round(height * ratio)
    canvas.style.width = `${width}px`
    canvas.style.height = `${height}px`
  }
  const ctx = canvas.getContext('2d')
  if (!ctx) return
  ctx.setTransform(ratio, 0, 0, ratio, 0, 0)
  ctx.clearRect(0, 0, width, height)
  const css = getComputedStyle(document.documentElement)
  const tick = css.getPropertyValue('--ruler-tick').trim() || '#444'
  const ink = css.getPropertyValue('--ruler-ink').trim() || '#222'
  const step = 16 * zoom
  ctx.strokeStyle = tick
  ctx.fillStyle = ink
  ctx.font = '9px Tahoma, sans-serif'
  ctx.lineWidth = 1
  for (let i = 0, pos = 48 * zoom; pos < (wide ? width : height); i++, pos += step / 8) {
    const major = i % 8 === 0
    const mid = i % 4 === 0
    ctx.beginPath()
    if (wide) {
      const top = major ? 2 : mid ? 8 : 12
      ctx.moveTo(pos + 0.5, height); ctx.lineTo(pos + 0.5, top)
      if (major && i) ctx.fillText(String(i / 8), pos + 2, 9)
    } else {
      const left = major ? 2 : mid ? 8 : 12
      ctx.moveTo(width, pos + 0.5); ctx.lineTo(left, pos + 0.5)
      if (major && i) ctx.save()
      if (major && i) { ctx.translate(2, pos - 1); ctx.rotate(-Math.PI / 2); ctx.fillText(String(i / 8), 0, 0); ctx.restore() }
    }
    ctx.stroke()
  }
}

export default RWord
