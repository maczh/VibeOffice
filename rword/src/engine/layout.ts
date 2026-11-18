/* Pagination and caret helpers. Blocks are measured at the content width and distributed
 * into pages of the content height; paragraphs and tables stay whole on a page. */
import { type Doc, type Block } from './model'
import { renderBlock } from './render'

export interface Caret { id: string; offset: number }

export function paginate(doc: Doc, contentW: number, contentH: number, measure: HTMLElement): Block[][] {
  measure.style.width = contentW + 'px'
  measure.innerHTML = ''
  const pages: Block[][] = [[]]
  let used = 0
  for (const b of doc.blocks) {
    const wrap = document.createElement('div')
    wrap.innerHTML = renderBlock(b, doc)
    const el = wrap.firstElementChild as HTMLElement
    measure.appendChild(el)
    const h = el.offsetHeight
    measure.removeChild(el)
    const forceBreak = b.kind === 'p' && b.pPr?.pageBreakBefore
    if ((used + h > contentH || forceBreak) && pages[pages.length - 1].length) {
      pages.push([b]); used = h
    } else {
      pages[pages.length - 1].push(b); used += h
    }
  }
  return pages
}

/** resolve list markers (.lm spans) in a rendered container, numbering sequential paragraphs */
export function resolveMarkers(root: HTMLElement, doc: Doc): void {
  const counters = new Map<string, number>()
  let lastKey = ''
  for (const lm of Array.from(root.querySelectorAll<HTMLElement>('.lm'))) {
    const numId = lm.getAttribute('data-num') || '0'
    const ilvl = lm.getAttribute('data-ilvl') || '0'
    const fmt = doc.numbering?.[+numId] || 'bullet'
    const key = numId + ':' + ilvl
    if (key !== lastKey) { /* a different list restarts only if numId differs */ }
    if (fmt === 'bullet') {
      lm.textContent = ['•', '◦', '▪'][+ilvl % 3] + ' '
    } else {
      const n = (counters.get(key) || 0) + 1
      counters.set(key, n)
      lm.textContent = (fmt === 'lowerLetter' ? String.fromCharCode(96 + n) : fmt === 'upperLetter' ? String.fromCharCode(64 + n) : fmt === 'lowerRoman' ? roman(n) : String(n)) + '. '
    }
    lastKey = key
  }
}

function roman(n: number): string {
  const table: [number, string][] = [[10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]
  let out = ''
  for (const [v, s] of table) while (n >= v) { out += s; n -= v }
  return out
}

/** capture the caret as (block id, character offset within that block) */
export function getCaret(root: HTMLElement): Caret | null {
  const sel = window.getSelection()
  if (!sel || !sel.rangeCount) return null
  const range = sel.getRangeAt(0)
  let node: Node | null = range.startContainer
  let block: HTMLElement | null = null
  while (node && node !== root) {
    if (node.nodeType === 1 && (node as HTMLElement).hasAttribute?.('data-id')) { block = node as HTMLElement; break }
    node = node.parentNode
  }
  if (!block) return null
  const pre = range.cloneRange()
  pre.selectNodeContents(block)
  pre.setEnd(range.startContainer, range.startOffset)
  const offset = pre.toString().length
  return { id: block.getAttribute('data-id')!, offset }
}

/** restore a caret captured by getCaret after a re-render */
export function restoreCaret(root: HTMLElement, caret: Caret | null): void {
  if (!caret) return
  const block = root.querySelector<HTMLElement>(`[data-id="${caret.id}"]`)
  if (!block) return
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT)
  let remaining = caret.offset
  let node: Node | null = walker.nextNode()
  let last: Node | null = null
  while (node) {
    const len = (node.nodeValue || '').length
    if (remaining <= len) {
      const sel = window.getSelection()!
      const r = document.createRange()
      r.setStart(node, remaining)
      r.collapse(true)
      sel.removeAllRanges()
      sel.addRange(r)
      return
    }
    remaining -= len
    last = node
    node = walker.nextNode()
  }
  const sel = window.getSelection()!
  const r = document.createRange()
  if (last) r.setStart(last, (last.nodeValue || '').length)
  else r.selectNodeContents(block)
  r.collapse(false)
  sel.removeAllRanges()
  sel.addRange(r)
}
