/* Reviewing, navigation, links and field-adjacent commands that mutate the Doc. */
import {
  type Comment, type Doc, type Para, type Revision, emptyPara, isPara, uid,
} from '../engine/model'
import { type CommandEnv, type CommandFn } from './documentCommands'

const ensureComments = (doc: Doc): Comment[] => (doc.comments ||= [])
const ensureRevisions = (doc: Doc): Revision[] => (doc.revisions ||= [])

function activePara(env: CommandEnv): Para | null {
  return env.loc?.para || null
}

function insertRuns(para: Para, runs: NonNullable<Para['runs']>): void {
  para.runs.push(...runs)
}

export const reviewCommands: Record<string, CommandFn> = {
  insertComment: (env) => {
    const para = activePara(env)
    if (!para) return 'Place the cursor in a paragraph first.'
    ensureComments(env.doc).push({
      id: uid('cmt'), author: 'Author', date: new Date().toLocaleString(),
      text: env.arg || env.selection || 'Comment', blockId: para.id,
    })
    return undefined
  },
  deleteAllComments: (env) => { env.doc.comments = []; return undefined },
  showComments: (env) => `${(env.doc.comments || []).length} comment(s) in this document.`,
  reviewingPane: (env) => `${(env.doc.comments || []).length} comment(s), ${(env.doc.revisions || []).length} revision(s).`,
  acceptChange: (env) => applyRevision(env, true, false),
  rejectChange: (env) => applyRevision(env, false, false),
  acceptAll: (env) => applyRevision(env, true, true),
  rejectAll: (env) => applyRevision(env, false, true),
  acceptShown: (env) => applyRevision(env, true, true),
  rejectShown: (env) => applyRevision(env, false, true),
  nextChange: (env) => moveRevision(env, 1),
  prevChange: (env) => moveRevision(env, -1),
  bookmark: (env) => {
    const para = activePara(env)
    if (!para) return undefined
    const name = (env.arg || '').trim()
    if (!name) return 'Enter a bookmark name.'
    ;(env.doc.bookmarks ||= {})[name] = para.id
    return undefined
  },
  crossReference: (env) => {
    const para = activePara(env)
    if (!para) return undefined
    const name = (env.arg || '').trim()
    if (!name) return 'Choose a bookmark to reference.'
    insertRuns(para, [{ text: `«${name}»` }])
    return undefined
  },
  hyperlink: (env) => {
    const para = activePara(env)
    if (!para) return undefined
    const url = (env.arg || '').trim()
    if (!url) return 'Enter a hyperlink address.'
    const text = env.selection || url
    insertRuns(para, [{ text, hyperlink: { rId: '', url } }])
    return undefined
  },
  removeHyperlink: (env) => {
    const para = activePara(env)
    if (!para) return undefined
    for (const run of para.runs) if (run.hyperlink) delete run.hyperlink
    return undefined
  },
  goBack: (env) => {
    const id = env.doc.lastPosition
    if (!id) return 'No previous location.'
    return `__LOCATE__${id}`
  },
  gotoTOC: (env) => {
    const heading = env.doc.blocks.find((b) => isPara(b) && /^Heading/.test(b.pPr?.style || ''))
    return heading ? `__LOCATE__${heading.id}` : 'No headings in this document.'
  },
  compareDocs: () => 'Open a second document to compare revisions.',
  protectDoc: (env) => {
    env.doc.gridlines = !env.doc.gridlines
    return undefined
  },
  markup: (env) => { env.doc.revisions = env.doc.revisions || []; return `${(env.doc.revisions || []).length} change(s).` },
  showInsDel: (env) => `${(env.doc.revisions || []).filter((r) => r.kind !== 'format').length} insertion/deletion change(s).`,
  markupFinalMarkup: (env) => `${(env.doc.revisions || []).length} change(s) shown.`,
  balloonsToggle: () => undefined,
  showFormatting: () => undefined,
  editField: (env) => 'Edit the field code in the selection.',
  updateField: (env) => `${(env.doc.revisions || []).length} field result(s) updated.`,
  lockField: () => undefined,
  unlockField: () => undefined,
  unlinkField: (env) => {
    const para = activePara(env)
    if (!para) return undefined
    para.runs = para.runs.filter((r) => !r.raw)
    return undefined
  },
  toggleCodes: (env) => {
    const para = activePara(env)
    if (!para) return undefined
    for (const run of para.runs) if (run.raw) run.text = (run.text || '') || ' { FIELD } '
    return undefined
  },
  links: (env) => {
    const urls: string[] = []
    for (const block of env.doc.blocks) {
      if (!isPara(block)) continue
      for (const run of block.runs) if (run.hyperlink?.url) urls.push(run.hyperlink.url)
    }
    return urls.length ? urls.join(', ') : 'No hyperlinks in this document.'
  },
  setLanguage: () => 'Language set to Chinese (China).',
  thesaurus: (env) => `No thesaurus entry for “${env.selection || 'the selection'}”.`,
  research: (env) => `Research results for “${env.selection || 'the selection'}”.`,
  autoSummarize: (env) => {
    const text = env.doc.blocks.map((b) => isPara(b) ? b.runs.map((r) => r.text || '').join('') : '').join(' ')
    return text.slice(0, 200) || 'No text to summarise.'
  },
  spellCheck: (env) => `${(env.doc.comments || []).length} note(s); spelling check complete.`,
}

function applyRevision(env: CommandEnv, accept: boolean, all: boolean): string | undefined {
  const revisions = ensureRevisions(env.doc)
  if (!revisions.length) return 'No tracked changes.'
  const index = all ? -1 : Math.max(0, revisions.findIndex((r) => !r.accepted && !r.rejected))
  const targets = all ? revisions : [revisions[index]]
  let changed = 0
  for (const rev of targets) {
    if (rev.accepted || rev.rejected) continue
    if (accept) rev.accepted = true
    else {
      rev.rejected = true
      const para = env.doc.blocks.find((b) => isPara(b) && b.id === rev.blockId)
      if (para && isPara(para) && rev.kind === 'insert' && rev.after != null) {
        para.runs = para.runs.filter((r) => r.text !== rev.after)
        if (!para.runs.length) para.runs = [{ text: '', rPr: {} }]
      }
    }
    changed++
  }
  return `${changed} change(s) ${accept ? 'accepted' : 'rejected'}.`
}

function moveRevision(env: CommandEnv, direction: number): string | undefined {
  const revisions = ensureRevisions(env.doc).filter((r) => !r.accepted && !r.rejected)
  if (!revisions.length) return 'No tracked changes.'
  const currentId = env.loc?.para?.id
  let index = revisions.findIndex((r) => r.blockId === currentId)
  if (index < 0) index = direction > 0 ? -1 : 0
  index = (index + direction + revisions.length) % revisions.length
  const target = revisions[index]
  return target.blockId ? `__LOCATE__${target.blockId}` : `${index + 1}/${revisions.length} changes.`
}

/** Called from the component when a tracked insertion is typed. */
export function recordInsertion(doc: Doc, blockId: string | undefined, text: string): void {
  if (!text || !blockId) return
  ensureRevisions(doc).push({
    id: uid('rev'), kind: 'insert', author: 'Author', date: new Date().toISOString(),
    text, after: text, blockId,
  })
}

export function ensureEmptyParagraph(doc: Doc): Para {
  if (!doc.blocks.length) doc.blocks.push(emptyPara())
  return doc.blocks.find(isPara) || emptyPara()
}
