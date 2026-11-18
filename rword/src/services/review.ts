import type { Doc, Para, Run } from '../engine/model'
export interface Revision { id: string; kind: 'insert' | 'delete' | 'format'; author: string; date: string; text: string; accepted?: boolean }
export interface Comment { id: string; author: string; date: string; text: string; resolved?: boolean }
export function addRevision(run: Run, revision: Revision): Run { return { ...run, raw: run.raw, rPr: { ...run.rPr }, text: revision.kind === 'delete' ? undefined : run.text } }
export function acceptRevision(revision: Revision): Revision { return { ...revision, accepted: true } }
export function rejectRevision(revision: Revision): Revision { return { ...revision, accepted: false } }
export function documentComments(doc: Doc): Comment[] { const comments: Comment[] = []; for (const block of doc.blocks) if (block.kind === 'p') for (const run of block.runs) if (run.raw?.includes('comment')) comments.push({ id: run.raw, author: '', date: '', text: run.text ?? '' }); return comments }
export function plainParagraph(text: string): Para { return { kind: 'p', id: `review-${Date.now()}`, runs: text ? [{ text }] : [] } }
