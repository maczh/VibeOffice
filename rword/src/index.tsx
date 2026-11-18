import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react'
import { ErrorBoundary, downloadBlob, fetchOnline, saveOnline, type OfficeApi, type OfficeProps, type OpenInput, type SaveTarget } from '@voffice/react-shared'
import { docText, emptyPara, newDoc, type Doc } from './engine'
import type { Package } from '@voffice/react-shared'
import { readDocx, writeDocx } from './engine'
import '../../rshared/src/style.css'

export type { OfficeApi, OfficeProps, OpenInput, SaveTarget }
export interface RWordProps extends OfficeProps { onReady?: () => void; onChange?: (dirty: boolean) => void; onSave?: (blob: Blob, name: string) => void; onError?: (error: Error) => void }
function docFromText(value: string, format = 'docx'): Doc { const doc = newDoc(); doc.format = format; doc.blocks = value.split(/\r?\n/).map((line) => ({ ...emptyPara(), runs: line ? [{ text: line }] : [] })); return doc }

export const RWord = forwardRef<OfficeApi, RWordProps>(function RWord({ src, fileName = 'Document.docx', readonly, onReady, onChange, onSave, onError }, ref) {
  const [name, setName] = useState(fileName); const [dirty, setDirty] = useState(false); const [busy, setBusy] = useState(''); const [text, setText] = useState('')
  const docRef = useRef<Doc>(newDoc()); const pkgRef = useRef<Package | null>(null); const partRef = useRef('word/document.xml'); const sourceRef = useRef<OpenInput | null>(null)
  const mark = (value: boolean) => { setDirty(value); onChange?.(value) }
  const open = async (input: OpenInput, requestedName?: string): Promise<void> => { try { setBusy('Opening…'); const file = input instanceof Blob ? { blob: input, name: requestedName ?? (input instanceof File ? input.name : 'Document.docx') } : await fetchOnline(typeof input === 'string' ? { kind: 'http', url: input } : input); const result = await readDocx(new Uint8Array(await file.blob.arrayBuffer())); docRef.current = result.doc; pkgRef.current = result.pkg; partRef.current = result.part; sourceRef.current = input; setText(docText(result.doc)); setName(file.name); mark(false); onReady?.() } catch (error) { onError?.(error instanceof Error ? error : new Error(String(error))) } finally { setBusy('') } }
  const getBlob = async (): Promise<Blob> => writeDocx(docRef.current, { pkg: pkgRef.current ?? undefined, part: partRef.current, format: docRef.current.format })
  const saveAs = async (requestedName = name): Promise<Blob> => { const blob = await getBlob(); const output = /\.[a-z0-9]+$/i.test(requestedName) ? requestedName : `${requestedName}.docx`; downloadBlob(blob, output); mark(false); return blob }
  const save = async (): Promise<void> => { const blob = await getBlob(); const source = sourceRef.current; if (source && typeof source !== 'string' && 'kind' in source && source.kind !== 'http') await saveOnline(source as SaveTarget, blob, name); else onSave?.(blob, name); mark(false) }
  const newDocument = () => { docRef.current = newDoc(); pkgRef.current = null; sourceRef.current = null; setText(''); setName('Document.docx'); mark(false) }
  useImperativeHandle(ref, () => ({ open, newDocument, getBlob, save, saveAs, saveTo: async (target, requestedName) => saveOnline(target, await getBlob(), requestedName ?? name), get name() { return name }, get dirty() { return dirty } }), [name, dirty, text])
  useEffect(() => { if (src) void open(src) }, [])
  const updateText = (value: string) => { setText(value); docRef.current = docFromText(value, docRef.current.format); mark(true) }
  return <ErrorBoundary><div className="vo-office"><div className="vo-title">Quire — {name}{dirty ? ' *' : ''}</div><div className="vo-toolbar"><button onClick={() => void save()}>Save</button><button onClick={() => void saveAs()}>Save As</button><button onClick={newDocument}>New</button>{busy && <span>{busy}</span>}</div><main className="vo-editor"><textarea readOnly={readonly} value={text} onChange={(event) => updateText(event.target.value)} aria-label="Document editor" /></main><div className="vo-status">{text.trim().split(/\s+/).filter(Boolean).length} words</div></div></ErrorBoundary>
})
export default RWord
