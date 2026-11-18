import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react'
import { ErrorBoundary, downloadBlob, fetchOnline, saveOnline, type OfficeApi, type OfficeProps, type OpenInput, type SaveTarget, type Package } from '@voffice/react-shared'
import { newPresentation, newTextShape, type Presentation } from './engine'
import { readPptx, writePptx } from './engine'
import '../../rshared/src/style.css'
export type { OfficeApi, OfficeProps, OpenInput, SaveTarget }
export interface RPptxProps extends OfficeProps { onReady?: () => void; onChange?: (dirty: boolean) => void; onSave?: (blob: Blob, name: string) => void; onError?: (error: Error) => void }
function presentationText(pres: Presentation): string { return pres.slides.map((slide, i) => `Slide ${i + 1}\n${slide.shapes.flatMap((shape) => shape.text ?? []).flatMap((para) => para.runs.map((run) => run.text)).join('')}`).join('\n\n') }
function presentationFromText(value: string): Presentation { const pres = newPresentation(); pres.slides[0].shapes = [newTextShape(914400, 914400, 10058400, 4572000, value)]; return pres }
export const RPptx = forwardRef<OfficeApi, RPptxProps>(function RPptx({ src, fileName = 'Presentation.pptx', readonly, onReady, onChange, onSave, onError }, ref) {
  const [name, setName] = useState(fileName); const [dirty, setDirty] = useState(false); const [busy, setBusy] = useState(''); const [text, setText] = useState('')
  const presRef = useRef<Presentation>(newPresentation()); const pkgRef = useRef<Package | null>(null); const partsRef = useRef<string[]>([]); const sourceRef = useRef<OpenInput | null>(null)
  const mark = (value: boolean) => { setDirty(value); onChange?.(value) }
  const open = async (input: OpenInput, requestedName?: string): Promise<void> => { try { setBusy('Opening…'); const file = input instanceof Blob ? { blob: input, name: requestedName ?? (input instanceof File ? input.name : 'Presentation.pptx') } : await fetchOnline(typeof input === 'string' ? { kind: 'http', url: input } : input); const result = await readPptx(new Uint8Array(await file.blob.arrayBuffer())); presRef.current = result.pres; pkgRef.current = result.pkg; partsRef.current = result.slideParts; sourceRef.current = input; setText(presentationText(result.pres)); setName(file.name); mark(false); onReady?.() } catch (error) { onError?.(error instanceof Error ? error : new Error(String(error))) } finally { setBusy('') } }
  const getBlob = async (): Promise<Blob> => writePptx(presRef.current, { pkg: pkgRef.current ?? undefined, slideParts: partsRef.current, format: presRef.current.format })
  const saveAs = async (requestedName = name): Promise<Blob> => { const blob = await getBlob(); const output = /\.[a-z0-9]+$/i.test(requestedName) ? requestedName : `${requestedName}.pptx`; downloadBlob(blob, output); mark(false); return blob }
  const save = async (): Promise<void> => { const blob = await getBlob(); const source = sourceRef.current; if (source && typeof source !== 'string' && 'kind' in source && source.kind !== 'http') await saveOnline(source as SaveTarget, blob, name); else onSave?.(blob, name); mark(false) }
  const newDocument = () => { presRef.current = newPresentation(); pkgRef.current = null; partsRef.current = []; sourceRef.current = null; setText(''); setName('Presentation.pptx'); mark(false) }
  useImperativeHandle(ref, () => ({ open, newDocument, getBlob, save, saveAs, saveTo: async (target, requestedName) => saveOnline(target, await getBlob(), requestedName ?? name), get name() { return name }, get dirty() { return dirty } }), [name, dirty, text])
  useEffect(() => { if (src) void open(src) }, [])
  const updateText = (value: string) => { setText(value); presRef.current = presentationFromText(value); mark(true) }
  return <ErrorBoundary><div className="vo-office"><div className="vo-title">Lectern — {name}{dirty ? ' *' : ''}</div><div className="vo-toolbar"><button onClick={() => void save()}>Save</button><button onClick={() => void saveAs()}>Save As</button><button onClick={newDocument}>New</button>{busy && <span>{busy}</span>}</div><main className="vo-editor"><textarea readOnly={readonly} value={text} onChange={(event) => updateText(event.target.value)} aria-label="Presentation editor" /></main><div className="vo-status">Ready</div></div></ErrorBoundary>
})
export default RPptx
