import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react'
import { ErrorBoundary, downloadBlob, fetchOnline, saveOnline, type OfficeApi, type OfficeProps, type OpenInput, type SaveTarget, type Package } from '@voffice/react-shared'
import { addr, getCell, newWorkbook, setCell, type Workbook } from './engine'
import { readXlsx, writeXlsx } from './engine'
import '../../rshared/src/style.css'
export type { OfficeApi, OfficeProps, OpenInput, SaveTarget }
export interface RExcelProps extends OfficeProps { onReady?: () => void; onChange?: (dirty: boolean) => void; onSave?: (blob: Blob, name: string) => void; onError?: (error: Error) => void }
function workbookText(wb: Workbook): string { const sheet = wb.sheets[wb.active]; return [...sheet.cells.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([key, cell]) => `${key}\t${cell.f ? '=' + cell.f : String(cell.v ?? '')}`).join('\n') }
function workbookFromText(value: string): Workbook { const wb = newWorkbook(); const sheet = wb.sheets[0]; value.split(/\r?\n/).forEach((line, row) => line.split('\t').forEach((part, col) => { if (part.startsWith('=')) setCell(sheet, row, col, { f: part.slice(1) }); else if (part !== '') { const num = Number(part); setCell(sheet, row, col, Number.isFinite(num) && part.trim() !== '' ? { v: num, t: 'n' } : { v: part, t: 's' }) } })); return wb }
export const RExcel = forwardRef<OfficeApi, RExcelProps>(function RExcel({ src, fileName = 'Book1.xlsx', readonly, onReady, onChange, onSave, onError }, ref) {
  const [name, setName] = useState(fileName); const [dirty, setDirty] = useState(false); const [busy, setBusy] = useState(''); const [text, setText] = useState('')
  const wbRef = useRef<Workbook>(newWorkbook()); const pkgRef = useRef<Package | null>(null); const partsRef = useRef<string[]>([]); const sourceRef = useRef<OpenInput | null>(null)
  const mark = (value: boolean) => { setDirty(value); onChange?.(value) }
  const open = async (input: OpenInput, requestedName?: string): Promise<void> => { try { setBusy('Opening…'); const file = input instanceof Blob ? { blob: input, name: requestedName ?? (input instanceof File ? input.name : 'Book1.xlsx') } : await fetchOnline(typeof input === 'string' ? { kind: 'http', url: input } : input); const result = await readXlsx(new Uint8Array(await file.blob.arrayBuffer())); wbRef.current = result.wb; pkgRef.current = result.pkg; partsRef.current = result.sheetParts; sourceRef.current = input; setText(workbookText(result.wb)); setName(file.name); mark(false); onReady?.() } catch (error) { onError?.(error instanceof Error ? error : new Error(String(error))) } finally { setBusy('') } }
  const getBlob = async (): Promise<Blob> => writeXlsx(wbRef.current, { pkg: pkgRef.current ?? undefined, sheetParts: partsRef.current, format: 'xlsx' })
  const saveAs = async (requestedName = name): Promise<Blob> => { const blob = await getBlob(); const output = /\.[a-z0-9]+$/i.test(requestedName) ? requestedName : `${requestedName}.xlsx`; downloadBlob(blob, output); mark(false); return blob }
  const save = async (): Promise<void> => { const blob = await getBlob(); const source = sourceRef.current; if (source && typeof source !== 'string' && 'kind' in source && source.kind !== 'http') await saveOnline(source as SaveTarget, blob, name); else onSave?.(blob, name); mark(false) }
  const newDocument = () => { wbRef.current = newWorkbook(); pkgRef.current = null; partsRef.current = []; sourceRef.current = null; setText(''); setName('Book1.xlsx'); mark(false) }
  useImperativeHandle(ref, () => ({ open, newDocument, getBlob, save, saveAs, saveTo: async (target, requestedName) => saveOnline(target, await getBlob(), requestedName ?? name), get name() { return name }, get dirty() { return dirty } }), [name, dirty, text])
  useEffect(() => { if (src) void open(src) }, [])
  const updateText = (value: string) => { setText(value); wbRef.current = workbookFromText(value); mark(true) }
  return <ErrorBoundary><div className="vo-office"><div className="vo-title">Ledger — {name}{dirty ? ' *' : ''}</div><div className="vo-toolbar"><button onClick={() => void save()}>Save</button><button onClick={() => void saveAs()}>Save As</button><button onClick={newDocument}>New</button>{busy && <span>{busy}</span>}</div><main className="vo-editor"><textarea readOnly={readonly} value={text} onChange={(event) => updateText(event.target.value)} aria-label="Spreadsheet editor" /></main><div className="vo-status">Ready</div></div></ErrorBoundary>
})
export default RExcel
