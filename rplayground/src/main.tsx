import { StrictMode, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { RWord } from '@voffice/rword'
import { RExcel } from '@voffice/rexcel'
import { RPptx } from '@voffice/rpptx'
import { FileDropZone, StartCenter, useLook, useRecentFiles, type OfficeApi, type SuiteApp } from '@voffice/react-shared'
import '@voffice/react-shared/styles'
import '../../rshared/src/style.css'

function App() {
  const [app, setApp] = useState<SuiteApp>('word'); const [start, setStart] = useState(false); const { look, setLook } = useLook(); const { recent, remember, remove } = useRecentFiles()
  const word = useRef<OfficeApi>(null); const excel = useRef<OfficeApi>(null); const pptx = useRef<OfficeApi>(null)
  const active = () => app === 'word' ? word.current : app === 'excel' ? excel.current : pptx.current
  const accept = app === 'word' ? '.docx,.docm,.dotx,.dotm' : app === 'excel' ? '.xlsx,.xlsm,.xltx,.xltm,.csv,.txt' : '.pptx,.pptm,.ppsx,.ppsm,.potx,.potm'
  const open = async (file: File) => { await active()?.open(file); remember({ name: file.name, app }) }
  const create = (next: SuiteApp) => { setApp(next); const target = next === 'word' ? word : next === 'excel' ? excel : pptx; target.current?.newDocument(); setStart(false) }
  return <div className="vo-react-host" data-look={look}><header className="vo-host-bar"><strong>VibeOffice React</strong>{(['word', 'excel', 'pptx'] as SuiteApp[]).map((id) => <button key={id} className={app === id ? 'active' : ''} onClick={() => setApp(id)}>{id === 'word' ? 'Quire' : id === 'excel' ? 'Ledger' : 'Lectern'}</button>)}<span /><button onClick={() => setStart(true)}>Start Center</button><select value={look} onChange={(e) => setLook(e.target.value as typeof look)}><option value="classic">Classic</option><option value="paper">Paper</option><option value="paper2016">Paper 2016</option></select></header><FileDropZone accept={accept} onFiles={(files) => { if (files[0]) void open(files[0]) }}><div className="vo-react-pane">{app === 'word' && <RWord ref={word} onSave={(blob, name) => { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click() }} />} {app === 'excel' && <RExcel ref={excel} />} {app === 'pptx' && <RPptx ref={pptx} />}</div></FileDropZone>{start && <div className="vo-react-start"><StartCenter recent={recent} onOpen={() => { setStart(false); document.body.querySelector<HTMLInputElement>('#react-host-file')?.click() }} onCreate={create} onRecent={(file) => { setApp(file.app); if (file.url) void active()?.open(file.url, file.name); setStart(false) }} onRemove={remove} /></div>}<input id="react-host-file" hidden type="file" accept={accept} onChange={(e) => { const file = e.target.files?.[0]; if (file) void open(file) }} /></div>
}
createRoot(document.getElementById('root')!).render(<StrictMode><App /></StrictMode>)
