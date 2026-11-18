import { useRef, useState } from 'react'
import { RWord, type RWordProps } from '../index'
import { MailMergePanel } from './MailMergePanel'
import { ReviewPanel } from './ReviewPanel'
import { TaskPane } from './TaskPane'
export type WorkspacePanel = 'none' | 'task' | 'mailmerge' | 'review'
export interface RWordWorkspaceProps extends RWordProps { initialPanel?: WorkspacePanel }
/** Composed React workspace for applications that need Quire's editor and side panels together. */
export function RWordWorkspace({ initialPanel = 'task', ...props }: RWordWorkspaceProps) { const ref = useRef<import('@voffice/react-shared').OfficeApi>(null); const [panel, setPanel] = useState<WorkspacePanel>(initialPanel); return <div className="vo-word-workspace"><div className="vo-word-main"><RWord ref={ref} {...props} /><div className="vo-workspace-actions"><button type="button" onClick={() => setPanel(panel === 'task' ? 'none' : 'task')}>Task Pane</button><button type="button" onClick={() => setPanel(panel === 'mailmerge' ? 'none' : 'mailmerge')}>Mail Merge</button><button type="button" onClick={() => setPanel(panel === 'review' ? 'none' : 'review')}>Review</button></div></div>{panel === 'task' && <TaskPane text="" onNew={() => ref.current?.newDocument()} onOpen={() => undefined} onClose={() => setPanel('none')} />}{panel === 'mailmerge' && <MailMergePanel onClose={() => setPanel('none')} />}{panel === 'review' && <ReviewPanel onClose={() => setPanel('none')} />}</div> }
