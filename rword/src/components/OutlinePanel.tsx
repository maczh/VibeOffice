import { useMemo } from 'react'
import { outline, type OutlineNode } from '../engine'
import type { Doc } from '../engine'
export interface OutlinePanelProps { doc: Doc; onSelect?: (node: OutlineNode) => void; onClose?: () => void }
export function OutlinePanel({ doc, onSelect, onClose }: OutlinePanelProps) { const tree = useMemo(() => outline(doc), [doc]); const draw = (nodes: OutlineNode[]) => <ul>{nodes.map(node => <li key={node.id}><button type="button" onClick={() => onSelect?.(node)}>{node.title}</button>{node.children.length ? draw(node.children) : null}</li>)}</ul>; return <aside className="vo-task-pane" aria-label="Document Map"><header><strong>Document Map</strong><button type="button" onClick={onClose}>×</button></header><div className="vo-task-body">{tree.length ? draw(tree) : <p>(No headings)</p>}</div></aside> }
