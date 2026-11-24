import { useEffect, useState, type CSSProperties, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

export interface PptxSelectionOverlayNode {
  key: string;
  className: string;
  style?: CSSProperties;
  handle?: string;
  shapeId?: string;
  owned?: boolean;
  children?: PptxSelectionOverlayNode[];
}

interface SelectionOverlayRuntime {
  ed?: { reactSelectionOverlayCatalog?(): PptxSelectionOverlayNode[] };
  bus?: { on(event: string, listener: () => void): () => void };
}

function runtime(host: HTMLElement): SelectionOverlayRuntime | undefined {
  return (host.ownerDocument.defaultView as (Window & { L?: SelectionOverlayRuntime }) | null)?.L;
}

function renderNode(node: PptxSelectionOverlayNode): ReactNode {
  return <div
    key={node.key}
    className={node.className}
    style={node.style}
    data-h={node.handle}
    data-sid={node.shapeId}
    data-react-owned={node.owned ? 'pptx-selection-control' : undefined}
  >{node.children?.map(renderNode)}</div>;
}

/** React owns PPTX selection frames and their delegated pointer targets inside the engine overlay host. */
export function PptxSelectionOverlay({ host }: { host: HTMLElement }) {
  const [nodes, setNodes] = useState<PptxSelectionOverlayNode[]>(() => runtime(host)?.ed?.reactSelectionOverlayCatalog?.() || []);
  useEffect(() => {
    const refresh = () => setNodes(runtime(host)?.ed?.reactSelectionOverlayCatalog?.() || []);
    let attachedBus: SelectionOverlayRuntime['bus'];
    let offs: Array<() => void> = [];
    const attach = () => {
      const bus = runtime(host)?.bus;
      if (bus && bus !== attachedBus) {
        offs.forEach((off) => off());
        attachedBus = bus;
        offs = ['pptx-selection-overlay', 'selection'].map((event) => bus.on(event, refresh));
      }
      refresh();
    };
    host.addEventListener('vibeoffice-engine-runtime-ready', attach);
    attach();
    return () => {
      host.removeEventListener('vibeoffice-engine-runtime-ready', attach);
      offs.forEach((off) => off());
    };
  }, [host]);
  const target = host.querySelector<HTMLElement>('.ed-overlay');
  return target ? createPortal(<div className="pptx-react-selection-overlay" data-react-owned="pptx-selection-overlay" style={{ position: 'absolute', inset: 0, pointerEvents: 'none' }}>{nodes.map(renderNode)}</div>, target) : null;
}
