import { type OutlineNode } from '../engine';
import type { Doc } from '../engine';
export interface OutlinePanelProps {
    doc: Doc;
    onSelect?: (node: OutlineNode) => void;
    onClose?: () => void;
}
export declare function OutlinePanel({ doc, onSelect, onClose }: OutlinePanelProps): import("react").JSX.Element;
