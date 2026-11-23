import type { Block, Doc, Para } from '../engine/model';
export interface OutlineNode {
    id: string;
    title: string;
    level: number;
    block: Para;
    children: OutlineNode[];
}
export declare function outline(doc: Doc): OutlineNode[];
export declare function headingLevel(block: Block): number;
