import type { Block, Doc } from '../engine/model';
export interface PageSlice {
    page: number;
    blocks: Block[];
    height: number;
}
export declare function paginateBlocks(doc: Doc, blockHeights: ReadonlyMap<string, number>, pageHeight?: number): PageSlice[];
export declare function pageCount(doc: Doc, heights: ReadonlyMap<string, number>, pageHeight?: number): number;
