import { type Doc, type Block } from './model';
export interface Caret {
    id: string;
    offset: number;
}
export declare function paginate(doc: Doc, contentW: number, contentH: number, measure: HTMLElement): Block[][];
/** resolve list markers (.lm spans) in a rendered container, numbering sequential paragraphs */
export declare function resolveMarkers(root: HTMLElement, doc: Doc): void;
/** capture the caret as (block id, character offset within that block) */
export declare function getCaret(root: HTMLElement): Caret | null;
/** restore a caret captured by getCaret after a re-render */
export declare function restoreCaret(root: HTMLElement, caret: Caret | null): void;
