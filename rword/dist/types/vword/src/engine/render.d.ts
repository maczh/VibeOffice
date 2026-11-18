import { type Doc, type Block, type Para } from './model';
export declare function renderBlock(b: Block, doc: Doc): string;
export declare function syncPara(el: HTMLElement, prev?: Para): Para;
export declare function syncBlock(el: HTMLElement, doc: Doc, prev?: Block): Block;
