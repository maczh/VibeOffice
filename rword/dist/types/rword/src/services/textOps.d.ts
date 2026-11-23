import type { Para, RPr } from '../engine/model';
export declare function replaceRange(text: string, start: number, end: number, replacement: string): string;
export declare function toggleCase(text: string, mode: 'upper' | 'lower' | 'title'): string;
export declare function splitParagraph(paragraph: Para, offset: number): [Para, Para];
export declare function applyRunStyle(text: string, style: RPr): {
    text: string;
    rPr: RPr;
};
