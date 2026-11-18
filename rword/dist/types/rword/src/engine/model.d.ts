export interface RPr {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    font?: string;
    /** point size */
    size?: number;
    color?: string;
    highlight?: string;
    vertAlign?: 'superscript' | 'subscript';
}
export interface ImageRef {
    rId: string;
    part: string;
    url: string;
    /** display size in px */
    w: number;
    h: number;
}
export interface Run {
    text?: string;
    br?: boolean;
    tab?: boolean;
    image?: ImageRef;
    hyperlink?: {
        rId: string;
        url: string;
    };
    rPr?: RPr;
    /** verbatim XML for content we do not model (fields, drawings, …) */
    raw?: string;
}
export interface PPr {
    align?: 'left' | 'center' | 'right' | 'justify';
    style?: string;
    numId?: number;
    ilvl?: number;
    /** left indent in px */
    indent?: number;
    hanging?: number;
    pageBreakBefore?: boolean;
    spaceBefore?: number;
    spaceAfter?: number;
    line?: number;
}
export interface Para {
    kind: 'p';
    id: string;
    runs: Run[];
    pPr?: PPr;
    raw?: string;
}
export interface Cell {
    paras: Para[];
    gridSpan?: number;
    vMerge?: 'restart' | 'continue';
    fill?: string;
    width?: number;
    raw?: string;
}
export interface Table {
    kind: 'tbl';
    id: string;
    rows: Cell[][];
    /** column widths in px */
    widths: number[];
    raw?: string;
}
export type Block = Para | Table;
export interface Section {
    /** twips (1/1440 inch) */
    pageW: number;
    pageH: number;
    margins: {
        top: number;
        right: number;
        bottom: number;
        left: number;
    };
}
export interface StyleDef {
    id: string;
    name: string;
    basedOn?: string;
    rPr?: RPr;
    pPr?: PPr;
}
export interface Doc {
    blocks: Block[];
    section: Section;
    styles: StyleDef[];
    defaultFont: string;
    defaultSize: number;
    /** numId → level-0 number format (bullet/decimal/…), from numbering.xml */
    numbering?: Record<number, string>;
    format: string;
}
export declare const uid: (p?: string) => string;
export declare const emptyPara: (pPr?: PPr) => Para;
export declare const defaultSection: () => Section;
export declare const newDoc: () => Doc;
export declare const isPara: (b: Block) => b is Para;
export declare const isTable: (b: Block) => b is Table;
/** plain text of a block (for word count / find) */
export declare function blockText(b: Block): string;
export declare function paraText(p: Para): string;
export declare function docText(d: Doc): string;
export declare function wordCount(d: Doc): number;
/** deep clone via JSON (model is plain data; image blob URLs survive as strings) */
export declare const cloneDoc: (d: Doc) => Doc;
